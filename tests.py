"""スモークテスト（画面と API をひととおり叩く）

    python tests.py              全部を1本ずつ別プロセスで実行
    python tests.py smoke_board  1本だけ実行

各テストは一時ファイルの SQLite を使い、instance/ の本番データには触れない。
"""
import json
import os
import subprocess
import sys
import tempfile
import urllib.error

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import ldap_auth  # noqa: E402
from app import conf, create_app  # noqa: E402
from app import Employee, FlowMap, MonthlyResponse, MonthlySurvey, Respondent, StageLog, Task, _migrate_monthly_questions, answerable_periods, current_period, db, lookup, period_label, public_period, shift_period  # noqa: E402

for _k, _v in (("ANTHROPIC_API_KEY", ""), ("LDAP_API_URL", ""), ("LDAP_SERVER", ""), ("ADMIN_ID", "admin"), ("ADMIN_PASSWORD", "adminpass")):
    conf.override(_k, _v)  # テストは config.py の値に依存しない


def smoke_bizmap_edit():
    """業務マップ編集（グラフAPI・マージ）のスモークテスト"""




    tmpdir = tempfile.mkdtemp()
    dbfile = os.path.join(tmpdir, "test.sqlite3").replace("\\", "/")
    app = create_app({"SQLALCHEMY_DATABASE_URI": "sqlite:///" + dbfile, "TESTING": True})
    client = app.test_client()
    client.post("/login", data={"emp_id": "admin", "password": "adminpass"})
    fails = []


    def check(cond, msg):
        print(("OK  " if cond else "NG  ") + msg)
        if not cond:
            fails.append(msg)


    client.post("/admin/demo/load")

    # 1. 自動生成グラフ
    r = client.get("/bizmap/api/graph?kind=info")
    g = r.get_json()
    check(r.status_code == 200 and g["needs_layout"] is True, "GET graph (auto) needs_layout")
    check(len(g["nodes"]) > 20 and len(g["edges"]) > 20, f"auto graph nodes={len(g['nodes'])} edges={len(g['edges'])}")
    task_node = next(n for n in g["nodes"] if n["id"].startswith("t"))
    some_edge = g["edges"][0]
    check(task_node.get("url") and task_node["cls"] in ("task", "done"), "task node has url/cls")

    # 2. 手修正して保存: 位置・名前・色、ノード非表示、エッジ非表示、手動ノード・手動エッジ、反転、ラベル
    data = {"nodes": [], "edges": [], "suppressed_nodes": [], "suppressed_edges": [], "view": {"x": 10, "y": 20, "zoom": 0.8}}
    hidden_node = next(n for n in g["nodes"] if n["id"].startswith("s_"))
    for i, n in enumerate(g["nodes"]):
        if n["id"] == hidden_node["id"]:
            continue
        m = dict(n)
        m["x"], m["y"] = 100 + i * 10, 50 + i * 5
        if n["id"] == task_node["id"]:
            m["text"] = "手で直した業務名"
            m["text_edited"] = True
            m["color"] = "yellow"
        data["nodes"].append(m)
    data["suppressed_nodes"].append(hidden_node["id"])
    manual_node = {"id": "m_abc123", "text": "手で足した箱", "cls": "place", "source": "manual", "x": 900, "y": 300, "color": None, "text_edited": False}
    data["nodes"].append(manual_node)
    for e in g["edges"]:
        if e["id"] == some_edge["id"]:
            data["suppressed_edges"].append(e["id"])
            continue
        m = dict(e)
        if e["from"] == task_node["id"] or e["to"] == task_node["id"]:
            m["reversed"] = True
            m["label"] = "ラベル修正"
            m["label_edited"] = True
        data["edges"].append(m)
    data["edges"].append({"id": f"{task_node['id']}->m_abc123", "from": task_node["id"], "to": "m_abc123", "label": "手動リンク", "kind": "manual", "source": "manual", "reversed": False, "label_edited": True})
    r = client.post("/bizmap/api/graph", json={"kind": "info", "data": data})
    check(r.status_code == 200 and r.get_json()["ok"], "POST save")
    with app.app_context():
        check(FlowMap.query.count() == 1, "FlowMap row created")

    # 3. マージ結果
    r = client.get("/bizmap/api/graph?kind=info")
    m = r.get_json()
    ids = {n["id"]: n for n in m["nodes"]}
    check(m["needs_layout"] is False and m["view"]["zoom"] == 0.8, "merged: needs_layout False, view kept")
    check(hidden_node["id"] not in ids, "merged: suppressed node hidden")
    check(ids[task_node["id"]]["text"] == "手で直した業務名" and ids[task_node["id"]]["color"] == "yellow", "merged: text/color override kept")
    check(ids[task_node["id"]]["x"] is not None, "merged: positions kept")
    check("m_abc123" in ids and ids["m_abc123"]["source"] == "manual", "merged: manual node kept")
    eids = {e["id"]: e for e in m["edges"]}
    check(some_edge["id"] not in eids, "merged: suppressed edge hidden")
    check(f"{task_node['id']}->m_abc123" in eids, "merged: manual edge kept")
    rev = [e for e in m["edges"] if e.get("reversed")]
    check(rev and all(e["label"] == "ラベル修正" for e in rev), f"merged: reversed+label kept ({len(rev)})")
    check(m["meta"]["saved_at"] is not None, "merged: saved_at")

    # 4. 新しい業務を登録すると new_node_ids に出る
    client.post("/survey/api/respondent", json={"emp_id": "E001", "name": "", "workplace_id": "", "role": "現場リーダー", "pc_hours": "1〜2時間"})
    client.post("/survey/api/tasks", json={
        "name": "新規に登録した業務", "domain_id": "2", "purpose": "テスト", "output": "新しい帳票",
        "output_users": ["上司"], "inputs": ["メール"], "frequency": "毎週", "duration": "15分未満", "headcount": "1",
        "tools": ["Excel"], "nature": ["資料作成"], "pains": ["特になし"], "wish": "減らしたい", "self_improve": "分からない",
    })
    with app.app_context():
        new_task = Task.query.filter_by(name="新規に登録した業務").first()
    r = client.get("/bizmap/api/graph?kind=info")
    m2 = r.get_json()
    check(f"t{new_task.id}" in m2["new_node_ids"], "merged: new task appears in new_node_ids")
    check(f"t{new_task.id}" in {n["id"] for n in m2["nodes"]}, "merged: new task node included")
    check("m_abc123" in {n["id"] for n in m2["nodes"]}, "merged: manual node still kept after new task")

    # 5. 閲覧ページに反映（Mermaid に手修正が出る）
    r = client.get("/bizmap/")
    html = r.data.decode("utf-8")
    check(r.status_code == 200 and "手で直した業務名" in html and "手で足した箱" in html, "index renders merged Mermaid")
    check("手修正あり" in html, "index shows saved badge")
    r = client.get("/bizmap/edit?kind=info")
    check(r.status_code == 200 and "fe-app" in r.data.decode("utf-8"), "edit page renders")
    for kind in ("people", "goods"):
        r = client.get(f"/bizmap/api/graph?kind={kind}&workplace_id=1")
        check(r.status_code == 404, f"以前の図 kind={kind} はもうない")
    r = client.get("/bizmap/edit?kind=info&workplace_id=1&domain_id=2")
    check(r.status_code == 200, "edit page with filters")
    r = client.get("/bizmap/")
    html = r.data.decode("utf-8")
    check("tab-people" not in html and "tab-goods" not in html and "nav-tabs" not in html and "システム・情報源と業務のつながり" in html and "職場・部署の間の受け渡し" in html and "領域別の把握状況" in html and "同じアウトプットを複数人が作成" in html and "classDef dupout" in html, "業務マップはタブなしの1ページ（図・表・把握状況・重複）")
    r = client.get("/bizmap/domain/2")
    check(r.status_code == 200, "domain page renders")

    # 5b. 同じアウトプットを複数人が作っていると、図のアウトプットの箱が「統合候補」になる
    with app.app_context():
        rs = Respondent.query.order_by(Respondent.id).limit(2).all()
        for i, r_ in enumerate(rs):
            db.session.add(Task(name=f"重複テスト業務{i}", respondent_id=r_.id, workplace_id=r_.workplace_id, output="重複テストの集計表", purpose="テスト", frequency="毎週", duration="15〜30分", headcount=1))
        db.session.commit()
        dup_ids = [t.id for t in Task.query.filter(Task.name.like("重複テスト業務%")).all()]
    g4 = client.get("/bizmap/api/graph?kind=info").get_json()
    cls_by_id = {n["id"]: n["cls"] for n in g4["nodes"]}
    check(len(dup_ids) == 2 and all(cls_by_id.get(f"o{i}") == "dupout" for i in dup_ids), "同じアウトプットの箱は dupout（統合候補）")
    r = client.get("/bizmap/")
    check("dupout" in r.data.decode("utf-8"), "閲覧ページの図にも統合候補の箱が出る")

    # 6. リセット
    r = client.delete("/bizmap/api/graph?kind=info")
    check(r.status_code == 200, "DELETE reset")
    m3 = client.get("/bizmap/api/graph?kind=info").get_json()
    check(m3["needs_layout"] is True and hidden_node["id"] in {n["id"] for n in m3["nodes"]}, "after reset: auto graph again")

    # 7. 不正データ
    r = client.post("/bizmap/api/graph", json={"kind": "info", "data": "oops"})
    check(r.status_code == 400, "POST invalid data -> 400")
    r = client.get("/bizmap/api/graph?kind=nope")
    check(r.status_code == 404, "GET unknown kind -> 404")

    return fails


def smoke_board():
    """活動状況（7段階）のスモークテスト"""




    tmpdir = tempfile.mkdtemp()
    dbfile = os.path.join(tmpdir, "test.sqlite3").replace("\\", "/")
    app = create_app({"SQLALCHEMY_DATABASE_URI": "sqlite:///" + dbfile, "TESTING": True})
    client = app.test_client()
    client.post("/login", data={"emp_id": "admin", "password": "adminpass"})
    fails = []


    def check(cond, msg):
        print(("OK  " if cond else "NG  ") + msg)
        if not cond:
            fails.append(msg)


    def page(url):
        r = client.get(url)
        return r.status_code, r.data.decode("utf-8")


    # 空の状態でも掲示板が出る
    code, html = page("/board/")
    check(code == 200 and "進行中の案件はありません" in html, "空の掲示板")

    client.post("/admin/demo/load")

    # 掲示板（通常）
    code, html = page("/board/")
    check(code == 200, "掲示板 200")
    for name in ("開始", "要求定義", "要件定義", "テスト", "確認", "完了！"):
        check(name in html, f"段階名 {name} が出る")
    check("説明" not in html and "登録してください" not in html, "説明文がない")
    check("検査成績書の作成" in html and "テスト中" in html, "進行中の案件とひとことが出る")
    check("保留中" in html and "組立実績の日次入力" in html, "保留の案件が出る")
    check("週次生産計画の作成と配布" in html and "伴走" in html, "自主改善の案件は「伴走」と表示")
    check("不良品発生報告のメール配信" in html and "手作業のメール作成をなくしました" in html, "完了案件とひとことが出る")
    check("最近の動き" not in html and "pb-funnel" not in html, "最近の動き・集計帯がない")
    check("田中 一郎" in html, "依頼者名が出る（既定）")
    check("大画面" not in html and "pb-chip" not in html and "pb-scope-select" in html and ">製造1課</option>" in html and "<footer" not in html, "大画面リンクなし・職場は見出しのプルダウン・フッターなし")
    check("pb-strip" in html and "累計改善案件" in html and "年間削減工数" in html and "削減率" in html and "自主改善者" not in html and "自動化・廃止業務" not in html and "先月より減った" not in html, "活動状況: 指標は累計改善案件と年間削減工数の2枚")
    check(html.count("pb-case-name") == 2 and "不良品発生報告のメール配信" in html and "pb-case-num" in html, "活動状況: 直近の完了2件が同じ行に並ぶ")
    with app.app_context():
        wp_tasks = Task.query.filter_by(workplace_id=2).all()
        wp_total = len(wp_tasks)
        wp_done = len([t for t in wp_tasks if t.status == "完了"])
    code, html = page("/board/?workplace_id=2")
    check(code == 200 and f'<div class="stat-value">{wp_total}<span class="stat-unit">件</span></div>' in html and f"完了 {wp_done} " in html, f"職場を選ぶとカードの数字もその職場の値（{wp_total} 件）")
    check("依頼してみませんか" in html and 'href="/survey/"' in html and 'href="/monthly/"' in html, "掲示板: 自動化依頼と月次アンケートへの導線")

    # 職場で絞り込み
    with app.app_context():
        t_test = Task.query.filter_by(name="検査成績書の作成").first()
        wp_id = t_test.workplace_id
        other = Task.query.filter(Task.workplace_id != wp_id, Task.status == "対応中").first()
    code, html = page(f"/board/?workplace_id={wp_id}")
    check(code == 200 and "検査成績書の作成" in html and (other is None or other.name not in html), "職場で絞り込める")

    # 掲示モードは廃止（パラメータを付けても通常表示）
    code, html = page("/board/?kiosk=1")
    check(code == 200 and "pb-kiosk" not in html and "navbar" in html, "掲示モードは出ない")
    code, html = page("/")
    check("<footer" not in html, "ダッシュボードにもフッターなし")

    # 氏名を隠す設定
    r = client.post("/admin/settings", data={"project_name": "P"})
    code, html = page("/board/")
    check("pb-name" not in html and "伊藤 大輔" not in html, "設定: 依頼者名が消える")
    client.post("/admin/settings", data={"project_name": "P", "board_show_names": "1"})
    code, html = page("/board/")
    check("pb-name" in html and "伊藤 大輔" in html, "設定: 氏名を再表示")

    # ダッシュボード・案件一覧・詳細
    code, html = page("/")
    check(code == 200 and "活動状況で案件の進み具合を見る" in html and "pb-funnel" not in html and "職場別の登録件数と月間時間" not in html and "業務領域別の把握状況" not in html, "ダッシュボード: 重複を外してリンクに")
    check("職場別の回答状況" in html and "最近登録された業務" in html and "次のアクション提案" in html, "ダッシュボード: 残した区画")
    code, html = page("/survey/")
    check(code == 200 and "自動化依頼" in html and 'id="rq"' in html, "自動化依頼は1画面")
    code, html = page("/cases/")
    check("標準手順で検討" not in html, "案件一覧: 標準手順の説明文を削除")
    code, html = page("/cases/")
    check(code == 200 and "進捗（活動状況）" in html and "pb-steps-sm" in html, "案件一覧に進捗列")
    code, html = page("/cases/?stage=test")
    check(code == 200 and "検査成績書の作成" in html and "受注情報のExcel転記" not in html, "案件一覧を段階で絞り込める")

    with app.app_context():
        t = Task.query.filter_by(name="原材料の払出し依頼票の作成").first()
        tid = t.id
        check(t.stage == "request" and t.status == "対応中", f"デモ: 要求定義/対応中 ({t.stage}/{t.status})")
        check(len(t.stage_logs) == 1 and t.stage_logs[0].stage == "request", "デモ: 履歴が1件")
        done = Task.query.filter_by(name="出勤簿の集計").first()
        check(done.stage == "done" and done.completed_at is not None and len(done.stage_logs) == 6, "デモ: 完了案件は履歴6件・完了日あり")
        check(done.stage_dates["done"] == done.completed_at, "デモ: 完了日と履歴が一致")

    code, html = page(f"/cases/{tid}")
    check(code == 200 and "次の段階へ ▶" in html and 'name="stage"' not in html and 'name="status"' not in html and "保留にする" in html and 'name="progress_note"' in html, "案件詳細: ステータス・段階の選択なし、保留ボタンあり")

    # 次の段階へ（request -> spec）
    r = client.post(f"/cases/{tid}/stage", data={"direction": "next"})
    check(r.status_code == 302, "次へ 302")
    with app.app_context():
        t = db.session.get(Task, tid)
        check(t.stage == "spec" and t.status == "対応中", f"次へ: spec/対応中 ({t.stage}/{t.status})")
        check(len(t.stage_logs) == 2 and t.stage_logs[-1].stage == "spec", "次へ: 履歴が増える")

    # 戻す（spec -> request）
    client.post(f"/cases/{tid}/stage", data={"direction": "prev"})
    with app.app_context():
        t = db.session.get(Task, tid)
        check(t.stage == "request" and len(t.stage_logs) == 3, "戻す: request・履歴3件")

    # 直接指定 done -> ステータス完了・完了日
    client.post(f"/cases/{tid}/stage", data={"stage": "done", "progress_note": "運用開始しました"})
    with app.app_context():
        t = db.session.get(Task, tid)
        check(t.stage == "done" and t.status == "完了" and t.completed_at is not None and t.progress_note == "運用開始しました", "直接指定: 完了に同期")

    # フォームはステータス・段階を変えない（ひとことだけ）
    form = {"decision": "未定", "assignment": "未定", "assignee": "", "after_minutes": "", "result_note": "", "progress_note": "活動状況のひとこと"}
    client.post(f"/cases/{tid}/update", data=form)
    with app.app_context():
        t = db.session.get(Task, tid)
        check(t.stage == "done" and t.status == "完了" and t.progress_note == "活動状況のひとこと", "フォーム更新では段階・ステータスは変わらない")

    # 完了した案件は保留にできない
    client.post(f"/cases/{tid}/hold")
    with app.app_context():
        t = db.session.get(Task, tid)
        check(t.status == "完了", "完了は保留にできない")

    # 段階を test に戻し、保留にする -> 段階は保ったまま保留
    client.post(f"/cases/{tid}/stage", data={"stage": "test"})
    client.post(f"/cases/{tid}/hold")
    with app.app_context():
        t = db.session.get(Task, tid)
        check(t.stage == "test" and t.status == "保留" and t.completed_at is None, f"保留で段階を保つ ({t.stage}/{t.status})")
    code, html = page("/board/")
    check("保留中" in html, "掲示板に保留中セクション")
    code, html = page(f"/cases/{tid}")
    check("保留を解除" in html, "案件詳細に保留解除ボタン")

    # 保留を解除 -> 対応中
    client.post(f"/cases/{tid}/hold")
    with app.app_context():
        t = db.session.get(Task, tid)
        check(t.stage == "test" and t.status == "対応中", f"保留解除で対応中 ({t.stage}/{t.status})")

    # 保留中に次へ -> 再開して進む
    client.post(f"/cases/{tid}/hold")
    client.post(f"/cases/{tid}/stage", data={"direction": "next"})
    with app.app_context():
        t = db.session.get(Task, tid)
        check(t.stage == "check" and t.status == "対応中", f"保留から次へで再開 ({t.stage}/{t.status})")

    # 開始に戻すと受付
    client.post(f"/cases/{tid}/stage", data={"stage": "start"})
    with app.app_context():
        t = db.session.get(Task, tid)
        check(t.stage == "start" and t.status == "受付", f"開始に戻すと受付 ({t.stage}/{t.status})")
    client.post(f"/cases/{tid}/stage", data={"stage": "check"})

    # 受付のまま（同じ段階を指定しても履歴は増えない）
    with app.app_context():
        n_before = StageLog.query.filter_by(task_id=tid).count()
    client.post(f"/cases/{tid}/stage", data={"stage": "check"})
    with app.app_context():
        check(StageLog.query.filter_by(task_id=tid).count() == n_before, "同じ段階なら履歴は増えない")

    # 新規登録した業務は開始
    with app.app_context():
        new_task = Task.query.filter_by(name="日報の実績をExcelの集計表に転記する").first()
        check(new_task.stage_dates["start"] == new_task.created_at, "開始の日付は登録日")

    # CSV
    r = client.get("/cases/export.csv")
    csv = r.data.decode("utf-8-sig")
    check("進捗段階" in csv and "段階更新日" in csv and "テスト" in csv, "CSVに進捗段階")

    # デモ削除で履歴も消える
    with app.app_context():
        before = StageLog.query.count()
    client.post("/admin/demo/clear")
    with app.app_context():
        check(before > 0 and StageLog.query.count() == 0 and Task.query.count() == 0, f"デモ削除で履歴も消える（{before} -> 0）")
    code, html = page("/board/")
    check(code == 200, "削除後も掲示板 200")

    return fails


def smoke_errors():
    """存在しないURLの 404 とエラー画面（404 / 500）のスモークテスト"""




    tmpdir = tempfile.mkdtemp()
    dbfile = os.path.join(tmpdir, "test.sqlite3").replace("\\", "/")
    # TESTING だと例外がそのまま伝わるので、500 の画面を見るために本番と同じ扱いにする
    app = create_app({"SQLALCHEMY_DATABASE_URI": "sqlite:///" + dbfile, "TESTING": False, "PROPAGATE_EXCEPTIONS": False})


    @app.route("/_boom")
    def boom():
        raise RuntimeError("テスト用の失敗")


    client = app.test_client()
    fails = []


    def check(cond, msg):
        print(("OK  " if cond else "NG  ") + msg)
        if not cond:
            fails.append(msg)


    def get(url):
        r = client.get(url)
        return r.status_code, r.data.decode("utf-8"), r


    # 存在しないURL（ログインしていなくても 404。以前はログイン画面へ飛んでいた）
    code, html, r = get("/no-such-page")
    check(code == 404 and "ページが見つかりません" in html and "app-header" in html, "存在しないURLは 404 の画面（ヘッダーつき）")
    code, html, r = get("/cases/999999")
    check(code == 302 and "/login" in r.headers["Location"], "ログインが必要なページは未ログインならログインへ")
    client.post("/login", data={"emp_id": "admin", "password": "adminpass"})
    code, html, r = get("/cases/999999")
    check(code == 404 and "ページが見つかりません" in html, "ログイン後、ない案件は 404 の画面")
    code, html, r = get("/no-such-page")
    check(code == 404 and "ページが見つかりません" in html, "ログイン後も存在しないURLは 404")

    # 500
    code, html, r = get("/_boom")
    check(code == 500 and "エラーが発生しました" in html and "app-header" in html, "失敗したときは 500 の画面（ヘッダーつき）")
    code, html, r = get("/board/")
    check(code == 200, "500 のあとも次のリクエストは普通に動く")

    return fails


def smoke_login_monthly():
    """ログイン（管理者アカウント）と月次アンケート（設問セット・1人1回・当月だけ）のスモークテスト"""




    tmpdir = tempfile.mkdtemp()
    dbfile = os.path.join(tmpdir, "test.sqlite3").replace("\\", "/")
    app = create_app({"SQLALCHEMY_DATABASE_URI": "sqlite:///" + dbfile, "TESTING": True})
    client = app.test_client()
    fails = []


    def check(cond, msg):
        print(("OK  " if cond else "NG  ") + msg)
        if not cond:
            fails.append(msg)


    def get(url, **kw):
        r = client.get(url, **kw)
        return r.status_code, r.data.decode("utf-8"), r


    # ---- 未ログイン: 公開ページは見られ、それ以外はログインへ ----
    code, html, r = get("/")
    check(code == 302 and r.headers["Location"].endswith("/board/"), "未ログインのトップは掲示板へ")
    code, html, r = get("/board/")
    check(code == 200 and "管理者メニュー" in html and "ログイン</a>" not in html and "ダッシュボード" not in html and 'href="/cases/' not in html, "未ログイン: 掲示板は見られ、右上は管理者メニュー、案件へのリンクなし")
    code, html, r = get("/survey/")
    check(code == 200, "未ログイン: 棚卸しアンケート 200")
    code, html, r = get("/monthly/")
    check(code == 200 and "今月のアンケート" in html and "mq-option" in html, "未ログイン: 月次フォーム 200（設問セットが自動作成）")
    for url in ("/cases/", "/bizmap/", "/people/", "/admin/", "/monthly/results", "/monthly/admin"):
        code, html, r = get(url)
        check(code == 302 and "/login" in r.headers["Location"], f"未ログイン: {url} はログインへ")
    code, html, r = get("/login")
    check(code == 200 and "パスワード" in html and "許可リスト" not in html, "ログイン画面")

    # ---- ログイン（固定の管理者アカウント） ----
    r = client.post("/login", data={"emp_id": "admin", "password": "wrong"})
    check(r.status_code == 200 and "正しくありません" in r.data.decode("utf-8"), "パスワード違いは拒否")
    r = client.post("/login", data={"emp_id": "someone", "password": "adminpass"})
    check(r.status_code == 200 and "正しくありません" in r.data.decode("utf-8"), "ID違いは拒否")
    r = client.post("/login", data={"emp_id": "admin", "password": ""})
    check(r.status_code == 200 and "入力してください" in r.data.decode("utf-8"), "空パスワードは拒否")
    r = client.post("/login", data={"emp_id": "admin", "password": "adminpass", "next": "/admin/"})
    check(r.status_code == 302 and r.headers["Location"].endswith("/admin/"), "admin / adminpass でログイン → next へ")
    code, html, r = get("/admin/")
    check(code == 200 and "管理者" in html and "ログイン" in html and "許可リスト" not in html, "ログイン後に管理画面（ヘッダーに管理者）")
    code, html, r = get("/admin/users")
    check(code == 404, "許可リスト画面はなくなった")
    code, html, r = get("/")
    check(code == 200 and "活動の成果" in html, "ログイン後のトップはダッシュボード")
    client.post("/logout")
    code, html, r = get("/")
    check(code == 302, "ログアウト後はトップが掲示板へ")
    conf.override("ADMIN_ID", "chief")
    conf.override("ADMIN_PASSWORD", "secret")
    r = client.post("/login", data={"emp_id": "admin", "password": "adminpass"})
    check(r.status_code == 200, "config.py で変えると既定のIDは使えない")
    r = client.post("/login", data={"emp_id": "chief", "password": "secret"})
    check(r.status_code == 302, "config.py のIDとパスワードでログイン")
    conf.override("ADMIN_ID", "admin")
    conf.override("ADMIN_PASSWORD", "adminpass")
    client.post("/logout")
    r = client.post("/login", data={"emp_id": "admin", "password": "adminpass"})
    check(r.status_code == 302, "既定に戻る")

    # ---- デモデータ（月次は設問セット方式で作られる） ----
    client.post("/admin/demo/load")
    with app.app_context():
        check(MonthlySurvey.query.count() >= 4, f"デモ: 月度ごとの設問セット {MonthlySurvey.query.count()} 件")
        n = MonthlyResponse.query.count()
        sample = MonthlyResponse.query.first()
        check(n > 20 and sample.answers.get("has") in ("ある", "少しある", "ない"), f"デモ: 回答 {n} 件（answers_json）")
    code, html, r = get("/monthly/results")
    check(code == 200 and "繰り返し業務あり" in html and "設問ごとの推移" in html and "trend-table" in html and "コメント" in html, "集計: 推移・設問ごとの推移・月度の詳細")
    code, html, r = get("/")
    check("月次アンケート（時系列）" in html and "自分でやりたい" in html, "ダッシュボードの月次カード")
    code, html, r = get("/monthly/export.csv")
    csv = r.data.decode("utf-8-sig")
    check(code == 200 and "PCでおこなう定型業務" in csv and "コメント" in csv, "CSV に設問の文章が列名で出る")

    # ---- 回答（ログイン不要・1人1回・その場で完了） ----
    client.post("/logout")

    period = current_period()
    code, html, r = get(f"/monthly/?period={period}&emp=E009")
    check(code == 200 and "小林 直樹" in html, "メールリンク: 社員IDで氏名が入る")
    payload = {"period": period, "emp_id": "E009", "name": "小林 直樹", "workplace_id": "", "answers": {"has": "ある", "hours": "30分〜1時間", "trend": "減った", "wants": "ある", "wants_tasks": "日報の転記", "stance": "自分で取り組みたい", "new_tasks": ""}, "comment": "テストのコメント"}
    with app.app_context():
        MonthlyResponse.query.filter_by(period=period, emp_id="E009").delete()
        db.session.commit()
    r = client.post("/monthly/submit", json=payload)
    j = r.get_json()
    check(r.status_code == 200 and j["success"] and j["wants"] and "/survey/start" in j["register_url"], "JSON 送信で受付（本登録への誘導あり）")
    r = client.post("/monthly/submit", json=payload)
    check(r.status_code == 400 and "回答済み" in r.get_json()["message"], "2回目は回答済みで拒否（1人1回）")
    code, html, r = get(f"/monthly/?period={period}&emp=E009")
    check("回答済み" in html and 'id="mq-form" class="card mq-card d-none"' in html, "回答済みの人はフォームの代わりに回答済み表示")
    r = client.get(f"/monthly/api/employee/E009?period={period}")
    check(r.get_json()["answered"] is True, "照会 API が回答済みを返す")
    bad = dict(payload, emp_id="E010", name="加藤 恵", answers=dict(payload["answers"], has=""))
    r = client.post("/monthly/submit", json=bad)
    check(r.status_code == 400 and "has" in r.get_json()["errors"], "必須の選択が空ならエラー")
    bad = dict(payload, emp_id="E010", answers=dict(payload["answers"], has="へんな値"))
    r = client.post("/monthly/submit", json=bad)
    check(r.status_code == 400, "選択肢にない値は拒否")
    # フォーム送信（JavaScript なし）でも受け付ける（デモの回答が残っていれば消してから）
    with app.app_context():
        MonthlyResponse.query.filter_by(period=period, emp_id="E010").delete()
        db.session.commit()
    form = {"period": period, "emp_id": "E010", "name": "加藤 恵", "workplace_id": ""}
    for k, v in payload["answers"].items():
        form["ans-" + k] = v
    form["comment"] = ""
    r = client.post("/monthly/submit", data=form)
    check(r.status_code == 302 and "/monthly/done" in r.headers["Location"], "通常のフォーム送信でも受付")
    code, html, r = get(f"/monthly/done?period={period}")
    check(code == 200 and "回答を受け付けました" in html, "完了ページ")

    # ---- 設問の編集（要ログイン） ----
    r = client.post("/login", data={"emp_id": "admin", "password": "adminpass"})
    code, html, r = get(f"/monthly/admin?period={period}")
    check(code == 200 and "qe-list" in html and "基本設問に戻す" in html, "設問エディタ")
    r = client.get("/monthly/admin/questions?period=default")
    check(r.get_json()["questions"][0]["key"] == "has", "基本設問の取得")
    with app.app_context():
        survey = MonthlySurvey.query.filter_by(period=period).first()
        qs = survey.questions
    qs.append({"key": "", "type": "choice", "text": "改善の勉強会に参加したいですか", "options": ["参加したい", "内容による", "参加しない"], "required": True})
    qs.append({"key": "", "type": "text", "text": "勉強会で学びたいこと", "required": False})
    r = client.post("/monthly/admin/save", json={"period": period, "questions": qs, "is_active": True})
    check(r.status_code == 200 and r.get_json()["success"] and r.get_json()["count"] == len(qs), "設問を追加して保存")
    with app.app_context():
        survey = MonthlySurvey.query.filter_by(period=period).first()
        keys = [q["key"] for q in survey.questions]
        check(len(keys) == len(set(keys)) and all(keys) and "has" in keys, f"保存後のキーは一意で基本設問も残る: {keys[-2:]}")
    code, html, r = get(f"/monthly/results?period={period}")
    check(code == 200 and "設問ごとの推移" in html and "職場別の推移" in html and 'id="chart-q-has"' in html and 'id="chart-q-stance"' in html and "月度の詳細" in html, "集計は時系列中心（設問ごとの推移・職場別の推移・月度の詳細）")
    # 旧2問（頼みたいか／自分でやりたいか）の回答と設問セットは、起動時に新1問へ読み替えられる
    with app.app_context():
        old = MonthlySurvey(period="2024-01", is_active=False)
        old.set_questions([q for q in MonthlySurvey.query.filter_by(period=period).first().questions if q["key"] != "stance"] + [
            {"key": "ask", "type": "choice", "text": "改善チームに頼みたいですか", "options": ["頼みたい", "相談したい", "自分で検討する", "特にない"], "required": True},
            {"key": "self", "type": "choice", "text": "自分で改善に取り組みたいですか", "options": ["取り組みたい", "教えてもらえれば取り組みたい", "今は難しい"], "required": True},
        ])
        db.session.add(old)
        db.session.add(MonthlyResponse(period="2024-01", emp_id="E001", name="田中 一郎", workplace_id=1, answers_json=json.dumps({"has": "ある", "hours": "5〜10時間", "trend": "初回なので分からない", "ask": "相談したい", "self": "今は難しい"}, ensure_ascii=False)))
        db.session.add(MonthlyResponse(period="2024-01", emp_id="E002", name="佐藤 花子", workplace_id=1, answers_json=json.dumps({"has": "ない", "ask": "特にない", "self": "教えてもらえれば取り組みたい"}, ensure_ascii=False)))
        db.session.add(MonthlyResponse(period="2024-01", emp_id="E003", name="鈴木 健", workplace_id=1, answers_json=json.dumps({"wants": "ある", "stance": "教えてもらいながら自分で取り組みたい"}, ensure_ascii=False)))
        db.session.add(MonthlyResponse(period="2024-01", emp_id="E004", name="高橋 美咲", workplace_id=1, answers_json=json.dumps({"wants": "ない", "stance": "今は特にない"}, ensure_ascii=False)))
        db.session.commit()
        check(_migrate_monthly_questions() is True, "移行が走る")
        db.session.commit()
        old = MonthlySurvey.query.filter_by(period="2024-01").first()
        keys = [q["key"] for q in old.questions]
        r1 = MonthlyResponse.query.filter_by(period="2024-01", emp_id="E001").first()
        r2 = MonthlyResponse.query.filter_by(period="2024-01", emp_id="E002").first()
        a2 = r2.answer("stance")
        check(r2.answer("hours") == "0時間", "「ない」で時間が空の回答は 0時間 になる")
        check("stance" in keys and "ask" not in keys and "self" not in keys and r1.answer("stance") == "改善チームに依頼したい" and a2 == "自分で取り組みたい", "旧2問は新1問に置き換わり、回答も読み替えられる")
        a3 = MonthlyResponse.query.filter_by(period="2024-01", emp_id="E003").first().answer("stance")
        a4 = MonthlyResponse.query.filter_by(period="2024-01", emp_id="E004").first().answer("stance")
        check(a3 == "自分で取り組みたい" and a4 == "", "以前の4択は2択に読み替え、「ない」の人の答えは消える")
        check(r1.answer("hours") == "1〜2時間" and r1.answer("trend") == "" and r1.hours_shift == 1.5, "週あたりの時間は1シフトあたりに、「初回なので分からない」は未回答に読み替える")
        hq = next(q for q in old.questions if q["key"] == "hours")
        check(hq["options"][0] == "0時間" and hq["required"] and hq["text"].startswith("1シフト"), "設問セットの時間の設問も新しい文面・選択肢になる")
        check(_migrate_monthly_questions() is False, "2回目は何もしない")
        MonthlyResponse.query.filter_by(period="2024-01").delete()
        db.session.delete(old)
        db.session.commit()
    # 基本設問は文面だけ変えられる。選択肢を変えて保存しても既定に戻り、外すこともできない
    tampered = [dict(q, text="今月、繰り返しのPC作業はありましたか", options=["はい", "いいえ"], required=False) if q["key"] == "has" else q for q in qs]
    r = client.post("/monthly/admin/save", json={"period": period, "questions": tampered, "is_active": True})
    with app.app_context():
        has_q = next(q for q in MonthlySurvey.query.filter_by(period=period).first().questions if q["key"] == "has")
        check(r.status_code == 200 and has_q["options"] == ["ある", "少しある", "ない"] and has_q["required"] and has_q["text"] == "今月、繰り返しのPC作業はありましたか", "基本設問は文面だけ変わり、選択肢と必須は固定のまま")
    r = client.post("/monthly/admin/save", json={"period": period, "questions": [q for q in qs if q["key"] != "stance"], "is_active": True})
    check(r.status_code == 400 and "基本設問は外せません" in r.get_json()["message"], "基本設問を外した保存は拒否")
    code, html, r = get(f"/monthly/?period={period}&emp=E011")
    check("改善の勉強会に参加したいですか" in html and "参加したい" in html, "回答フォームに追加した設問が出る")
    r = client.post("/monthly/admin/save", json={"period": period, "questions": [{"key": "x", "type": "choice", "text": "一択", "options": ["a"]}]})
    check(r.status_code == 400 and "選択肢" in r.get_json()["message"], "選択肢が1つなら保存できない")
    code, html, r = get(f"/monthly/results?period={period}")
    check(code == 200 and "改善の勉強会に参加したいですか" in html and "テストのコメント" in html, "集計に追加設問とコメントが出る")

    # 受付停止
    r = client.post("/monthly/admin/save", json={"period": period, "questions": qs, "is_active": False})
    client.post("/logout")
    code, html, r = get(f"/monthly/?period={period}&emp=E012")
    check("準備中" in html, "受付停止中は準備中の表示")
    r = client.post("/monthly/submit", json=dict(payload, emp_id="E012", name="松本 玲奈"))
    check(r.status_code == 400, "受付停止中は送信できない")

    # ---- 回答できる月度は当月（と設問セットのある前月）だけ ----
    # URLで別の月度を指定しても当月に戻り、設問セットも作られない

    client.post("/login", data={"emp_id": "admin", "password": "adminpass"})
    client.post("/monthly/admin/save", json={"period": period, "questions": qs, "is_active": True})  # 受付を再開
    client.post("/logout")
    with app.app_context():  # 追加した設問にも答え、職場も指定した送信内容
        cur = MonthlySurvey.query.filter_by(period=period).first()
        extra = {q["key"]: (q["options"][0] if q["type"] == "choice" else "") for q in cur.questions if q["key"] not in payload["answers"]}
    full = dict(payload, answers=dict(payload["answers"], **extra), workplace_id="1")
    nxt = shift_period(period, 1)
    prev = shift_period(period, -1)
    old = shift_period(period, -6)
    code, html, r = get(f"/monthly/?period={nxt}")
    with app.app_context():
        check(MonthlySurvey.query.filter_by(period=nxt).first() is None and f'name="period" value="{period}"' in html, "翌月を指定しても当月のフォーム（翌月の設問セットは作られない）")
    code, html, r = get(f"/monthly/?period={old}")
    with app.app_context():
        check(MonthlySurvey.query.filter_by(period=old).first() is None and f'name="period" value="{period}"' in html, "半年前を指定しても当月のフォーム")
        has_prev = MonthlySurvey.query.filter_by(period=prev).first() is not None
        check((prev in answerable_periods()) == has_prev and period in answerable_periods() and nxt not in answerable_periods(), f"回答できる月度は当月{'と前月' if has_prev else 'だけ'}")
        check(public_period(nxt) == period and public_period("") == period and public_period(prev) == (prev if has_prev else period), "public_period の丸め")
        MonthlySurvey.query.filter_by(period=prev).delete()
        db.session.commit()
    r = client.post("/monthly/submit", json=dict(full, period=old, emp_id="E013", name="試験 昔"))
    with app.app_context():
        check(r.status_code == 200 and MonthlyResponse.query.filter_by(period=old).count() == 0 and MonthlyResponse.query.filter_by(period=period, emp_id="E013").count() == 1, "半年前の月度で送っても当月の回答として受け付ける")
    r = client.post("/monthly/submit", json=dict(full, period=prev, emp_id="E014", name="試験 前月"))
    with app.app_context():
        check(r.status_code == 200 and MonthlySurvey.query.filter_by(period=prev).first() is None and MonthlyResponse.query.filter_by(period=prev, emp_id="E014").count() == 0 and MonthlyResponse.query.filter_by(period=period, emp_id="E014").count() == 1, "設問セットのない前月も同じ（前月の設問セットは作られない）")
        cur = MonthlySurvey.query.filter_by(period=period).first()
        db.session.add(MonthlySurvey(period=prev, is_active=True, questions_json=cur.questions_json))
        db.session.commit()
        check(prev in answerable_periods(), "前月の設問セットを作ると前月も回答できる")
    code, html, r = get(f"/monthly/?period={prev}&emp=E015")
    check(code == 200 and f'name="period" value="{prev}"' in html and period_label(prev) in html, "前月のフォームが開く")
    r = client.post("/monthly/submit", json=dict(full, period=prev, emp_id="E015", name="試験 前月回答"))
    with app.app_context():
        check(r.status_code == 200 and MonthlyResponse.query.filter_by(period=prev, emp_id="E015").count() == 1, "前月に回答できる")

    # ---- 職場は必須（照会で分からなければ選ぶ） ----
    r = client.post("/monthly/submit", json=dict(full, period=period, emp_id="X999", name="新人 太郎", workplace_id=""))
    check(r.status_code == 400 and "workplace_id" in r.get_json()["errors"], "職場が分からないときは送信できない")
    r = client.post("/monthly/submit", json=dict(full, period=period, emp_id="X999", name="新人 太郎"))
    check(r.status_code == 200 and r.get_json()["success"], "職場を選べば受付")
    code, html, r = get("/monthly/?emp=X999")
    check('id="workplace_id" class="form-select" required' in html and "職場を選んでください" in html, "職場の選択は必須の入力欄")

    # ---- 条件つきの設問 ----
    base_ans = dict(full["answers"], wants="ある", wants_tasks="", trend="変わらない", new_tasks="")
    r = client.post("/monthly/submit", json=dict(full, emp_id="X101", name="条件 一", workplace_id="1", answers=base_ans))
    check(r.status_code == 400 and "wants_tasks" in r.get_json()["errors"], "「ある」なのに業務名が空なら送信できない")
    r = client.post("/monthly/submit", json=dict(full, emp_id="X101", name="条件 一", workplace_id="1", answers=dict(base_ans, wants="ない", wants_tasks="消えるはず")))
    with app.app_context():
        row = MonthlyResponse.query.filter_by(period=period, emp_id="X101").first()
        check(r.status_code == 200 and row is not None and row.answer("wants_tasks") == "", "「ない」なら業務名は受け取らない（空で保存）")
    r = client.post("/monthly/submit", json=dict(full, emp_id="X102", name="条件 二", workplace_id="1", answers=dict(base_ans, wants="ない", trend="増えた", new_tasks="")))
    check(r.status_code == 400 and "new_tasks" in r.get_json()["errors"], "「増えた」なのに新規業務が空なら送信できない")
    r = client.post("/monthly/submit", json=dict(full, emp_id="X102", name="条件 二", workplace_id="1", answers=dict(base_ans, wants="ない", trend="減った", new_tasks="消えるはず")))
    with app.app_context():
        row = MonthlyResponse.query.filter_by(period=period, emp_id="X102").first()
        check(r.status_code == 200 and row is not None and row.answer("new_tasks") == "", "「減った」なら新規業務は受け取らない（空で保存）")
    code, html, r = get("/monthly/?emp=X103")
    check('data-dep-key="wants"' in html and 'data-dep-key="trend"' in html and 'data-badge="req"' in html, "フォームに条件つき設問の情報が出る")
    check('<span class="mq-no">3-1</span>' in html and '<span class="mq-no">4-1</span>' in html and '<span class="mq-no">4-2</span>' in html and 'data-dep-show=' in html, "枝番: 3-1（増えた業務名）、4-1（業務名）、4-2（進め方）")
    r = client.post("/monthly/submit", json=dict(full, emp_id="X105", name="条件 五", workplace_id="1", answers=dict(base_ans, wants="ない", trend="変わらない", new_tasks="消えるはず")))
    with app.app_context():
        row = MonthlyResponse.query.filter_by(period=period, emp_id="X105").first()
        check(r.status_code == 200 and row is not None and row.answer("new_tasks") == "", "「変わらない」なら増えた業務名は出ない（受け取らない）")
    r = client.post("/monthly/submit", json=dict(full, emp_id="X104", name="条件 三", workplace_id="1", answers=dict(base_ans, wants="ある", wants_tasks="日報の転記", stance="")))
    check(r.status_code == 400 and r.get_json()["errors"].get("stance") == "選択してください", "「ある」なのに進め方を選んでいないと送信できない")
    r = client.post("/monthly/submit", json=dict(full, emp_id="X104", name="条件 三", workplace_id="1", answers=dict(base_ans, wants="ない", stance="自分で取り組みたい")))
    with app.app_context():
        row = MonthlyResponse.query.filter_by(period=period, emp_id="X104").first()
        check(r.status_code == 200 and row is not None and row.answer("stance") == "", "「ない」なら進め方は受け取らない（空で保存）")

    # ---- 照会が「IDがあるか」しか返さない環境（氏名・職場なし） ----


    class _FakeResp:
        def __init__(self, body):
            self.body = body

        def read(self):
            return self.body

        def __enter__(self):
            return self

        def __exit__(self, *args):
            return False


    def _fake_urlopen(req, timeout=0):
        url = req.full_url
        if "E999" in url:
            raise urllib.error.HTTPError(url, 404, "Not Found", {}, None)
        if "E998" in url:
            return _FakeResp(b'{"exists": false}')
        if "E997" in url:
            return _FakeResp(json.dumps({"exists": True, "name": "名前 あり", "department": "製造1課"}).encode("utf-8"))
        if "E996" in url:
            return _FakeResp(json.dumps({"employee_id": "E996", "display_name": "表示名 太郎"}).encode("utf-8"))
        return _FakeResp(b'{"exists": true}')


    _real_urlopen = ldap_auth.urllib.request.urlopen
    ldap_auth.urllib.request.urlopen = _fake_urlopen
    conf.override("LDAP_API_URL", "http://ldap.test/api?employee_id={id}")
    with app.app_context():
        check(lookup("E999") is None and lookup("E998") is None, "404 や exists:false は「いない」")
        info = lookup("E777")
        check(info is not None and info["name"] == "" and info["workplace_id"] is None and Employee.query.filter_by(emp_id="E777").first() is not None, "exists:true だけでも「いる」と分かる（氏名・職場は空。控えに行ができる）")
        info = lookup("E997")
        check(info is not None and info["name"] == "名前 あり" and info["workplace"] == "製造1課", "氏名・所属が返れば従来どおり")
        info = lookup("E996")
        check(info is not None and info["name"] == "表示名 太郎" and info["workplace_id"] is None, "氏名が display_name で返る API でも氏名が入る（職場は空）")
    r = client.get(f"/monthly/api/employee/E777?period={period}")
    j = r.get_json()
    check(j["found"] and j["name"] == "" and not j["workplace_id"], "照会APIで見つかるが氏名・職場は空")
    r = client.get(f"/monthly/api/employee/E996?period={period}")
    j = r.get_json()
    check(j["found"] and j["name"] == "表示名 太郎" and not j["workplace_id"], "display_name の人は氏名だけ入り、職場を選ぶ")
    r = client.post("/monthly/submit", json=dict(full, period=period, emp_id="E996", name="", workplace_id=""))
    check(r.status_code == 400 and set(r.get_json()["errors"]) == {"workplace_id"}, "初回に足りないのは職場だけ")
    r = client.post("/monthly/submit", json=dict(full, period=period, emp_id="E996", name="", workplace_id="3"))
    check(r.status_code == 200, "職場を選んで送信")
    r = client.get(f"/monthly/api/employee/E996?period={prev}")
    j = r.get_json()
    check(j["found"] and j["name"] == "表示名 太郎" and j["workplace_id"] == 3, "2回目からは職場も自動で入る")
    r = client.post("/monthly/submit", json=dict(full, period=period, emp_id="E777", name="", workplace_id=""))
    check(r.status_code == 400 and {"name", "workplace_id"} <= set(r.get_json()["errors"]), "初回は氏名と職場を入れるまで送信できない")
    r = client.post("/monthly/submit", json=dict(full, period=period, emp_id="E777", name="初回 入力", workplace_id="2"))
    check(r.status_code == 200, "初回に氏名と職場を入れて送信")
    r = client.get(f"/monthly/api/employee/E777?period={prev}")
    j = r.get_json()
    with app.app_context():
        emp = Employee.query.filter_by(emp_id="E777").first()
        check(j["found"] and j["name"] == "初回 入力" and j["workplace_id"] == 2 and emp.name == "初回 入力" and emp.workplace_id == 2, "2回目からは氏名・職場が自動で入る（控えに残る）")
    client.post("/survey/api/reset")
    r = client.post("/survey/api/respondent", json={"emp_id": "E777", "name": "", "workplace_id": ""})
    j = r.get_json()
    check(r.status_code == 200 and j["respondent"]["name"] == "初回 入力" and j["respondent"]["workplace_id"] == 2, "自動化依頼でも社員IDだけで氏名・職場が入る")
    ldap_auth.urllib.request.urlopen = _real_urlopen
    conf.override("LDAP_API_URL", "")

    # ---- DX人財情報 ----
    client.post("/login", data={"emp_id": "admin", "password": "adminpass"})
    code, html, r = get("/people/")
    check(code == 200 and "DX人財情報" in html and "人材育成" not in html and "田中 一郎" in html and "初回 入力" in html, "DX人財情報: 自動化依頼の回答者と、月次だけ答えた人が並ぶ")
    check("自分で改善したい" in html and "直近の月次で「自分で取り組みたい」" in html and "人財一覧" in html, "DX人財情報: 集計カードと一覧")
    r = client.get("/people/", query_string={"wish": "positive"})
    html = r.data.decode("utf-8")
    check(r.status_code == 200 and "佐藤 花子" in html, "自分で改善したい人で絞り込み")
    r = client.get("/people/", query_string={"wish": "今は考えていない"})
    html = r.data.decode("utf-8")
    check(r.status_code == 200 and "渡辺 由美" in html and "佐藤 花子" not in html, "意欲の選択肢で絞り込み")
    r = client.get("/people/", query_string={"workplace_id": 99})
    check(r.status_code == 200 and "該当する人はいません" in r.data.decode("utf-8"), "該当なしの表示")

    return fails


def smoke_request():
    """自動化依頼（1画面・JSON API）のスモークテスト"""




    tmpdir = tempfile.mkdtemp()
    dbfile = os.path.join(tmpdir, "test.sqlite3").replace("\\", "/")
    app = create_app({"SQLALCHEMY_DATABASE_URI": "sqlite:///" + dbfile, "TESTING": True})
    client = app.test_client()
    fails = []


    def check(cond, msg):
        print(("OK  " if cond else "NG  ") + msg)
        if not cond:
            fails.append(msg)


    # 入口（ログイン不要・1画面）
    r = client.get("/survey/")
    html = r.data.decode("utf-8")
    check(r.status_code == 200 and 'id="rq"' in html and "自動化依頼" in html and "パート" not in html and "rq-options" in html, "入口は1画面（定義を埋め込み）")
    check("業務登録</a>" not in html and "自動化依頼</a>" in html, "ナビは自動化依頼")
    r = client.get("/survey/start?emp=E001")
    check(r.status_code == 302 and r.headers["Location"].endswith("/survey/?emp=E001"), "旧リンクは入口へ（社員IDつき）")

    task = {
        "name": "日報の数字を週報に転記する", "domain_id": "1", "purpose": "週報を作るため", "output": "週報（Excel）",
        "output_users": ["上司"], "output_users_note": "", "inputs": ["他のExcel"], "inputs_note": "",
        "frequency": "毎週", "duration": "15〜30分", "headcount": 1, "tools": ["Excel"], "nature": ["転記・コピー"],
        "pains": ["時間がかかる"], "pains_note": "", "wish": "自動化したい", "self_improve": "取り組みたい", "note": "",
        "prev_task": "", "next_task": "", "actors": "", "goods": [], "goods_note": "", "goods_from": "", "goods_to": "", "info_from": "",
    }
    r = client.post("/survey/api/tasks", json=task)
    check(r.status_code == 400 and r.get_json()["errors"].get("_") == "respondent", "依頼者未保存では業務を登録できない")

    # あなたのこと
    r = client.post("/survey/api/respondent", json={"emp_id": "", "name": "x", "workplace_id": ""})
    j = r.get_json()
    check(r.status_code == 400 and {"emp_id", "workplace_id"} <= set(j["errors"]), "あなたのこと: 社員IDと職場の検証")
    r = client.post("/survey/api/respondent", json={"emp_id": "ZZZ9", "name": "", "workplace_id": ""})
    check(r.status_code == 400 and {"name", "workplace_id"} <= set(r.get_json()["errors"]), "見つからないIDは氏名・職場が必要")
    r = client.post("/survey/api/respondent", json={"emp_id": "T001", "name": "試験 太郎", "workplace_id": 1})
    j = r.get_json()
    check(r.status_code == 200 and j["success"] and not j["resumed"] and j["respondent"]["name"] == "試験 太郎" and j["tasks"] == [] and "role" not in j["respondent"], "あなたのこと: 保存（役割・PC利用時間の項目はない）")
    # 同姓同名・同じ職場でも社員IDが違えば別人
    client.post("/survey/api/reset")
    r = client.post("/survey/api/respondent", json={"emp_id": "T002", "name": "試験 太郎", "workplace_id": 1})
    j = r.get_json()
    check(r.status_code == 200 and not j["resumed"] and j["respondent"]["emp_id"] == "T002", "同じ氏名・職場でも社員IDが違えば新しい依頼者")
    with app.app_context():
        check(Respondent.query.filter_by(name="試験 太郎").count() == 2, "依頼者は2人になる")
    client.post("/survey/api/reset")
    r = client.post("/survey/api/respondent", json={"emp_id": "T001", "name": "", "workplace_id": ""})
    j = r.get_json()
    check(r.status_code == 200 and j["resumed"], "T001 に戻る")

    # 業務の登録・更新・削除
    r = client.post("/survey/api/tasks", json=dict(task, name="", frequency="週2回"))
    j = r.get_json()
    check(r.status_code == 400 and "name" in j["errors"] and "frequency" in j["errors"], "業務: 必須と選択肢の検証")
    r = client.post("/survey/api/tasks", json=task)
    j = r.get_json()
    check(r.status_code == 200 and j["success"] and j["task"]["hours"] > 0 and len(j["tasks"]) == 1, f"業務: 登録（月 {j['task']['hours']} h）")
    tid = j["task"]["id"]
    r = client.put(f"/survey/api/tasks/{tid}", json=dict(task, name="日報の数字を週報へ自動転記", frequency="毎日"))
    j = r.get_json()
    check(r.status_code == 200 and j["task"]["name"].endswith("自動転記") and j["task"]["frequency"] == "毎日", "業務: 更新")
    r = client.post("/survey/api/tasks", json=dict(task, name="2件目の業務"))
    check(len(r.get_json()["tasks"]) == 2, "業務: 2件目")
    r = client.delete(f"/survey/api/tasks/{tid}")
    j = r.get_json()
    check(r.status_code == 200 and len(j["tasks"]) == 1 and j["tasks"][0]["name"] == "2件目の業務", "業務: 削除")

    # 関わり方・全体
    r = client.post("/survey/api/part-c", json={"c1_wish": "", "c2_skills": []})
    check(r.status_code == 400 and "c1_wish" in r.get_json()["errors"], "関わり方: 必須の検証")
    r = client.post("/survey/api/part-c", json={"c1_wish": "ぜひ取り組みたい", "c2_skills": ["Excel関数"], "c3_learn": ["Python"], "c4_hours": "1〜2時間", "c5_targets": "2件目の業務", "c6_share": "できる"})
    check(r.status_code == 200 and r.get_json()["respondent"]["part_c_done"], "関わり方: 保存")
    r = client.post("/survey/api/part-d", json={"d1": "総務へ毎月送っている", "d2": "", "d3": "", "d4": ""})
    j = r.get_json()
    check(r.status_code == 200 and j["respondent"]["part_d_done"] and len(j["tasks"]) == 1, "全体: 送信して完了")
    with app.app_context():
        rsp = Respondent.query.filter_by(emp_id="T001").first()
        check(rsp.part_c_done and rsp.part_d_done and rsp.d1 == "総務へ毎月送っている" and Task.query.filter_by(respondent_id=rsp.id).count() == 1, "DB に保存されている")

    # 同じ社員IDで再開 → 前回の内容が返る
    client.post("/survey/api/reset")
    r = client.get("/survey/")
    check('"respondent": null' in r.data.decode("utf-8"), "別の人: 状態が消える")
    r = client.post("/survey/api/respondent", json={"emp_id": "T001", "name": "", "workplace_id": ""})
    j = r.get_json()
    check(r.status_code == 200 and j["resumed"] and j["respondent"]["name"] == "試験 太郎" and len(j["tasks"]) == 1, "同じ社員IDで続きから（氏名はDBから）")
    r = client.get("/survey/")
    check('"name": "試験 太郎"' in r.data.decode("utf-8"), "再表示で状態が埋め込まれる")

    # 月次に答えたことがある人は、社員IDだけで氏名・職場が入る
    with app.app_context():
        db.session.add(MonthlyResponse(period="2026-01", emp_id="M001", name="月次 花子", workplace_id=2, answers_json="{}"))
        db.session.commit()
    client.post("/survey/api/reset")
    r = client.post("/survey/api/respondent", json={"emp_id": "M001", "name": "", "workplace_id": ""})
    j = r.get_json()
    check(r.status_code == 200 and not j["resumed"] and j["respondent"]["name"] == "月次 花子" and j["respondent"]["workplace_id"] == 2, "月次の回答から氏名・職場を補う")

    # ログイン後の案件一覧・詳細の文言
    client.post("/login", data={"emp_id": "admin", "password": "adminpass"})
    r = client.get("/cases/")
    html = r.data.decode("utf-8")
    check("職場／依頼者" in html and "自動化依頼を出す" in html, "案件一覧の文言")

    return fails


def smoke_writeback():
    """書き戻し（writeback）のスモークテスト"""




    tmpdir = tempfile.mkdtemp()
    dbfile = os.path.join(tmpdir, "test.sqlite3").replace("\\", "/")
    app = create_app({"SQLALCHEMY_DATABASE_URI": "sqlite:///" + dbfile, "TESTING": True})
    client = app.test_client()
    client.post("/login", data={"emp_id": "admin", "password": "adminpass"})
    fails = []


    def check(cond, msg):
        print(("OK  " if cond else "NG  ") + msg)
        if not cond:
            fails.append(msg)


    client.post("/admin/demo/load")

    # 書き戻し: 情報の流れ
    g = client.get("/bizmap/api/graph?kind=info").get_json()
    nodes = {n["id"]: n for n in g["nodes"]}
    with app.app_context():
        t1 = Task.query.filter_by(name="日報の実績をExcelの集計表に転記する").first()
        t2 = Task.query.filter_by(name="週報の作成").first()
        t3 = Task.query.filter_by(name="残業予定表の提出").first()
        t1_id, t2_id, t3_id = t1.id, t2.id, t3.id
        # 自動生成の「次の業務」矢印（t1 -> t2）が存在することを確認
        check(f"t{t1_id}->t{t2_id}" in {e["id"] for e in g["edges"]}, "auto seq edge t1->t2 exists")
        sys_node = next(nid for nid, n in nodes.items() if n["cls"] == "system")
        ext_node = next(nid for nid, n in nodes.items() if n["cls"] == "ext")
        src_node = next(nid for nid, n in nodes.items() if n["cls"] == "source")

    data = {"nodes": [], "edges": [], "suppressed_nodes": [], "suppressed_edges": [f"t{t1_id}->t{t2_id}"], "view": None}
    for n in g["nodes"]:
        m = dict(n)
        m["x"], m["y"] = 10, 10
        if n["id"] == f"t{t3_id}":
            m["text"] = "残業予定の連絡"
            m["text_edited"] = True
        data["nodes"].append(m)
    data["nodes"].append({"id": "m_note1", "text": "翌日の計画確認", "cls": "note", "source": "manual", "x": 5, "y": 5})
    for e in g["edges"]:
        if e["id"] == f"t{t1_id}->t{t2_id}":
            continue
        data["edges"].append(dict(e))
    manual = [
        {"from": f"t{t1_id}", "to": f"t{t3_id}"},            # 業務→業務: 前後の業務
        {"from": sys_node, "to": f"t{t3_id}"},                # システム→業務: 入力情報
        {"from": src_node, "to": f"t{t3_id}"},                # 情報源→業務: 情報の受け取り元
        {"from": f"t{t3_id}", "to": ext_node},                # 業務→他部署: 利用者
        {"from": f"t{t3_id}", "to": "m_note1"},               # 業務→メモ: 後の業務
    ]
    for m in manual:
        data["edges"].append({"id": f"{m['from']}->{m['to']}", "from": m["from"], "to": m["to"], "label": m.get("label", ""), "kind": "manual", "source": "manual", "reversed": False, "label_edited": bool(m.get("label"))})
    r = client.post("/bizmap/api/graph", json={"kind": "info", "data": data})
    check(r.status_code == 200, "save graph with manual edits")

    r = client.get("/bizmap/api/writeback?kind=info")
    cands = r.get_json()["candidates"]
    texts = [c["text"] for c in cands]
    print("candidates:")
    for t in texts:
        print("   -", t)
    check(any("業務名「残業予定表の提出」を「残業予定の連絡」に変更" in t for t in texts), "候補: 業務名の変更")
    check(any("後の業務に「残業予定表の提出」" in t for t in texts), "候補: 前後の業務（業務→業務）")
    check(any("入力情報に「" in t for t in texts), "候補: システム→業務")
    check(any("情報の受け取り元に「" in t for t in texts), "候補: 情報源→業務")
    check(any("アウトプットの利用者に「" in t for t in texts), "候補: 業務→他部署")
    check(any("後の業務に「翌日の計画確認」" in t for t in texts), "候補: 業務→メモ")
    check(any("後の業務から「週報の作成」" in t for t in texts), "候補: 消した前後の矢印")

    ids = [c["id"] for c in cands]
    r = client.post("/bizmap/api/writeback", json={"kind": "info", "ids": ids})
    res = r.get_json()
    check(r.status_code == 200 and res["ok"] and res["applied"] >= 7, f"apply writeback applied={res.get('applied')}")
    with app.app_context():
        t1 = db.session.get(Task, t1_id)
        t2 = db.session.get(Task, t2_id)
        t3 = db.session.get(Task, t3_id)
        check(t3.name == "残業予定の連絡", "業務名が変更された")
        check("残業予定" in (t1.next_task or ""), f"t1.next_task に追加: {t1.next_task}")
        check("日報" in (t3.prev_task or ""), f"t3.prev_task に追加: {t3.prev_task}")
        check("週報" not in (t1.next_task or ""), f"t1.next_task から週報を外した: {t1.next_task}")
        check("日報" not in (t2.prev_task or ""), f"t2.prev_task から日報を外した: {t2.prev_task}")
        check(len(t3.inputs_list) >= 2, f"inputs に追加: {t3.inputs_list}")
        check(t3.info_from, f"info_from: {t3.info_from}")
        check("他部署" in t3.output_users_list and t3.output_users_note, f"利用者: {t3.output_users_list} / {t3.output_users_note}")
        check("翌日の計画確認" in (t3.next_task or ""), f"next_task メモ: {t3.next_task}")

    # 反映後に再生成すると、手動の矢印が自動の矢印として現れる（重複しない）
    g2 = client.get("/bizmap/api/graph?kind=info").get_json()
    eids = [e["id"] for e in g2["edges"]]
    check(len(eids) == len(set(eids)), "再生成後にエッジIDの重複なし")
    check(f"t{t1_id}->t{t3_id}" in eids, "反映後: 業務→業務が自動生成される")
    auto_seq = next(e for e in g2["edges"] if e["id"] == f"t{t1_id}->t{t3_id}")
    check(auto_seq["source"] == "auto", "反映後: 自動エッジとして扱われる")
    check(f"t{t1_id}->t{t2_id}" not in eids, "反映後: 外した前後の矢印は再生成されない")
    r = client.get("/bizmap/api/writeback?kind=info")
    check(len(r.get_json()["candidates"]) <= 2, f"再反映の候補は残らない（{len(r.get_json()['candidates'])}件）")

    r = client.get("/bizmap/edit?kind=info")
    check(r.status_code == 200 and "fe-writeback" in r.data.decode("utf-8"), "editor has writeback button")
    r = client.get("/bizmap/api/writeback?kind=info&domain_id=99")
    check(r.status_code == 200 and r.get_json()["candidates"] == [], "writeback without saved map -> empty")

    return fails


TESTS = [("smoke_bizmap_edit", smoke_bizmap_edit), ("smoke_board", smoke_board), ("smoke_errors", smoke_errors), ("smoke_login_monthly", smoke_login_monthly), ("smoke_request", smoke_request), ("smoke_writeback", smoke_writeback)]


def _run_one(name):
    fails = dict(TESTS)[name]()
    print()
    print("FAILURES:", len(fails))
    for f in fails:
        print(" -", f)
    return 1 if fails else 0


if __name__ == "__main__":
    if len(sys.argv) > 1:
        sys.exit(_run_one(sys.argv[1]))
    results = []
    for name, _fn in TESTS:
        print("=" * 70)
        print(name)
        print("=" * 70)
        rc = subprocess.run([sys.executable, "-X", "utf8", os.path.abspath(__file__), name], cwd=os.path.dirname(os.path.abspath(__file__))).returncode
        results.append((name, rc))
        print()
    print("=" * 70)
    failed = [n for n, rc in results if rc != 0]
    for name, rc in results:
        print(("PASS  " if rc == 0 else "FAIL  ") + name)
    print(f"{len(results) - len(failed)} / {len(results)} passed")
    sys.exit(1 if failed else 0)
