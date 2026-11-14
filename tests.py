"""スモークテスト。python tests.py で全部、python tests.py smoke_board のように名前を付けると1本。一時ファイルの SQLite を使う"""
import io
import json
import os
import subprocess
import sys
import tempfile
import urllib.error

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import ldap_auth
from app import conf, create_app
from app import Employee, FlowMap, MonthlyResponse, MonthlySurvey, Respondent, StageLog, Task, _migrate_monthly_questions, answerable_periods, current_period, db, lookup, period_label, public_period, shift_period

for _k, _v in (("LDAP_API_URL", ""), ("LDAP_SERVER", ""), ("ADMIN_ID", "admin"), ("ADMIN_PASSWORD", "adminpass")):
    conf.override(_k, _v)

def smoke_bizmap_edit():
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

    import app as app_module
    with app.app_context():
        media = {t: app_module.medium_of(t) for t in ("残業予定表（Excel）をメール送付", "不良報告メール", "メール", "紙の日報", "生産管理システム", "共有フォルダのExcel", "パスワード表", "ＰＤＦ", "各ラインリーダー（口頭）", "謎のもの")}
        check(list(media.values()) == ["excel", "mail", "unknown", "paper", "system", "excel", "doc", "pdf", "verbal", "unknown"], f"情報の媒体を名前から判定（区分の設定の語・除く語・全角）: {media}")
        who = {t: (app_module.who_of(t, "製造1課")["actor"], app_module.who_of(t, "製造1課")["label"]) for t in ("課長", "品質保証課 課長", "生産管理課の進捗確認", "製造1課（職場内）", "山田さん", "顧客", "日常点検", "分からない")}
        check(who == {"課長": ("boss", "製造1課 課長"), "品質保証課 課長": ("dept", "検査・品質保証課 課長"), "生産管理課の進捗確認": ("dept", "生産管理課の進捗確認"),
                      "製造1課（職場内）": ("team", "製造1課の人"), "山田さん": ("", "担当者"), "顧客": ("external", "顧客"), "日常点検": ("", "日常点検"), "分からない": ("unknown", "利用者不明")},
              f"作業のだれ: 職場マスタ → 区分の語 → 未確認。役職だけなら職場名を付け、氏名は担当者（だれ未確認）に置き換える: {who}")
        more = {t: (app_module.who_of(t, "製造1課")["actor"], app_module.who_of(t, "製造1課")["label"], app_module.who_of(t, "製造1課")["state"])
                for t in ("佐藤課長", "営業さん", "品証の佐藤さん", "課長へ提出する日報の作成", "田中 一郎")}
        check(more == {"佐藤課長": ("boss", "製造1課 課長", "ok"), "営業さん": ("dept", "営業さん", "ok"), "品証の佐藤さん": ("dept", "品証の担当者", "ok"),
                       "課長へ提出する日報の作成": ("", "課長へ提出する日報の作成", "unconfirmed"), "田中 一郎": ("", "担当者", "unconfirmed")},
              f"名字＋役職は役職、部署に敬称はそのまま、社員名簿の氏名は担当者、行き先（〜へ）はだれにしない: {more}")
        masked = [app_module.mask_person_names(t) for t in ("佐藤課長に渡す", "お客様と営業さんに送る", "記入様式に書く", "佐藤、鈴木", "千葉工場の検査員さん")]
        check(masked == ["課長に渡す", "お客様と営業さんに送る", "記入様式に書く", "担当者、担当者", "千葉工場の検査員さん"], f"氏名の置き換えは名字・敬称・役職の形だけ（様式・工場名はそのまま）: {masked}")
        check(app_module.same_info("日次実績集計表", "日次実績集計表（Excel）") and not app_module.same_info("実績表", "日次実績集計表") and not app_module.same_info("週報", "週報の作成"),
              "業務をまたぐ情報の一致は完全一致か「名前（媒体）」だけ（部分一致でつながない）")
        check([app_module.handoff_of(t) for t in ("メールで届く", "システムで見る", "朝会で配る")] == ["push", "pull", "push"], "受け手側の動詞（届く・見る）でも引き継ぎ方を決める")
        check([app_module.field_role(t) for t in ("メール", "生産管理システム", "Excel（共有フォルダ）", "紙の日報")] == [("line", "mail"), ("both", "system"), ("both", "folder"), ("box", "")],
              "手順の欄: 渡し方は線の方法、置き場は箱と方法、それ以外は情報の箱")
        check([app_module.handoff_of(t) for t in ("課長に渡す", "各ラインから日報を集める", "前週比のコメントを書く", "生産管理課に問い合わせる")] == ["push", "pull", "", "pull"],
              "引き継ぎ方は動詞（伝えに行く語・取りに行く語）で決める")
        check(app_module.mask_person_names("山田さん、課長さん、品質保証課") == "担当者、課長、品質保証課", "氏名は担当者に置き換え、役職は残す")
        dg = app_module.Diagram()
        dg.node("s_a", "日次実績集計表（Excel）", "source")
        dg.node("s_b", "生産管理システム", "source")
        dg.node("t1", "週報の作成", "task", who={"actor": "team", "label": "製造1課", "state": "ok"})
        dg.node("p_a", "朝会で確認する", "work", who=app_module.who_of("課長", "製造1課"))
        dg.edge("s_a", "t1", "info")
        dg.edge("t1", "p_a", "work", handoff="push")
        dg.edge("s_b", "t1", "info", method="system", show_method=app_module.show_method_for("system", "system"))
        src = dg.render()
        check("s_a:::source@{ shape: doc, label:" in src and "bi-file-earmark-excel-fill" in src and 's_b[("' in src and 't1["週報の作成"]:::task' in src
              and "<b>製造1課 課長</b>" in src and 'p_a["' in src and '== "前が伝えに行く" ==>' in src and "linkStyle 1 stroke:#495057" in src and "システム（登録・参照）" not in src,
              "Mermaid: 作業は四角（だれを太字で）、情報は媒体の形とアイコン、仕事の流れは太い灰の矢印に引き継ぎ方、媒体から分かる方法は書かない")
        legend = app_module.map_legend_html(dg)
        check("箱と線" in legend and "仕事の流れ" in legend and "情報の流れ" in legend and "Excel" in legend and "PowerPoint" not in legend and "上司・役職" in legend, "凡例は箱と線の2種類と、図に出てくる色・媒体・だれだけ")
        check('"how"' in app_module.INTAKE_PROMPT and '"use"' in app_module.INTAKE_PROMPT and '"how_from"' in app_module.INTAKE_PROMPT and "誰にどうやって渡し" in app_module.INTAKE_PROMPT
              and "氏名・〇〇さん" in app_module.INTAKE_PROMPT and "{methods}" not in app_module._intake_prompt() and "state.purpose" in app_module.FINALIZE_PROMPT,
              "聞き取り: 渡す相手・渡し方・使いみちを1問で聞き、受け取り方も集め、氏名は書かない")
    html = client.get("/bizmap/").data.decode("utf-8")
    check("map-legend" in html and "箱と線" in html, "業務マップ: 箱と線・色・媒体・だれの凡例")
    js = client.get("/static/flows.js").data.decode("utf-8")
    check("最大表示" in js and "requestFullscreen" in js and "flow-max" in js and "bootstrap-icons" in js, "閲覧: 図の最大表示（Esc で戻る）とアイコンの読み込み待ち")
    html = client.get("/bizmap/edit?kind=info").data.decode("utf-8")
    fe = client.get("/static/flowedit.js").data.decode("utf-8")
    check('id="fe-max"' in html and 'id="fe-add-work"' in html and 'id="fe-add-info"' in html and 'id="fe-type-row"' in html and 'id="fe-method-select"' in html and 'id="fe-handoff-select"' in html
          and "classify(" in fe and "情報どうしは結べません" in fe and "requestFullscreen" in fe, "編集: ＋作業・＋情報、種類・だれ・媒体、線の方法・引き継ぎ方、情報どうしは結べない")
    r = client.post("/bizmap/api/classify", json={"items": [{"id": "a", "type": "info", "text": "週報（Excel）"}, {"id": "b", "type": "work", "who": "課長", "workplace": "製造1課"}]})
    check(r.get_json()["items"] == [{"id": "a", "medium": "excel"}, {"id": "b", "actor": "boss", "who": "製造1課 課長", "who_state": "ok"}], "編集画面の判定はサーバーに問い合わせる（区分の設定と同じ判定）")

    from html.parser import HTMLParser

    class DivTree(HTMLParser):
        """div の入れ子だけを追い、id のある div がどの div の中にあるかを控える"""
        def __init__(self):
            super().__init__()
            self.stack, self.inside = [], {}

        def handle_starttag(self, tag, attrs):
            if tag != "div":
                return
            i = dict(attrs).get("id")
            if i:
                self.inside[i] = [x for x in self.stack if x]
            self.stack.append(i)

        def handle_endtag(self, tag):
            if tag == "div" and self.stack:
                self.stack.pop()

    tree = DivTree()
    tree.feed(html)
    check("fe-app" in tree.inside.get("fe-container", []) and "fe-app" in tree.inside.get("fe-node-popover", []) and "fe-app" in tree.inside.get("fe-edge-popover", []),
          f"編集画面: 図の領域とポップアップは #fe-app の中（閉じタグの数が合っている）: {tree.inside.get('fe-container')}")
    check("auto_method" in fe and "classifySeq" in fe and "history[historyIndex] = snapshot()" in fe and "SELECT" in fe and "pop.offsetHeight" in fe,
          "編集: 「自動」で元の方法に戻す、判定の結果は履歴に入れ古い応答は捨てる、選択欄で Delete しても箱を消さない、ポップアップは実際の高さで収める")

    # ---- 区分のキー: 追加した行のキーは保存し直しても変わらず、消したキーは使い回さない ----
    with app.app_context():
        v0 = app_module.vocab_public()
    row = lambda label: {"key": "", "label": label, "words": label, "exclude": "", "icon": "bi-file-earmark", "shape": "doc"}
    r1 = client.post("/admin/map-vocab", json={"vocab": dict(v0, media=v0["media"] + [row("図面")])}).get_json()["vocab"]
    k_zumen = next(x["key"] for x in r1["media"] if x["label"] == "図面")
    r2 = client.post("/admin/map-vocab", json={"vocab": dict(r1, media=[row("動画")] + r1["media"])}).get_json()["vocab"]
    keys2 = {x["label"]: x["key"] for x in r2["media"]}
    r3 = client.post("/admin/map-vocab", json={"vocab": dict(r2, media=[x for x in r2["media"] if x["label"] not in ("図面", "動画")] + [row("音声")])}).get_json()["vocab"]
    keys3 = {x["label"]: x["key"] for x in r3["media"]}
    check(keys2["図面"] == k_zumen and keys2["動画"] != k_zumen and keys3["音声"] not in (k_zumen, keys2["動画"]) and "図面" not in keys3,
          f"区分のキー: 並べ替え・追加でも変わらず、消したキーは使い回さない: {k_zumen} {keys2.get('動画')} {keys3.get('音声')}")
    client.post("/admin/map-vocab", json={"reset": True})
    r4 = client.post("/admin/map-vocab", json={"vocab": dict(v0, media=v0["media"] + [row("写真台帳")])}).get_json()["vocab"]
    check(next(x["key"] for x in r4["media"] if x["label"] == "写真台帳") not in (k_zumen, keys2["動画"], keys3["音声"]), "既定に戻しても、前に使ったキーは使い回さない")
    client.post("/admin/map-vocab", json={"reset": True})

    with app.app_context():
        # ---- 氏名・方法・手順図（実装レビューの指摘） ----
        check([app_module.method_of(t) for t in ("人に聞く", "紙で届く")] == ["verbal", "hand"], "受け取り方のチップ（人に聞く・紙で届く）も方法になる")
        check(app_module.mask_person_names("製品仕様") == "製品仕様" and app_module.mask_person_names("生産管理課の渡辺さん") == "生産管理課の担当者"
              and app_module.mask_person_names("山田さん（生産管理課）") == "担当者（生産管理課）", "仕様は氏名にしない、部署の後ろ・かっこの前の氏名だけ担当者")
        w = app_module.who_of("製造１課 課長", "製造1課")
        check((w["actor"], w["label"]) == ("boss", "製造1課 課長"), f"職場マスタは全角の数字でも同じ職場: {w}")
        d1 = app_module.build_steps_flow([{"no": 1, "action": "生産管理システムから日次表を出力する", "output": "日次表"},
                                          {"no": 2, "action": "週報を課長に渡す", "input": "週報", "to": "課長"}], "製造1課")
        into_boss = [d1.nodes[a]["text"] for (a, b), e in d1.edges.items() if e["type"] == "info" and d1.nodes[b].get("who") == "製造1課 課長"]
        check(into_boss == ["週報"], f"渡すものは手順の input（前の手順のできるものではない）: {into_boss}")
        d2 = app_module.build_steps_flow([{"no": 1, "action": "日次実績集計表から1週間分の数字を週報に転記する", "system": "Excel（週報）", "input": "日次実績集計表"}])
        check(any(d2.nodes[a]["text"] == "日次実績集計表" and b == "st1" for (a, b) in d2.edges) and any(a == "st1" and d2.nodes[b]["text"] == "Excel（週報）" for (a, b) in d2.edges),
              "書き込み先のシステムがある手順でも、input を取り出し元として描く")
        d3 = app_module.build_steps_flow([{"no": 1, "action": "実績を生産管理システムに登録する", "to": "生産管理システム"}])
        check(not any(n["text"] == "（渡すもの未確認）" for n in d3.nodes.values()), "相手がシステムだけなら「渡すもの未確認」を作らない")

        # ---- 手修正の重ね合わせ・昔の形の読み替え ----
        auto = {"nodes": [{"id": "o8", "text": "段取り時間推移グラフ", "type": "info", "cls": "output", "medium": "unknown", "source": "auto", "hint": None},
                          {"id": "p_x", "text": "使いみち未確認", "type": "work", "cls": "work", "who": "他部署", "actor": "dept", "who_state": "ok", "source": "auto"}],
                "edges": [{"id": "o8->p_x", "from": "o8", "to": "p_x", "type": "info", "method": "mail", "show_method": True, "handoff": "", "inferred": False, "label": "", "source": "auto"}]}
        saved = {"schema": 2, "nodes": [dict(auto["nodes"][0], text="段取り時間推移グラフ（Excel版）", text_edited=True),
                                        dict(auto["nodes"][1], actor="external", who_edited=True, actor_locked=True)],
                 "edges": [dict(auto["edges"][0], method="u99", method_edited=True)]}
        mg = app_module.merge_graph(auto, app_module.sanitize_graph(saved))
        nm, em = {n["id"]: n for n in mg["nodes"]}, mg["edges"][0]
        check(nm["o8"]["medium"] == "excel" and nm["p_x"]["actor"] == "external" and nm["p_x"].get("actor_locked") and em["method"] == "mail" and not em.get("method_edited") and em["auto_method"] == "mail",
              f"手修正: 名前を直した情報は新しい名前で媒体を判定、手で選んだだれ区分は残る、消された方法のキーは自動に戻る: {nm['o8']['medium']} {nm['p_x']['actor']} {em['method']}")
        legacy = {"nodes": [{"id": "p_05", "text": "生産管理課", "cls": "ext", "source": "auto"}, {"id": "m_k", "text": "kintone", "cls": "system", "source": "manual"}],
                  "edges": [{"id": "e1", "from": "m_k", "to": "t1", "kind": "manual", "source": "manual", "label": "毎朝8時までにメールで届く（遅れたら電話）", "label_edited": True},
                            {"id": "e2", "from": "m_k", "to": "t2", "kind": "manual", "source": "manual", "label": "メール"}]}
        up = app_module.upgrade_graph(legacy)
        un, ue = {n["id"]: n for n in up["nodes"]}, {e["id"]: e for e in up["edges"]}
        check(un["p_05"]["who"] == "生産管理課" and un["p_05"]["actor"] == "dept" and un["p_05"]["text"] == "使いみち未確認" and un["m_k"].get("medium_edited")
              and ue["e1"]["label"].startswith("毎朝") and not ue["e1"].get("method_edited") and ue["e2"]["method"] == "mail" and ue["e2"]["label"] == "",
              "昔の形: 人・部署の箱はだれのある作業に、手で選んだシステムは残り、文のラベルはそのまま・方法の名前だけのラベルは方法にする")

    # ---- 聞き取りの答えと業務マップの連動 ----
    with app.test_request_context():
        tasks = Task.query.all()
        by_name = {t.name: t for t in tasks}
        t_day, t_week, t_asm = by_name["日報の実績をExcelの集計表に転記する"], by_name["週報の作成"], by_name["組立実績の日次入力（生産管理システム）"]
        flow = app_module.build_info_flow(tasks)
        nd = flow.nodes
        into = lambda t: [(nd[a]["text"], nd[a]["type"], e["type"], e["inferred"]) for (a, b), e in flow.edges.items() if b == f"t{t.id}"]
        out_of = lambda t: [(nd[b].get("who") or nd[b]["text"], nd[b]["type"], app_module.edge_label(e)) for (a, b), e in flow.edges.items() if a in (f"t{t.id}", f"o{t.id}")]
        check(("紙の日報", "info", "info", False) in into(t_day) and not any(x[0] == "各ラインの手書き日報" for x in into(t_day)),
              f"取り出し元は手順で答えた名前（紙の日報）の情報の箱で、同じものは1つ: {into(t_day)}")
        check((t_day.output, "info", "info", True) in into(t_week) and not any(x[0] == "Excel（週報）" for x in into(t_week)),
              f"別の業務のできるもの（日次実績集計表）を使うと、点線（推定）の情報の流れで自動につながる: {into(t_week)}")
        check(("生産管理課", "work", "") in out_of(t_asm) and ("生産管理システム", "info", "") in out_of(t_asm) and not any(x[0] == "他部署" for x in out_of(t_asm)),
              f"渡す相手は相手の作業（だれ＝生産管理課）、入力するシステムは書き込む情報の箱: {out_of(t_asm)}")
        check(("製造1課 課長", "work", "メール") in out_of(t_week), f"渡し方（メール）を情報の流れのラベルに書き、課長は職場名つき: {out_of(t_week)}")
        types = {k: n["type"] for k, n in nd.items()}
        bad = [(a, b) for (a, b), e in flow.edges.items() if (e["type"] == "work") != (types[a] == "work" and types[b] == "work") or (types[a] == "info" and types[b] == "info")]
        check(not bad and set(n["type"] for n in nd.values()) == {"work", "info"} and set(e["type"] for e in flow.edges.values()) <= {"work", "info"}, f"箱は作業と情報、線は仕事の流れと情報の流れだけ（規則違反 {len(bad)} 本）")
        labels = {app_module.edge_label(e) for e in flow.edges.values()}
        check(not labels & {"他のExcel", "紙の帳票", "人から聞く", "その他", "未設定", "次の業務", "前の業務"}, f"区分名や「未設定」を線のラベルにしない: {labels}")
        steps = app_module.build_task_flow(t_day)
        st = {(steps.nodes[a].get("who") if steps.nodes[a]["cls"] == "work" else steps.nodes[a]["text"], steps.nodes[b].get("who") if steps.nodes[b]["cls"] == "work" else steps.nodes[b]["text"], e["type"], app_module.edge_label(e)) for (a, b), e in steps.edges.items()}
        check(("紙の日報", "① 各ラインから手書きの日報を集める", "info", "") in st and ("紙の日報", "② 日報の数字を日次実績集計表に転記する", "info", "") in st
              and ("② 日報の数字を日次実績集計表に転記する", "Excel（日次実績集計表）", "info", "") in st and ("④ 朝会用に印刷して課長に渡す", "製造1課 課長", "work", "前が伝えに行く") in st
              and ("印刷した集計表", "製造1課 課長", "info", "") in st and ("① 各ラインから手書きの日報を集める", "② 日報の数字を日次実績集計表に転記する", "work", "") in st,
              f"手順図: 帳票 → 手順、手順 → 転記先、渡すもの → 相手の作業（情報の流れ）と 渡す手順 → 相手の作業（仕事の流れ・前が伝えに行く）")
        recip = next(n for n in steps.nodes.values() if n["cls"] == "work" and n["who"] == "製造1課 課長")
        check(recip["text"] == t_day.purpose, "相手の作業の内容は使いみち（相手が1つなら業務の目的）")
        steps = app_module.build_task_flow(t_asm)
        st = {(steps.nodes[a].get("who") if steps.nodes[a]["cls"] == "work" else steps.nodes[a]["text"], steps.nodes[b].get("who") if steps.nodes[b]["cls"] == "work" else steps.nodes[b]["text"], e["type"]) for (a, b), e in steps.edges.items()}
        check(("② 生産管理システムに品番ごとの実績を入力する", "生産管理システム", "info") in st and ("システム上の実績データ", "生産管理課", "info") in st,
              "手順図: 入力するシステムは手順から情報の流れ、渡す相手の部署は相手の作業")
        rows = app_module.source_rows(tasks)
        check(any(r["name"] == "紙の日報" for r in rows), "システム・情報源の表も同じ名前で数える")
        wb_saved = {"schema": 2, "nodes": [{"id": f"t{t_day.id}", "text": t_day.name, "type": "work", "cls": "task", "source": "auto"},
                                           {"id": "u_x", "text": "出勤簿の転記（2人）", "type": "work", "cls": "pending", "who": "製造1課", "who_state": "ok", "actor": "team", "source": "auto"},
                                           {"id": "p_unknown", "text": "使う人が分からない", "type": "work", "cls": "warn", "source": "auto"},
                                           {"id": "p_y", "text": "使いみち未確認", "type": "work", "cls": "work", "who": "生産管理課", "actor": "dept", "who_state": "ok", "source": "auto"}],
                    "edges": [{"id": "a", "from": f"t{t_day.id}", "to": "u_x", "type": "work", "source": "manual"},
                              {"id": "b", "from": f"t{t_day.id}", "to": "p_unknown", "type": "info", "source": "manual"},
                              {"id": "c", "from": "p_y", "to": f"t{t_day.id}", "type": "work", "source": "manual"}]}
        wc = app_module.writeback_candidates(wb_saved, {t.id: t for t in tasks})
        got = sorted(op["field"] + ":" + op["value"] for c in wc for op in c["ops"])
        check(got == ["actors:生産管理課", "next_task:出勤簿の転記"], f"書き戻し: 未登録の業務は後の業務（人数は外す）、だれのある箱から業務へは関わる人、分からない箱は出さない: {got}")
        src_def = next(s_ for s_ in app_module.task_sources(by_name["不良品発生報告のメール配信"]) if "不良報告書" in s_["name"])
        check(src_def["method"] != "verbal", f"取り出し元の受け取り方は、名前に合う区分から取る（紙の報告書に「口頭」を付けない）: {src_def}")
    with app.app_context():
        check(app_module.users_from_names(["課長", "生産管理課", "顧客", "山田さん", "生産管理システム"]) == ["上司", "他部署", "顧客・社外", "職場内"],
              "渡す相手の名前から利用者の区分を決める（AI が使えないときの予備）")
        state = app_module._clean_state({"steps": [{"action": "集計する", "to": "課長", "how": "メール", "use": "朝会で確認"}], "recipients": ["品質保証課"], "purpose": "朝会で前日の実績を確認する"})
        tmp = Task(name="予備判定の確認")
        app_module._fill_from_state(tmp, state)
        check(state["purpose"] == "朝会で前日の実績を確認する" and state["steps"][0]["how"] == "メール" and state["steps"][0]["use"] == "朝会で確認" and tmp.purpose == "朝会で前日の実績を確認する" and tmp.output_users_list == ["他部署", "上司"],
              f"手順の渡し方・使いみちを受け取り、使いみちは目的に、渡す相手（品質保証課・課長）は他部署・上司になる: {tmp.output_users_list}")
        db.session.rollback()
        r0 = Respondent.query.filter(Respondent.emp_id == "DEMO-001").first()
        handlers = app_module.intake_handlers(r0)
        known = handlers["list_known_systems"]({})
        ph = handlers["person_history"]({})
        check(any(x["種類"] == "できるもの" for x in known["候補"]) and "氏名" not in ph and all("担当" not in app_module._task_brief(t) for t in Task.query.limit(3)),
              "名寄せの道具はできるものの名前も返し、AI には氏名を渡さない")
        db.session.add(MonthlyResponse(period=current_period(), emp_id="P100", name="", workplace_id=t_day.workplace_id,
                                       answers_json=json.dumps({"has": "ある", "task_names": "検査記録の月次まとめ\n週報の作成"}, ensure_ascii=False)))
        db.session.commit()
    g = client.get(f"/bizmap/api/graph?kind=info&workplace_id={t_day.workplace_id}").get_json()
    pend = [n["text"] for n in g["nodes"] if n["cls"] == "pending"]
    check(any(p.startswith("検査記録の月次まとめ") for p in pend) and not any(p.startswith("週報の作成") for p in pend) and all(n["type"] == "work" for n in g["nodes"] if n["cls"] == "pending"),
          f"月次で挙がったが未登録の業務は灰色の作業の箱（登録済みの業務は出さない）: {pend}")
    html = client.get(f"/bizmap/?workplace_id={t_day.workplace_id}").data.decode("utf-8")
    check("月次で挙がった未登録の業務" in html and "lg-pending" in html, "業務マップに未登録の業務の件数と凡例")
    g = client.get(f"/bizmap/api/graph?kind=info&domain_id=1").get_json()
    check(not any(n["cls"] == "pending" for n in g["nodes"]), "業務領域で絞ったときは未登録の業務を出さない（領域が分からないため）")
    with app.app_context():
        MonthlyResponse.query.filter_by(emp_id="P100").delete()
        db.session.commit()

    with app.app_context():
        auto = {"nodes": [{"id": "t1", "text": "業務", "cls": "task", "type": "work", "source": "auto"}, {"id": "s_new", "text": "紙の日報", "cls": "source", "type": "info", "medium": "paper", "source": "auto"}], "edges": []}
        saved = {"nodes": [{"id": "t1", "text": "業務", "cls": "task", "x": 1, "y": 1}, {"id": "s_old", "text": "各ラインの手書き日報", "cls": "source", "source": "auto", "x": 5, "y": 5},
                           {"id": "t99", "text": "消えた業務", "cls": "task", "source": "auto"}, {"id": "p_x", "text": "品質保証課", "cls": "ext", "source": "auto", "text_edited": True},
                           {"id": "m_x", "text": "Excelの台帳", "cls": "manual", "source": "manual"}, {"id": "m_y", "text": "紙の日報", "cls": "manual", "source": "manual"}],
                 "edges": [{"id": "s_old->t1", "from": "s_old", "to": "t1", "source": "manual", "kind": "manual"}, {"id": "t99->t1", "from": "t99", "to": "t1", "source": "manual", "kind": "manual"},
                           {"id": "m_x->t1", "from": "m_x", "to": "t1", "source": "manual", "kind": "manual", "label": "メール"}, {"id": "m_y->t1", "from": "m_y", "to": "t1", "source": "manual", "kind": "manual"}]}
        up = app_module.upgrade_graph(saved)
        check(up["schema"] == 2 and {n["id"]: (n["type"], n["cls"]) for n in up["nodes"]}["p_x"] == ("work", "work") and next(n for n in up["nodes"] if n["id"] == "p_x").get("who") == "品質保証課",
              "保存済みの図は読み込むときに新しい形へ読み替える（人の箱は作業、名前を直していればだれへ）")
        mg = app_module.merge_graph(auto, saved)
        ids = {n["id"] for n in mg["nodes"]}
        mx = next(e for e in mg["edges"] if e["from"] == "m_x")
        check({"t1", "s_new", "s_old", "m_x"} <= ids and "t99" not in ids and "m_y" not in ids and any(e["from"] == "s_new" and e["to"] == "t1" for e in mg["edges"])
              and mx["type"] == "info" and mx.get("method") == "mail" and mx.get("show_method") and not mx.get("label"),
              "手で引いた線の相手は残し、同じ名前の手で足した箱は自動の箱にまとめ、ラベルの方法名は線の方法に読み替える")

    r = client.get("/bizmap/api/graph?kind=info")
    g = r.get_json()
    check(r.status_code == 200 and g["needs_layout"] is True, "GET graph (auto) needs_layout")
    check(len(g["nodes"]) > 20 and len(g["edges"]) > 20, f"auto graph nodes={len(g['nodes'])} edges={len(g['edges'])}")
    task_node = next(n for n in g["nodes"] if n["id"].startswith("t"))
    some_edge = g["edges"][0]
    check(task_node.get("url") and task_node["cls"] in ("task", "done"), "task node has url/cls")

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

    with app.app_context():
        e001 = Employee.query.filter_by(emp_id="DEMO-001").first()
        wp_e001, pat_e001 = e001.workplace_id, e001.work_pattern or "普通勤務"
    client.post("/survey/api/respondent", json={"emp_id": "DEMO-001", "name": "", "workplace_id": wp_e001, "work_pattern": pat_e001})
    client.post("/survey/api/tasks", json={
        "name": "新規に登録した業務", "frequency": "毎週", "duration": "15分未満", "wish": "減らしたい", "self_improve": False, "messages": [],
        "state": {"title": "新規に登録した業務", "steps": [{"no": 1, "action": "メールで届いた数字を帳票に転記する", "system": "メール", "input": "", "output": "新しい帳票", "to": "上司"}], "sources": ["メール"], "outputs": ["新しい帳票"], "recipients": ["上司"]},
    })
    with app.app_context():
        new_task = Task.query.filter_by(name="新規に登録した業務").first()
    r = client.get("/bizmap/api/graph?kind=info")
    m2 = r.get_json()
    check(f"t{new_task.id}" in m2["new_node_ids"], "merged: new task appears in new_node_ids")
    check(f"t{new_task.id}" in {n["id"] for n in m2["nodes"]}, "merged: new task node included")
    check("m_abc123" in {n["id"] for n in m2["nodes"]}, "merged: manual node still kept after new task")

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

    r = client.delete("/bizmap/api/graph?kind=info")
    check(r.status_code == 200, "DELETE reset")
    m3 = client.get("/bizmap/api/graph?kind=info").get_json()
    check(m3["needs_layout"] is True and hidden_node["id"] in {n["id"] for n in m3["nodes"]}, "after reset: auto graph again")

    r = client.post("/bizmap/api/graph", json={"kind": "info", "data": "oops"})
    check(r.status_code == 400, "POST invalid data -> 400")
    r = client.get("/bizmap/api/graph?kind=nope")
    check(r.status_code == 404, "GET unknown kind -> 404")

    return fails

def smoke_board():
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

    code, html = page("/board/")
    check(code == 200 and "進行中の案件はありません" in html, "空の掲示板")
    code, html = page("/")
    check("アンケート・業務登録" not in html and 'href="/monthly/">月次アンケート</a>・<a href="/survey/">自動化を検討</a>から始める' in html and "デモデータを表示中" not in html,
          "ダッシュボードの空の案内は「月次アンケート・自動化を検討」。デモがなければ表示中の帯は出ない")
    code, html = page("/monthly/admin")
    check("自動化を検討へ誘導" in html or json.dumps("自動化を検討へ誘導")[1:-1] in html, "設問エディタの誘導の名前は「自動化を検討へ誘導」")

    import app as app_module
    with app.app_context():
        own_survey = MonthlySurvey.query.filter_by(period=current_period()).first().id
        wp_del = app_module.Workplace.query.filter_by(name="設備保全課").first().id
        dom_del = app_module.Domain.query.filter_by(name="人員・労務管理").first().id
    client.post(f"/admin/workplace/{wp_del}/delete")
    client.post(f"/admin/domain/{dom_del}/delete")
    client.post("/admin/demo/load")
    with app.app_context():
        made = json.loads(app_module.Setting.get("demo_masters", "{}"))
        check(len(made.get("workplaces", [])) == 1 and len(made.get("domains", [])) == 1 and len(made.get("surveys", [])) == 3 and own_survey not in made["surveys"],
              f"デモが作った職場・業務領域・月次の設問セットを控える（前からあった当月の設問セットは含めない）: {made}")
        demo = Task.query.filter_by(is_demo=True).all()
        by = {t.name: t for t in demo}
        check(len(demo) == 26 and all(t.steps for t in demo) and all(t.ai_source == "rules" for t in demo) and all(t.wish in ("なくしたい", "減らしたい", "まずは相談したい") for t in demo),
              "デモ: 全業務に手順があり、本番の登録と同じくルールで項目を埋め、希望は3つの選択肢")
        check(by["安全パトロール結果の入力"].wish == "減らしたい" and by["作業手順書の改訂管理台帳の更新"].wish == "なくしたい" and by["設備日常点検表のファイリング確認"].wish == "まずは相談したい" and not by["設備日常点検表のファイリング確認"].is_candidate,
              "デモ: 以前の「このままでよい」は内容に合う希望にし、相談だけで困りごとのない業務は案件候補にしない")
        asm = by["組立実績の日次入力（生産管理システム）"]
        claim = by["クレーム対応の進捗報告"]
        check(asm.inputs_list == ["紙の帳票"] and "転記・コピー" not in claim.nature_list and claim.nature_list and by["日報の実績をExcelの集計表に転記する"].pains_list == ["時間がかかる", "同じ情報を何度も入力する"],
              f"デモ: 入力元は取り出し元だけ（入力先のシステムは入れない）、作業の性質は手順の言葉から、困りごとは言葉から: {asm.inputs_list} {claim.nature_list}")
        irregular = by["作業手順書の改訂管理台帳の更新"]
        check(irregular.frequency == "不定期" and irregular.monthly_minutes == 22.0, f"不定期は月1回として計算する（{irregular.monthly_minutes} 分）")
        emps = sorted(e.emp_id for e in app_module.Employee.query.filter_by(is_demo=True).all())
        check(emps == [f"DEMO-{i:03d}" for i in range(1, 13)] and Respondent.query.filter_by(emp_id="DEMO-001", is_demo=True).first() is not None, f"デモの社員IDは DEMO-001〜DEMO-012: {emps[:2]}…")
        irregular_id = irregular.id
    code, html = page(f"/cases/{irregular_id}")
    check("不定期（月1回として計算）" in html, "案件詳細: 不定期は「月1回として計算」と出す")

    code, html = page("/board/")
    check(code == 200 and "デモデータを表示中" in html, "掲示板 200・デモデータを表示中の帯")
    code, html = page("/cases/")
    check("デモデータを表示中" in html and html.count('text-warning-emphasis">デモ</span>') == 26, "改善案件の一覧: 帯とデモの印")
    code, html = page("/")
    check("デモデータを表示中" in html and 'text-warning-emphasis">デモ</span>' in html, "ダッシュボード: 帯と最近の業務にデモの印")
    client.post("/logout")
    code, html = page("/board/")
    check("デモデータを表示中" in html and 'href="/admin/"' not in html, "未ログインの活動状況にも帯が出る（管理へのリンクはない）")
    client.post("/login", data={"emp_id": "admin", "password": "adminpass"})
    code, html = page("/board/")
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
    check("自動化を検討してみませんか" in html and "自動化を検討する</a>" in html and 'href="/survey/"' in html and 'href="/monthly/"' in html, "掲示板: 自動化依頼と月次アンケートへの導線")

    with app.app_context():
        t_test = Task.query.filter_by(name="検査成績書の作成").first()
        wp_id = t_test.workplace_id
        other = Task.query.filter(Task.workplace_id != wp_id, Task.status == "対応中").first()
    code, html = page(f"/board/?workplace_id={wp_id}")
    check(code == 200 and "検査成績書の作成" in html and (other is None or other.name not in html), "職場で絞り込める")

    code, html = page("/board/?kiosk=1")
    check(code == 200 and "pb-kiosk" not in html and "navbar" in html, "掲示モードは出ない")
    code, html = page("/")
    check("<footer" not in html, "ダッシュボードにもフッターなし")

    r = client.post("/admin/settings", data={"project_name": "P"})
    code, html = page("/board/")
    check("pb-name" not in html and "伊藤 大輔" not in html, "設定: 依頼者名が消える")
    client.post("/admin/settings", data={"project_name": "P", "board_show_names": "1"})
    code, html = page("/board/")
    check("pb-name" in html and "伊藤 大輔" in html, "設定: 氏名を再表示")

    code, html = page("/")
    check(code == 200 and "活動状況で案件の進み具合を見る" in html and "pb-funnel" not in html and "職場別の登録件数と月間時間" not in html and "業務領域別の把握状況" not in html, "ダッシュボード: 重複を外してリンクに")
    check("職場別の回答状況" in html and "最近登録された業務" in html and "次のアクション提案" in html, "ダッシュボード: 残した区画")
    code, html = page("/survey/")
    check(code == 200 and "自動化を検討" in html and "自動化依頼" not in html and 'id="rq"' in html, "自動化を検討は1画面（名称は「自動化を検討」）")
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

    r = client.post(f"/cases/{tid}/stage", data={"direction": "next"})
    check(r.status_code == 302, "次へ 302")
    with app.app_context():
        t = db.session.get(Task, tid)
        check(t.stage == "spec" and t.status == "対応中", f"次へ: spec/対応中 ({t.stage}/{t.status})")
        check(len(t.stage_logs) == 2 and t.stage_logs[-1].stage == "spec", "次へ: 履歴が増える")

    client.post(f"/cases/{tid}/stage", data={"direction": "prev"})
    with app.app_context():
        t = db.session.get(Task, tid)
        check(t.stage == "request" and len(t.stage_logs) == 3, "戻す: request・履歴3件")

    client.post(f"/cases/{tid}/stage", data={"stage": "done", "progress_note": "運用開始しました"})
    with app.app_context():
        t = db.session.get(Task, tid)
        check(t.stage == "done" and t.status == "完了" and t.completed_at is not None and t.progress_note == "運用開始しました", "直接指定: 完了に同期")

    form = {"decision": "未定", "assignment": "未定", "assignee": "", "after_minutes": "", "result_note": "", "progress_note": "活動状況のひとこと"}
    client.post(f"/cases/{tid}/update", data=form)
    with app.app_context():
        t = db.session.get(Task, tid)
        check(t.stage == "done" and t.status == "完了" and t.progress_note == "活動状況のひとこと", "フォーム更新では段階・ステータスは変わらない")

    client.post(f"/cases/{tid}/hold")
    with app.app_context():
        t = db.session.get(Task, tid)
        check(t.status == "完了", "完了は保留にできない")

    client.post(f"/cases/{tid}/stage", data={"stage": "test"})
    client.post(f"/cases/{tid}/hold")
    with app.app_context():
        t = db.session.get(Task, tid)
        check(t.stage == "test" and t.status == "保留" and t.completed_at is None, f"保留で段階を保つ ({t.stage}/{t.status})")
    code, html = page("/board/")
    check("保留中" in html, "掲示板に保留中セクション")
    code, html = page(f"/cases/{tid}")
    check("保留を解除" in html, "案件詳細に保留解除ボタン")

    client.post(f"/cases/{tid}/hold")
    with app.app_context():
        t = db.session.get(Task, tid)
        check(t.stage == "test" and t.status == "対応中", f"保留解除で対応中 ({t.stage}/{t.status})")

    client.post(f"/cases/{tid}/hold")
    client.post(f"/cases/{tid}/stage", data={"direction": "next"})
    with app.app_context():
        t = db.session.get(Task, tid)
        check(t.stage == "check" and t.status == "対応中", f"保留から次へで再開 ({t.stage}/{t.status})")

    client.post(f"/cases/{tid}/stage", data={"stage": "start"})
    with app.app_context():
        t = db.session.get(Task, tid)
        check(t.stage == "start" and t.status == "受付", f"開始に戻すと受付 ({t.stage}/{t.status})")
    client.post(f"/cases/{tid}/stage", data={"stage": "check"})

    with app.app_context():
        n_before = StageLog.query.filter_by(task_id=tid).count()
    client.post(f"/cases/{tid}/stage", data={"stage": "check"})
    with app.app_context():
        check(StageLog.query.filter_by(task_id=tid).count() == n_before, "同じ段階なら履歴は増えない")

    code, html = page(f"/cases/{tid}")
    with app.app_context():
        steps_now = db.session.get(Task, tid).steps_done_list
    check("段階は「確認」まで進んでいます" in html and "1. アウトプットの確認・2. 目的の明確化・3. ECRSによる最適化 にまだチェックがありません" in html and steps_now == [],
          "段階が要件定義より先なのに標準手順のチェックが遅れていれば、案件詳細で更新を促す（自動ではチェックしない）")
    client.post(f"/cases/{tid}/update", data=dict(form, steps=["1", "2", "3"]))
    code, html = page(f"/cases/{tid}")
    check("にまだチェックがありません" not in html, "標準手順にチェックすれば案内は消える")

    with app.app_context():
        wk = Task.query.filter_by(name="週報の作成").first()
        wk_id, wk_minutes, wk_marked = wk.id, wk.monthly_minutes, "domain" in wk.ai_fields_list
        dom_q = app_module.Domain.query.filter_by(name="品質管理").first().id
    client.post(f"/cases/{wk_id}/update", data=dict(form, prev_task="日報の実績をExcelの集計表に転記する", next_task="月報の作成（佐藤さんへ）", actors="佐藤課長、生産管理課", domain_id=str(dom_q), headcount="3"))
    with app.app_context():
        wk = db.session.get(Task, wk_id)
        check(wk_marked and wk.prev_task == "日報の実績をExcelの集計表に転記する" and wk.next_task == "月報の作成（担当者へ）" and wk.actors == "課長、生産管理課" and wk.domain_id == dom_q and "domain" not in wk.ai_fields_list
              and wk.headcount == 3 and wk.monthly_minutes == wk_minutes * 3,
              "案件詳細で前の業務・後の業務・関わる人・業務領域・担当人数を直せる（氏名は置き換え、直した領域は推定の印を外し、人数は月間時間に掛ける）")
    client.post(f"/cases/{wk_id}/update", data=dict(form, domain_id=""))
    code, html = page(f"/cases/{wk_id}")
    check('name="adopt_ai_domain"' in html and 'name="prev_task"' in html and 'name="next_task"' in html and 'name="actors"' in html and 'name="domain_id"' in html and 'name="headcount"' in html,
          "業務領域が空でも「AI推定にする」を選べ、更新フォームに業務情報の欄がある")
    client.post(f"/cases/{wk_id}/verify", data={"verified_by": "推進メンバー", "adopt_ai_domain": "1"})
    with app.app_context():
        wk = db.session.get(Task, wk_id)
        check(wk.verified and wk.domain is not None and wk.domain.name == wk.ai_domain, f"空の業務領域を AI推定（{wk.ai_domain}）にして確認できる")
        eq = app_module.Domain.query.filter_by(name="設備管理").first()
        eq_id, eq_desc = eq.id, eq.description
        n_eq = Task.query.filter_by(ai_domain="設備管理").count()
        eq_task = Task.query.filter_by(ai_domain="設備管理").first().id
    client.post(f"/admin/domain/{eq_id}/update", data={"name": "設備保全", "description": eq_desc, "sort": "3"})
    with app.app_context():
        guess = app_module.rule_domain(Task(name="設備の点検記録を台帳に入力する"), app_module._domains())
        check(n_eq > 0 and Task.query.filter_by(ai_domain="設備管理").count() == 0 and Task.query.filter_by(ai_domain="設備保全").count() == n_eq and guess is not None and guess.name == "設備保全",
              f"業務領域の名前を変えると AI推定の領域名も新しい名前になり、ルールの推定も新しい名前で働く: {guess.name if guess else None}")
    code, html = page(f"/cases/{eq_task}")
    check("AI推定: 設備管理" not in html and "設備保全" in html, "名前を変えたあと、案件詳細に古い領域名を AI推定として出さない")
    r = client.post(f"/cases/{eq_task}/reanalyze", follow_redirects=True)
    with app.app_context():
        t_eq = db.session.get(Task, eq_task)
        check("AI分析を更新しました（ルールベース）" in r.data.decode("utf-8") and t_eq.ai_source == "rules" and t_eq.ai_domain == "設備保全", "AI なしの再分析は登録と同じルールで行い、業務領域の推定も新しい名前で更新する")

    with app.app_context():
        new_task = Task.query.filter_by(name="日報の実績をExcelの集計表に転記する").first()
        check(new_task.stage_dates["start"] == new_task.created_at, "開始の日付は登録日")

    r = client.get("/cases/export.csv")
    csv = r.data.decode("utf-8-sig")
    check("進捗段階" in csv and "段階更新日" in csv and "テスト" in csv, "CSVに進捗段階")

    with app.app_context():
        before = StageLog.query.count()
    client.post("/admin/demo/clear")
    with app.app_context():
        check(before > 0 and StageLog.query.count() == 0 and Task.query.count() == 0, f"デモ削除で履歴も消える（{before} -> 0）")
        check(app_module.Workplace.query.filter_by(name="設備保全課").first() is None and app_module.Domain.query.filter_by(name="人員・労務管理").first() is None
              and MonthlySurvey.query.count() == 1 and db.session.get(MonthlySurvey, own_survey) is not None and app_module.Setting.get("demo_masters", "") == "",
              "デモ削除で、デモが作ってほかに使われていない職場・業務領域・設問セットも消す（前からあった設問セットは残す）")
    code, html = page("/board/")
    check(code == 200 and "デモデータを表示中" not in html, "削除後も掲示板 200（帯は消える）")

    client.post("/admin/demo/load")
    with app.app_context():
        e1 = app_module.Employee.query.filter_by(emp_id="DEMO-001").first()
        wp1, pat1 = e1.workplace_id, e1.work_pattern
    client.post("/survey/api/respondent", json={"emp_id": "DEMO-001", "workplace_id": wp1, "work_pattern": pat1})
    r = client.post("/survey/api/tasks", json={"name": "本物の業務", "state": {"title": "本物の業務", "steps": [{"no": 1, "action": "紙の日報を集める", "system": "紙の日報"}], "sources": ["紙の日報"]},
                                               "messages": [], "frequency": "毎週", "duration": "15分未満", "wish": "減らしたい", "shift": "昼勤"})
    check(r.status_code == 200, "デモの社員IDでも本物の業務を登録できる")
    r = client.post("/admin/demo/load")
    with app.app_context():
        real = Task.query.filter_by(name="本物の業務").first()
        check(r.status_code == 302 and real is not None and real.respondent is not None and real.respondent.emp_id == "DEMO-001" and Respondent.query.filter_by(emp_id="DEMO-001").count() == 1
              and app_module.Employee.query.filter_by(emp_id="DEMO-001").count() == 1 and Task.query.filter_by(is_demo=True).count() == 26,
              "本物の業務がある社員IDのままデモを入れ直しても止まらない（同じIDの人は使い回す）")
    client.post("/admin/demo/clear")
    with app.app_context():
        real = Task.query.filter_by(name="本物の業務").first()
        check(real is not None and real.respondent is not None and real.respondent.emp_id == "DEMO-001" and not real.respondent.is_demo and Task.query.count() == 1 and Respondent.query.count() == 1
              and app_module.Employee.query.filter_by(emp_id="DEMO-001").first() is not None and app_module.Employee.query.filter_by(emp_id="DEMO-002").first() is None,
              "デモ削除: 本物の業務がある回答者は残してデモの印を外し、業務を付け替えない。ほかのデモの人は消す")

    return fails

def smoke_errors():
    tmpdir = tempfile.mkdtemp()
    dbfile = os.path.join(tmpdir, "test.sqlite3").replace("\\", "/")
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

    code, html, r = get("/no-such-page")
    check(code == 404 and "ページが見つかりません" in html and "app-header" in html, "存在しないURLは 404 の画面（ヘッダーつき）")
    code, html, r = get("/cases/999999")
    check(code == 302 and "/login" in r.headers["Location"], "ログインが必要なページは未ログインならログインへ")
    client.post("/login", data={"emp_id": "admin", "password": "adminpass"})
    code, html, r = get("/cases/999999")
    check(code == 404 and "ページが見つかりません" in html, "ログイン後、ない案件は 404 の画面")
    code, html, r = get("/no-such-page")
    check(code == 404 and "ページが見つかりません" in html, "ログイン後も存在しないURLは 404")

    code, html, r = get("/_boom")
    check(code == 500 and "エラーが発生しました" in html and "app-header" in html, "失敗したときは 500 の画面（ヘッダーつき）")
    code, html, r = get("/board/")
    check(code == 200, "500 のあとも次のリクエストは普通に動く")

    return fails

def smoke_login_monthly():
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
        check(n > 20 and sample.answers.get("has") in ("ある", "ない"), f"デモ: 回答 {n} 件（answers_json）")
    code, html, r = get("/monthly/results")
    check(code == 200 and "PCの定型・定期業務あり" in html and "設問ごとの推移" in html and "trend-table" in html and "コメント" in html and "勤務形態別の推移" in html and "交代勤務" in html, "集計: 推移・設問ごとの推移・勤務形態別の推移・月度の詳細")
    code, html, r = get("/")
    check("月次アンケート（時系列）" in html and "一緒に取り組みたい" in html, "ダッシュボードの月次カード")
    code, html, r = get("/monthly/export.csv")
    csv = r.data.decode("utf-8-sig")
    check(code == 200 and "PCの定型・定期業務（繰り返しの手動作業）はありますか" in csv and "コメント" in csv, "CSV に設問の文章が列名で出る")

    # ---- 回答（ログイン不要・1人1回・その場で完了） ----
    client.post("/logout")

    with app.app_context():
        for e in ("E013", "E014", "E015", "X101", "X102", "X104", "X105", "X106", "X107", "X108", "X109", "X110", "X111", "X998", "X999"):
            if Employee.query.filter_by(emp_id=e).first() is None:
                db.session.add(Employee(emp_id=e, name="", source="manual"))
        db.session.commit()
    period = current_period()
    code, html, r = get(f"/monthly/?period={period}&emp=DEMO-009")
    check(code == 200 and 'id="name"' not in html and "小林 直樹" not in html and "mq-fallback" not in html and "社員IDを確認しました" in html,
          "メールリンク: 社員IDだけで開く（氏名の欄も表示もない）")
    r = client.post("/monthly/submit", json={"period": period, "emp_id": "NOPE01", "workplace_id": "1", "work_pattern": "普通勤務", "answers": {}})
    check(r.status_code == 400 and "見つかりません" in r.get_json()["errors"].get("emp_id", ""), "存在しない社員IDは送信できない（氏名を入れて通すことはない）")
    j = client.get(f"/monthly/api/employee/NOPE01?period={period}").get_json()
    check(j["found"] is False, "照会 API は存在しないIDを「見つからない」と返す")
    j = client.get(f"/monthly/api/employee/%EF%BC%A4%EF%BC%A5%EF%BC%AD%EF%BC%AF%EF%BC%8D%EF%BC%90%EF%BC%90%EF%BC%99?period={period}").get_json()
    check(j["found"] and j["emp_id"] == "DEMO-009", "全角で入れた社員ID（ＤＥＭＯ－００９）も同じ人として扱う")
    with app.app_context():
        wp009 = str(Employee.query.filter_by(emp_id="DEMO-009").first().workplace_id)
    payload = {"period": period, "emp_id": "DEMO-009", "name": "小林 直樹", "workplace_id": wp009, "work_pattern": "普通勤務", "answers": {"has": "ある", "hours": "45分〜1時間", "task_names": "日報の集計", "trend": "減った", "decreased_reason": "日報の転記を自動化した", "wants": "ある", "wants_tasks": "日報の転記", "stance": "学びながら改善チームと一緒に取り組みたい", "new_tasks": ""}, "comment": "テストのコメント"}
    with app.app_context():
        MonthlyResponse.query.filter_by(period=period, emp_id="DEMO-009").delete()
        db.session.commit()
    r = client.post("/monthly/submit", json=payload)
    j = r.get_json()
    check(r.status_code == 200 and j["success"] and j["wants"] and "/survey/start" in j["register_url"], "JSON 送信で受付（本登録への誘導あり）")
    r = client.post("/monthly/submit", json=payload)
    check(r.status_code == 400 and "回答済み" in r.get_json()["message"], "2回目は回答済みで拒否（1人1回）")
    code, html, r = get(f"/monthly/?period={period}&emp=DEMO-009")
    check("回答済み" in html and 'id="mq-form" class="card mq-card d-none"' in html, "回答済みの人はフォームの代わりに回答済み表示")
    r = client.get(f"/monthly/api/employee/DEMO-009?period={period}")
    check(r.get_json()["answered"] is True, "照会 API が回答済みを返す")
    bad = dict(payload, emp_id="DEMO-010", name="加藤 恵", answers=dict(payload["answers"], has=""))
    r = client.post("/monthly/submit", json=bad)
    check(r.status_code == 400 and "has" in r.get_json()["errors"], "必須の選択が空ならエラー")
    bad = dict(payload, emp_id="DEMO-010", answers=dict(payload["answers"], has="へんな値"))
    r = client.post("/monthly/submit", json=bad)
    check(r.status_code == 400, "選択肢にない値は拒否")
    with app.app_context():
        MonthlyResponse.query.filter_by(period=period, emp_id="DEMO-010").delete()
        db.session.commit()
        e010 = Employee.query.filter_by(emp_id="DEMO-010").first()
        wp010, pat010 = str(e010.workplace_id), e010.work_pattern
    form = {"period": period, "emp_id": "DEMO-010", "name": "加藤 恵", "workplace_id": wp010, "work_pattern": pat010}
    for k, v in payload["answers"].items():
        form["ans-" + k] = v
    form["comment"] = ""
    r = client.post("/monthly/submit", data=form)
    check(r.status_code == 302 and "/monthly/done" in r.headers["Location"], "通常のフォーム送信でも受付")
    code, html, r = get(f"/monthly/done?period={period}")
    check(code == 200 and "回答を受け付けました" in html and "自動化を検討する" in html and "src=monthly" in html and "日報の転記" in html, "完了ページ（なくしたい業務があれば「自動化を検討する」へ）")

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
    with app.app_context():
        old = MonthlySurvey(period="2024-01", is_active=False)
        old.set_questions([q for q in MonthlySurvey.query.filter_by(period=period).first().questions if q["key"] != "stance"] + [
            {"key": "ask", "type": "choice", "text": "改善チームに頼みたいですか", "options": ["頼みたい", "相談したい", "自分で検討する", "特にない"], "required": True},
            {"key": "self", "type": "choice", "text": "自分で改善に取り組みたいですか", "options": ["取り組みたい", "教えてもらえれば取り組みたい", "今は難しい"], "required": True},
        ])
        db.session.add(old)
        db.session.add(MonthlyResponse(period="2024-01", emp_id="E001", name="田中 一郎", workplace_id=1, answers_json=json.dumps({"has": "ある", "hours": "5〜10時間", "trend": "初回なので分からない", "ask": "相談したい", "self": "今は難しい"}, ensure_ascii=False)))
        db.session.add(MonthlyResponse(period="2024-01", emp_id="E002", name="佐藤 花子", workplace_id=1, answers_json=json.dumps({"has": "ない", "hours": "0時間", "ask": "特にない", "self": "教えてもらえれば取り組みたい"}, ensure_ascii=False)))
        db.session.add(MonthlyResponse(period="2024-01", emp_id="E003", name="鈴木 健", workplace_id=1, answers_json=json.dumps({"has": "少しある", "wants": "ある", "stance": "教えてもらいながら自分で取り組みたい"}, ensure_ascii=False)))
        db.session.add(MonthlyResponse(period="2024-01", emp_id="E004", name="高橋 美咲", workplace_id=1, answers_json=json.dumps({"wants": "ない", "stance": "今は特にない"}, ensure_ascii=False)))
        db.session.add(MonthlyResponse(period="2024-01", emp_id="E005", name="伊藤 大輔", workplace_id=1, answers_json=json.dumps({"has": "ある", "wants": "ある", "stance": "自分で取り組みたい"}, ensure_ascii=False)))
        db.session.commit()
        check(_migrate_monthly_questions() is True, "移行が走る")
        db.session.commit()
        old = MonthlySurvey.query.filter_by(period="2024-01").first()
        keys = [q["key"] for q in old.questions]
        check("task_names" in keys and "decreased_reason" in keys and keys.index("has") < keys.index("hours") < keys.index("task_names") < keys.index("trend") < keys.index("decreased_reason") < keys.index("new_tasks"), f"古い設問セットにも 1-3（業務名）と 2-2（減った理由）が入り、基本設問の並びになる: {keys}")
        tq = next(q for q in old.questions if q["key"] == "trend")
        nq = next(q for q in old.questions if q["key"] == "task_names")
        check(nq["type"] == "list" and nq["placeholder"] == "例：日報の集計" and tq["text"] == "先月と比べて、PCの定型・定期業務は", "古い設問セットの 1-3 は複数回答になり、文言も「PCの定型・定期業務」にそろう")
        check(tq.get("depends", {}).get("key") == "has" and tq.get("depends", {}).get("main") is True and not tq.get("required"), "古い設問セットの 2（先月と比べて）も「ある」のときだけ出る設問になる")
        r1 = MonthlyResponse.query.filter_by(period="2024-01", emp_id="E001").first()
        r2 = MonthlyResponse.query.filter_by(period="2024-01", emp_id="E002").first()
        a2 = r2.answer("stance")
        check(r2.answer("hours") == "", "「ない」の回答は時間を空にする（以前の 0時間 も消す）")
        check("stance" in keys and "ask" not in keys and "self" not in keys and r1.answer("stance") == "改善チームに依頼したい" and a2 == "学びながら改善チームと一緒に取り組みたい", "旧2問は新1問に置き換わり、回答も読み替えられる")
        a3 = MonthlyResponse.query.filter_by(period="2024-01", emp_id="E003").first().answer("stance")
        h3 = MonthlyResponse.query.filter_by(period="2024-01", emp_id="E003").first().answer("has")
        a4 = MonthlyResponse.query.filter_by(period="2024-01", emp_id="E004").first().answer("stance")
        check(a3 == "学びながら改善チームと一緒に取り組みたい" and a4 == "" and h3 == "ある", "以前の4択は2択に読み替え、「ない」の人の答えは消える。「少しある」は「ある」に")
        a5 = MonthlyResponse.query.filter_by(period="2024-01", emp_id="E005").first().answer("stance")
        sq = next(q for q in old.questions if q["key"] == "stance")
        check(a5 == "自分で取り組みたい" and sq["options"] == ["自分で取り組みたい", "学びながら改善チームと一緒に取り組みたい", "改善チームに依頼したい"], "「自分で取り組みたい」はそのまま残り、設問セットの選択肢は3択になる")
        wq = next(q for q in old.questions if q["key"] == "wants_tasks")
        check(wq["type"] == "list" and wq["text"] == "あれば、その業務名", "古い設問セットの 3-2 も複数回答になる")
        check(r1.answer("hours") == "1〜2時間" and r1.answer("trend") == "" and r1.hours_shift == 1.5, "週あたりの時間は1シフトあたりに、「初回なので分からない」は未回答に読み替える")
        hq = next(q for q in old.questions if q["key"] == "hours")
        check(hq["options"][0] == "15分未満" and "0時間" not in hq["options"] and len(hq["options"]) == 10 and "合計時間" in hq["text"] and not hq["required"] and hq.get("depends", {}).get("key") == "has" and hq["text"].startswith("1シフト") and "定型・定期" in hq["text"], "設問セットの時間の設問も新しい文面・選択肢になり、「ある」のときだけ出る")
        check(_migrate_monthly_questions() is False, "2回目は何もしない")
        MonthlyResponse.query.filter_by(period="2024-01").delete()
        db.session.delete(old)
        db.session.commit()
    tampered = [dict(q, text="今月、繰り返しのPC作業はありましたか", options=["はい", "いいえ"], required=False) if q["key"] == "has" else q for q in qs]
    r = client.post("/monthly/admin/save", json={"period": period, "questions": tampered, "is_active": True})
    with app.app_context():
        has_q = next(q for q in MonthlySurvey.query.filter_by(period=period).first().questions if q["key"] == "has")
        check(r.status_code == 200 and has_q["options"] == ["ある", "ない"] and has_q["required"] and has_q["text"] == "今月、繰り返しのPC作業はありましたか", "基本設問は文面だけ変わり、選択肢と必須は固定のまま")
    r = client.post("/monthly/admin/save", json={"period": period, "questions": [q for q in qs if q["key"] != "stance"], "is_active": True})
    check(r.status_code == 400 and "基本設問は外せません" in r.get_json()["message"], "基本設問を外した保存は拒否")
    swapped = list(qs)
    i_t = next(i for i, q in enumerate(swapped) if q["key"] == "trend")
    i_r = next(i for i, q in enumerate(swapped) if q["key"] == "decreased_reason")
    swapped[i_t], swapped[i_r] = swapped[i_r], swapped[i_t]
    r = client.post("/monthly/admin/save", json={"period": period, "questions": swapped, "is_active": True})
    with app.app_context():
        saved_keys = [q["key"] for q in MonthlySurvey.query.filter_by(period=period).first().questions]
    check(r.status_code == 200 and saved_keys.index("trend") < saved_keys.index("decreased_reason"), "基本設問を並べ替えても、親（先月と比べて）は子（なぜ減ったのか）より前に保存される")
    code, html, r = get(f"/monthly/?period={period}&emp=DEMO-011")
    check("改善の勉強会に参加したいですか" in html and "参加したい" in html, "回答フォームに追加した設問が出る")
    r = client.post("/monthly/admin/save", json={"period": period, "questions": [{"key": "x", "type": "choice", "text": "一択", "options": ["a"]}]})
    check(r.status_code == 400 and "選択肢" in r.get_json()["message"], "選択肢が1つなら保存できない")
    code, html, r = get(f"/monthly/results?period={period}")
    check(code == 200 and "改善の勉強会に参加したいですか" in html and "テストのコメント" in html, "集計に追加設問とコメントが出る")

    r = client.post("/monthly/admin/save", json={"period": period, "questions": qs, "is_active": False})
    client.post("/logout")
    code, html, r = get(f"/monthly/?period={period}&emp=DEMO-012")
    check("準備中" in html, "受付停止中は準備中の表示")
    r = client.post("/monthly/submit", json=dict(payload, emp_id="DEMO-012", name="松本 玲奈"))
    check(r.status_code == 400, "受付停止中は送信できない")

    # ---- 回答できる月度は当月（と設問セットのある前月）だけ ----

    client.post("/login", data={"emp_id": "admin", "password": "adminpass"})
    client.post("/monthly/admin/save", json={"period": period, "questions": qs, "is_active": True})
    client.post("/logout")
    with app.app_context():
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
    r = client.post("/monthly/submit", json=dict(full, period=period, emp_id="X998", name="新人 次郎", workplace_id="1", work_pattern=""))
    check(r.status_code == 400 and "work_pattern" in r.get_json()["errors"], "初めての人は勤務形態（普通勤務／交代勤務）を選ばないと送信できない")
    r = client.post("/monthly/submit", json=dict(full, period=period, emp_id="X998", name="新人 次郎", workplace_id="1", work_pattern="交代勤務"))
    j = client.get(f"/monthly/api/employee/X998?period={period}").get_json()
    with app.app_context():
        e = Employee.query.filter_by(emp_id="X998").first()
        x998_pat = e.work_pattern if e else None
    check(r.status_code == 200 and "work_pattern" not in j and x998_pat == "交代勤務", "選んだ勤務形態は対象者一覧に控えるが、照会では返さない（次回も本人が選ぶ）")
    with app.app_context():
        row = MonthlyResponse.query.filter_by(period=period, emp_id="X998").first()
        check(row is not None and row.work_pattern == "交代勤務", "回答に勤務形態が保存される")
    code, html, r = get("/monthly/?emp=X998")
    check('id="work_pattern" class="form-select" required' in html and 'value="交代勤務" selected' not in html, "回答フォームに勤務形態の選択があり、控えがあっても選択済みにしない（本人が選ぶ）")
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
    check('data-type="list"' in html and html.count('name="ans-task_names"') == 2 and "業務を追加" in html and 'value="6時間以上"' in html and 'value="0時間"' not in html and "合計時間は" in html and "PC定型業務は" not in html,
          "1-3 は複数回答（業務名を1つずつ）、1-2 は10区分の合計時間、文言は「PCの定型・定期業務」")
    check('data-list-key="wants_tasks"' in html and "1-3 で答えた業務名を転記" in html and "（1行に1件）" not in html and 'value="自分で取り組みたい"' in html and 'value="学びながら改善チームと一緒に取り組みたい"' in html and "mq-list-del" in html,
          "3-2 は複数回答で「1-3 で答えた業務名を転記」ボタンつき、3-3 は3択")
    r = client.post("/monthly/submit", json=dict(full, emp_id="X111", name="複数 四", workplace_id="1", answers=dict(base_ans, wants="ある", wants_tasks=["日報の集計", "", "週報の作成"], stance="自分で取り組みたい")))
    with app.app_context():
        row = MonthlyResponse.query.filter_by(period=period, emp_id="X111").first()
        j11 = r.get_json()
        check(r.status_code == 200 and row is not None and row.answer("wants_tasks") == "日報の集計\n週報の作成" and row.answer("stance") == "自分で取り組みたい" and "src=monthly" in j11["register_url"] and j11["wants_tasks"] == ["日報の集計", "週報の作成"],
              "3-2 の複数の業務名を1行ずつ保存し、完了画面から「自動化を検討」へ進むリンクと業務名を返す。3-3 の「自分で取り組みたい」も受け付ける")
    r = client.get(j11["register_url"])
    check(r.status_code == 302 and "src=monthly" in r.headers["Location"], "/survey/start は src を引き継いで検討の画面へ")
    html11 = client.get(r.headers["Location"]).data.decode("utf-8")
    check('"self_improve": true' in html11 and '"stance": ' in html11 and 'id="rq-from-monthly"' in html11 and "rq-next-monthly" in html11,
          "アンケートから来ると、答え（業務名・進め方）を入力済みで始める準備が整う")
    r = client.post("/monthly/submit", json=dict(full, emp_id="X108", name="複数 一", workplace_id="1", answers=dict(base_ans, wants="ない", task_names=["日報の集計", " ", "出勤簿の転記"])))
    with app.app_context():
        row = MonthlyResponse.query.filter_by(period=period, emp_id="X108").first()
        check(r.status_code == 200 and row is not None and row.answer("task_names") == "日報の集計\n出勤簿の転記", "1-3 は複数の業務名を受け付け、空欄は除いて1行ずつ保存")
    r = client.post("/monthly/submit", json=dict(full, emp_id="X109", name="複数 二", workplace_id="1", answers=dict(base_ans, wants="ない", task_names=["", " "])))
    check(r.status_code == 400 and "task_names" in r.get_json()["errors"], "1-3 が全部空なら送信できない")
    form2 = {"period": period, "emp_id": "X110", "name": "複数 三", "workplace_id": "1", "work_pattern": "普通勤務", "comment": ""}
    for k_, v_ in dict(base_ans, wants="ない").items():
        form2["ans-" + k_] = v_
    form2["ans-task_names"] = ["週報の作成", "", "残業予定表の提出"]
    r = client.post("/monthly/submit", data=form2)
    with app.app_context():
        row = MonthlyResponse.query.filter_by(period=period, emp_id="X110").first()
        check(r.status_code == 302 and row is not None and row.answer("task_names") == "週報の作成\n残業予定表の提出", "通常のフォーム送信でも 1-3 の複数入力を1行ずつ保存")
    check('data-dep-key="wants"' in html and 'data-dep-key="trend"' in html and 'data-badge="req"' in html, "フォームに条件つき設問の情報が出る")
    check('<span class="mq-no">1-2</span>' in html and '<span class="mq-no">1-3</span>' in html and html.count('<span class="mq-no">2-2</span>') == 2 and '<span class="mq-no">3-2</span>' in html and '<span class="mq-no">3-3</span>' in html and '<span class="mq-no">3-1</span>' not in html and '<span class="mq-no">4-1</span>' not in html and 'data-dep-show=' in html, "枝番: 1-2（時間）・1-3（業務名）、2-2（減った理由／増えた業務名）、3-2（業務名）・3-3（進め方）")
    check('data-key="trend" data-type="choice" data-required="0" data-dep-key="has"' in html and 'data-key="wants" data-type="choice" data-required="0" data-dep-key="has"' in html and 'class="mq-q d-none" data-key="trend"' in html, "2（先月と比べて）と 3（なくしたい業務）は「ある」のときだけ出る（枝問の見た目ではなく通し番号）")
    r = client.post("/monthly/submit", json=dict(full, emp_id="X106", name="条件 六", workplace_id="1", answers=dict(base_ans, wants="ない", trend="減った", decreased_reason="")))
    check(r.status_code == 400 and "decreased_reason" in r.get_json()["errors"], "「減った」なのに理由が空なら送信できない")
    r = client.post("/monthly/submit", json=dict(full, emp_id="X106", name="条件 六", workplace_id="1", answers=dict(base_ans, wants="ない", trend="増えた", new_tasks="新しい帳票の転記", decreased_reason="消えるはず")))
    with app.app_context():
        row = MonthlyResponse.query.filter_by(period=period, emp_id="X106").first()
        check(r.status_code == 200 and row is not None and row.answer("decreased_reason") == "" and row.answer("new_tasks") == "新しい帳票の転記", "「増えた」なら理由は受け取らず、増えた業務名を保存")
    r = client.post("/monthly/submit", json=dict(full, emp_id="X107", name="条件 七", workplace_id="1", answers=dict(base_ans, wants="ない", task_names="")))
    check(r.status_code == 400 and "task_names" in r.get_json()["errors"], "「ある」なのに業務名（1-3）が空なら送信できない")
    r = client.post("/monthly/submit", json=dict(full, emp_id="X107", name="条件 七", workplace_id="1", answers=dict(base_ans, wants="ない", hours="")))
    check(r.status_code == 400 and "hours" in r.get_json()["errors"], "「ある」なのに時間（1-2）が空なら送信できない")
    r = client.post("/monthly/submit", json=dict(full, emp_id="X107", name="条件 七", workplace_id="1", answers=dict(full["answers"], has="ない", hours="1〜2時間", task_names="消えるはず", trend="減った", decreased_reason="消えるはず", wants="ある", wants_tasks="消えるはず", stance="学びながら改善チームと一緒に取り組みたい")))
    with app.app_context():
        row = MonthlyResponse.query.filter_by(period=period, emp_id="X107").first()
        check(r.status_code == 200 and row is not None and row.answer("has") == "ない" and all(row.answer(k) == "" for k in ("hours", "task_names", "trend", "decreased_reason", "wants", "wants_tasks", "stance")), "「ない」ならコメント以外は受け取らない（すべて空で保存）")
    r = client.post("/monthly/submit", json=dict(full, emp_id="X105", name="条件 五", workplace_id="1", answers=dict(base_ans, wants="ない", trend="変わらない", new_tasks="消えるはず")))
    with app.app_context():
        row = MonthlyResponse.query.filter_by(period=period, emp_id="X105").first()
        check(r.status_code == 200 and row is not None and row.answer("new_tasks") == "", "「変わらない」なら増えた業務名は出ない（受け取らない）")
    r = client.post("/monthly/submit", json=dict(full, emp_id="X104", name="条件 三", workplace_id="1", answers=dict(base_ans, wants="ある", wants_tasks="日報の転記", stance="")))
    check(r.status_code == 400 and r.get_json()["errors"].get("stance") == "選択してください", "「ある」なのに進め方を選んでいないと送信できない")
    r = client.post("/monthly/submit", json=dict(full, emp_id="X104", name="条件 三", workplace_id="1", answers=dict(base_ans, wants="ない", stance="学びながら改善チームと一緒に取り組みたい")))
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
        check(info is not None and info["name"] == "" and "workplace_id" not in info and Employee.query.filter_by(emp_id="E777").first() is None, "exists:true だけでも「いる」と分かる（照会しただけでは対象者一覧に載らない）")
        info = lookup("E997")
        check(info is not None and info["name"] == "名前 あり" and "workplace_id" not in info and "department" not in info, "照会で使うのは氏名だけ（所属が返っても職場は決めない）")
        info = lookup("E996")
        check(info is not None and info["name"] == "表示名 太郎" and "workplace_id" not in info, "氏名が display_name で返る API でも氏名が入る")
    r = client.get(f"/monthly/api/employee/E777?period={period}")
    j = r.get_json()
    check(j["found"] and set(j) == {"found", "answered", "emp_id"}, f"照会 API は見つかったか・回答済みかと社員IDだけを返す（氏名・職場・勤務形態は返さない）: {sorted(j)}")
    r = client.get(f"/monthly/api/employee/E996?period={period}")
    j = r.get_json()
    check(j["found"] and "name" not in j and "workplace_id" not in j and "work_pattern" not in j, "display_name の人も、氏名・職場は画面に返さない")
    r = client.post("/monthly/submit", json=dict(full, period=period, emp_id="E996", name="", workplace_id=""))
    check(r.status_code == 400 and set(r.get_json()["errors"]) == {"workplace_id"}, "初回に足りないのは職場だけ")
    r = client.post("/monthly/submit", json=dict(full, period=period, emp_id="E996", name="", workplace_id="3"))
    with app.app_context():
        e996 = Employee.query.filter_by(emp_id="E996").first()
        check(r.status_code == 200 and e996 is not None and e996.source == "answer" and e996.name == "表示名 太郎" and e996.active, "職場を選んで送信すると、対象者一覧に1行で載る（氏名は照会の結果）")
    for missing in ("E999", "E998"):
        r = client.post("/monthly/submit", json=dict(full, period=period, emp_id=missing, workplace_id="1"))
        check(r.status_code == 400 and "emp_id" in r.get_json()["errors"], f"LDAP にいない社員ID（{missing}）は送信できない")
    r = client.get(f"/monthly/api/employee/E996?period={prev}")
    j = r.get_json()
    check(j["found"] and "workplace_id" not in j and "work_pattern" not in j, "2回目でも職場・勤務形態は自動で入らない（毎回本人が選ぶ）")
    code, html_e, r = get(f"/monthly/?period={prev}&emp=E996")
    check(code == 200 and "selected>" not in html_e.split('id="workplace_id"', 1)[-1].split("</select>", 1)[0], "社員ID付きのリンクで開いても、職場は選ばれていない")
    r = client.post("/monthly/submit", json=dict(full, period=period, emp_id="E777", name="", workplace_id=""))
    check(r.status_code == 400 and set(r.get_json()["errors"]) == {"workplace_id"}, "初回に足りないのは職場だけ（氏名は聞かない）")
    r = client.post("/monthly/submit", json=dict(full, period=period, emp_id="E777", name="初回 入力", workplace_id="2"))
    check(r.status_code == 200, "職場を選んで送信（送られてきた氏名は使わない）")
    r = client.get(f"/monthly/api/employee/E777?period={prev}")
    j = r.get_json()
    with app.app_context():
        emp = Employee.query.filter_by(emp_id="E777").first()
        check(j["found"] and "workplace_id" not in j and emp.name == "" and emp.workplace_id == 2, "職場は自動で入らないが、選んだ職場は対象者一覧に残る")
    client.post("/survey/api/reset")
    r = client.post("/survey/api/respondent", json={"emp_id": "E777", "workplace_id": ""})
    j = r.get_json()
    check(r.status_code == 400 and {"workplace_id", "work_pattern"} <= set(j["errors"]), "自動化を検討でも職場・勤務形態は本人が選ぶ（社員IDだけでは入らない）")
    r = client.post("/survey/api/respondent", json={"emp_id": "E999", "workplace_id": 1, "work_pattern": "普通勤務"})
    check(r.status_code == 400 and "emp_id" in r.get_json()["errors"], "自動化を検討でも LDAP にいない社員IDは進めない")
    ldap_auth.urllib.request.urlopen = _real_urlopen
    conf.override("LDAP_API_URL", "")

    with app.app_context():
        import app as app_module
        for i, (has, trend) in enumerate((("ない", ""), ("ない", ""), ("ある", "減った"), ("ある", "減った"))):
            db.session.add(MonthlyResponse(period="2001-01", emp_id=f"R{i}", name=f"率 {i}", workplace_id=1, answers_json=json.dumps({"has": has, "trend": trend}, ensure_ascii=False)))
        db.session.commit()
        row = next(x for x in app_module.monthly_summary() if x["period"] == "2001-01")
        check(row["has_rate"] == 50.0 and row["decreased_rate"] == 100.0, f"月次の割合は、その設問に答えた人が分母（あり {row['has_rate']}%、減った {row['decreased_rate']}%）")
        MonthlyResponse.query.filter_by(period="2001-01").delete()
        db.session.add(MonthlyResponse(period="2001-02", emp_id="H1", name="時間 旧", workplace_id=1, answers_json=json.dumps({"has": "ある", "hours": "30分〜1時間"}, ensure_ascii=False)))
        db.session.add(MonthlyResponse(period="2001-02", emp_id="H2", name="時間 新", workplace_id=1, answers_json=json.dumps({"has": "ある", "hours": "45分〜1時間"}, ensure_ascii=False)))
        db.session.commit()
        ht = next(x for x in app_module.monthly_question_trends() if x["q"]["key"] == "hours")
        k = ht["periods"].index("2001-02")
        band = next(x for x in ht["series"] if x["option"] == "30分〜1時間")
        h_old = MonthlyResponse.query.filter_by(period="2001-02", emp_id="H1").first().hours_shift
        h_new = MonthlyResponse.query.filter_by(period="2001-02", emp_id="H2").first().hours_shift
        check(band["percents"][k] == 100 and [x["option"] for x in ht["series"]] == ["30分未満", "30分〜1時間", "1〜2時間", "2〜4時間", "4時間以上"] and ht["note"] and h_old == 0.75 and h_new == 0.875,
              "1-2 の推移は以前の区切りにまとめ（以前の回答と新しい回答を同じ区切りで比べる）、換算は回答した区分の値")
        MonthlyResponse.query.filter_by(period="2001-02").delete()
        for i, st in enumerate(("自分で取り組みたい", "学びながら改善チームと一緒に取り組みたい", "改善チームに依頼したい")):
            db.session.add(MonthlyResponse(period="2001-03", emp_id=f"S{i}", name=f"進め方 {i}", workplace_id=1, answers_json=json.dumps({"has": "ある", "wants": "ある", "stance": st}, ensure_ascii=False)))
        db.session.commit()
        row = next(x for x in app_module.monthly_summary() if x["period"] == "2001-03")
        check(round(row["self_only_rate"]) == 33 and round(row["together_rate"]) == 33 and round(row["ask_rate"]) == 33 and round(row["self_rate"]) == 67,
              "3-3 は3択それぞれの割合と、自分で・一緒にの合計を出す")
        MonthlyResponse.query.filter_by(period="2001-03").delete()
        db.session.commit()

    # ---- 対象者（一度でも回答した人）と回答率の分母 ----
    with app.app_context():
        import app as app_module
        for e in ("M1", "M2", "M3"):
            db.session.add(Employee(emp_id=e, name=f"対象 {e}", source="answer", active=(e != "M3")))
        for p, ids in (("2002-01", ("M1",)), ("2002-02", ("M1", "M2")), ("2002-03", ("M2", "M3"))):
            for e in ids:
                db.session.add(MonthlyResponse(period=p, emp_id=e, name=f"対象 {e}", workplace_id=1, answers_json="{}"))
        db.session.commit()
        rows = [x for x in app_module.member_rows() if x["emp"].emp_id in ("M1", "M2", "M3")]
        st = app_module.member_stats(rows, ["2002-01", "2002-02", "2002-03"])
        check([(st[p]["answered"], st[p]["members"]) for p in ("2002-01", "2002-02", "2002-03")] == [(1, 1), (2, 2), (1, 2)] and st["2002-03"]["rate"] == 50.0,
              f"分母はその月までに一度でも回答した人（対象外は除く）: {[(st[p]['answered'], st[p]['members']) for p in st]}")
        db.session.add(Employee(emp_id="L001", name="照会だけ", source="ldap"))
        db.session.add(MonthlyResponse(period="2002-01", emp_id="M9", name="控え 九", workplace_id=1, answers_json="{}"))
        db.session.commit()
        app_module._sync_members()
        db.session.commit()
        check(Employee.query.filter_by(emp_id="L001").first() is None and Employee.query.filter_by(emp_id="M9").count() == 1, "起動時: 回答した人はマスタへ（重複なし）、照会しただけで残っていた行は消す")
        app_module._sync_members()
        db.session.commit()
        check(Employee.query.filter_by(emp_id="M9").count() == 1, "何度起動しても同じ人は1行")
        MonthlyResponse.query.filter(MonthlyResponse.period.in_(["2002-01", "2002-02", "2002-03"])).delete(synchronize_session=False)
        Employee.query.filter(Employee.emp_id.in_(["M1", "M2", "M3", "M9"])).delete(synchronize_session=False)
        db.session.commit()
        e009 = Employee.query.filter_by(emp_id="DEMO-009").first().id
    client.post("/login", data={"emp_id": "admin", "password": "adminpass"})
    code, html, r = get("/admin/employees")
    check(code == 200 and "対象者一覧" in html and "対象外" in html and "自動化を検討" in html and "member-yes" in html and "回答 ／ 対象（回答率）" in html and "E996" in html and "NOPE01" not in html,
          "対象者一覧: 社員ID・氏名・自動化を検討の回数 × 月ごとの回答チェック")
    r = client.post(f"/admin/employees/{e009}/toggle", json={"excluded": True})
    j = r.get_json()
    with app.app_context():
        check(j["success"] and j["active"] is False and not db.session.get(Employee, e009).active and j["stats"] and j["excluded_count"] >= 1, "対象外にチェックすると分母から外れる")
    r = client.post(f"/admin/employees/{e009}/toggle", json={"excluded": False})
    check(r.get_json()["active"] is True, "チェックを外すと対象に戻る")
    code, html, r = get("/")
    check("月次アンケート回答率" in html and "対象 " in html, "ダッシュボードの回答率は対象者一覧の最新月で出す")
    code, html, r = get("/admin/")
    check("アンケート対象者数" not in html and "目標期限" not in html and "target_respondents" not in html, "管理: 対象者数・目標期限の設定はない")

    # ---- DX人財情報 ----
    client.post("/login", data={"emp_id": "admin", "password": "adminpass"})
    code, html, r = get("/people/")
    check(code == 200 and "DX人財情報" in html and "人材育成" not in html and "田中 一郎" in html and "E777" in html, "DX人財情報: 自動化を検討の回答者と、月次だけ答えた人が並ぶ")
    check("自分で改善したい" in html and "直近の月次で「自分で・一緒に取り組みたい」" in html and "人財一覧" in html, "DX人財情報: 集計カードと一覧")
    r = client.get("/people/", query_string={"wish": "positive"})
    html = r.data.decode("utf-8")
    check(r.status_code == 200 and "佐藤 花子" in html, "自分で改善したい人で絞り込み")
    r = client.get("/people/", query_string={"wish": "スキルを身につけながら、一緒にやってみたい"})
    html = r.data.decode("utf-8")
    with app.app_context():
        import app as app_module
        yes = [p["name"] for p in app_module.dx_people(wish="スキルを身につけながら、一緒にやってみたい")]
        others = [p["name"] for p in app_module.dx_people() if p["name"] and p["name"] not in yes]
    check(r.status_code == 200 and yes and others and all(n in html for n in yes) and not any(n in html for n in others), f"意欲（依頼で選んだもの）で絞り込み（一緒に {len(yes)} 人／それ以外 {len(others)} 人）")
    r = client.post("/admin/employees/add", data={"emp_id": "N001", "name": "新規 社員", "department": "", "email": "", "workplace_id": "1", "work_pattern": "交代勤務"})
    with app.app_context():
        e = Employee.query.filter_by(emp_id="N001").first()
        check(r.status_code == 302 and e is not None and e.work_pattern == "交代勤務", "社員の手動追加で勤務形態を登録できる")
    r = client.get("/people/", query_string={"workplace_id": 99})
    check(r.status_code == 200 and "該当する人はいません" in r.data.decode("utf-8"), "該当なしの表示")

    return fails

def smoke_request():
    import urllib.request

    import app as app_module

    tmpdir = tempfile.mkdtemp()
    dbfile = os.path.join(tmpdir, "test.sqlite3").replace("\\", "/")
    app = create_app({"SQLALCHEMY_DATABASE_URI": "sqlite:///" + dbfile, "TESTING": True})
    client = app.test_client()
    fails = []

    def check(cond, msg):
        print(("OK  " if cond else "NG  ") + msg)
        if not cond:
            fails.append(msg)

    def get(url):
        r = client.get(url)
        return r.status_code, r.data.decode("utf-8"), r

    conf.override("OPENAI_REASONING_EFFORT", "low")
    conf.override("OPENAI_MAX_OUTPUT_TOKENS", 4000)
    conf.override("OPENAI_API_TIMEOUT", 90)
    client.post("/login", data={"emp_id": "admin", "password": "adminpass"})
    client.post("/admin/demo/load")
    client.post("/admin/ai", data={"openai_api_url": "https://ai.example.test/api/responses", "openai_api_key": "test-key", "openai_model": "gpt-6-test"})
    with app.app_context():
        db.session.add(Employee(emp_id="T001", name="試験 太郎", source="manual"))
        db.session.add(Employee(emp_id="T002", name="試験 次郎", source="manual"))
        db.session.add(MonthlyResponse(period=current_period(), emp_id="T001", name="試験 太郎", workplace_id=1, answers_json=json.dumps({"has": "ある", "hours": "30分〜1時間", "task_names": "日報の集計、出勤簿の転記", "trend": "変わらない", "wants": "ある", "wants_tasks": "日報の集計", "stance": "改善チームに依頼したい"}, ensure_ascii=False)))
        db.session.commit()
    client.post("/logout")

    code, html, r = get("/survey/?emp=T001&task=日報の集計")

    def in_html(text):
        return text in html or json.dumps(text)[1:-1] in html

    check(code == 200 and "聞き取りを始める" in html and "一気に書く" in html and "写真を貼る" in html and "最初から見る" in html and 'id="rq-dup-note"' in html and 'id="rq-dup"' in html and in_html("まずは相談したい") and in_html("スキルを身につけながら、一緒にやってみたい") and "mermaid" in html and in_html("日報の集計"), "依頼ページ: 聞き取り・一気に書く・写真・再生・3つの希望・一緒にやってみたい・業務名の引き継ぎ")
    check("検討する業務" in html and "改善への関わり方" not in html and "部全体について" not in html, "第3節・第4節は無くなっている")
    js = client.get("/static/survey.js").data.decode("utf-8")
    check(not in_html("氏名") and "rq-fallback" not in js and "fillFallback" not in js and "社員IDを確認しました" in js, "あなたのこと: 社員IDは確認だけで、氏名の欄はない")

    r = client.post("/survey/api/intake", json={"title": "x", "messages": [], "state": {}})
    check(r.status_code == 400 and "あなたのこと" in r.get_json()["message"], "あなたのこと未保存では聞き取りを始められない")
    r = client.post("/survey/api/respondent", json={"emp_id": "T001", "name": "試験 太郎", "workplace_id": 1})
    check(r.status_code == 400 and "work_pattern" in r.get_json()["errors"], "初めての人は勤務形態を選ばないと保存できない")
    r = client.post("/survey/api/respondent", json={"emp_id": "T001", "name": "試験 太郎", "workplace_id": 1, "work_pattern": "交代勤務"})
    j = r.get_json()
    check(r.status_code == 200 and j["suggested"] == ["日報の集計", "出勤簿の転記"] and j["suggested_wants"] == ["日報の集計"] and j["respondent"]["work_pattern"] == "交代勤務", f"アンケートで挙げた業務が候補に出て（なくしたい業務が先）、勤務形態が保存される: {j.get('suggested')}")
    code, html_m, r = get("/survey/?emp=T001&src=monthly")
    check(code == 200 and '"self_improve": false' in html_m and ('"wants_tasks": ["\\u65e5' in html_m or '"wants_tasks": ["日報の集計"]' in html_m), "アンケートから来ると、なくしたい業務と進め方（依頼したい＝チェックなし）を渡す")
    check('"respondent": null' in html_m, "開いたときは前の回答者を戻さず、空で始める")
    check('"workplace_id": 1' in html_m, "月次アンケートから来たときは、その回答で選んだ職場を引き継ぐ")
    client.post("/survey/api/respondent", json={"emp_id": "T001", "workplace_id": 1, "work_pattern": "交代勤務"})

    calls, script = [], []

    class _Resp:
        def __init__(self, payload):
            self.status = 200
            self._body = json.dumps(payload, ensure_ascii=False).encode("utf-8")

        def read(self):
            return self._body

        def __enter__(self):
            return self

        def __exit__(self, *exc):
            return False

    def fake_urlopen(req, timeout=None):
        raw = req.data.decode("utf-8")
        calls.append({"body": json.loads(raw), "raw": raw})
        nxt = script.pop(0)
        if isinstance(nxt, Exception):
            raise nxt
        return _Resp(nxt)

    def message(obj):
        return {"id": "resp", "object": "response", "status": "completed", "error": None, "output": [{"type": "message", "id": "m", "status": "completed", "role": "assistant", "content": [{"type": "output_text", "text": json.dumps(obj, ensure_ascii=False), "annotations": []}]}]}

    def state(**kw):
        base = {"title": "日報の集計", "steps": [], "sources": [], "outputs": [], "recipients": [], "frequency": None, "duration": None, "pains": [], "unknowns": []}
        base.update(kw)
        return base

    step1 = {"no": 1, "action": "生産管理システムから前日の実績をCSVで出す", "system": "生産管理システム", "input": "前日の実績", "output": "", "to": ""}
    step2 = {"no": 2, "action": "Excelの集計表に貼り付ける", "system": "Excel（集計表）", "input": "CSV", "output": "", "to": ""}
    step3 = {"no": 3, "action": "不良件数を手で数えて追記する", "system": "Excel（集計表）", "input": "", "output": "日次実績集計表", "to": ""}
    step4 = {"no": 4, "action": "課長にメールで送る", "system": "メール", "input": "", "output": "", "to": "課長"}
    real_urlopen = urllib.request.urlopen
    urllib.request.urlopen = fake_urlopen
    try:
        script.append(message({"reply": "「日報の集計」ですね。最初に何をしますか？", "chips": [], "done": False, "remaining": 5, "state": state()}))
        r = client.post("/survey/api/intake", json={"title": "日報の集計", "messages": [], "state": {}})
        j = r.get_json()
        check(r.status_code == 200 and j["success"] and "最初に何を" in j["reply"] and j["done"] is False and j["remaining"] == 5 and j["diagram"] == "" and j["state"]["title"] == "日報の集計", "聞き取り1ターン目: 最初の質問が返り、図はまだ空")
        b = calls[-1]["body"]
        check(b["model"] == "gpt-6-test" and "聞き取り担当" in b["instructions"] and "現在の state" in b["instructions"] and isinstance(b["input"], list) and b["input"][0]["role"] == "user" and "日報の集計" in b["input"][0]["content"] and "temperature" not in calls[-1]["raw"], "1ターン目の本文: 聞き取りのプロンプト・state・業務名。temperature なし")
        check([x["name"] for x in b["tools"]] == ["person_history", "list_known_systems", "search_similar_tasks", "check_duplicate"] and b["tool_choice"] == "auto" and b["store"] is False and "使える道具" in b["instructions"], "聞き取りには4つの道具を平らな形で渡す")
        check("交代勤務（昼勤・夜勤あり）" in b["instructions"] and "昼勤・夜勤の両方" in b["instructions"], "交代勤務の人には、いつやるかを聞く指示が入る")
        msgs = [{"role": "ai", "text": j["reply"]}, {"role": "user", "text": "生産管理システムから前日の実績をCSVで出す"}]
        script.append({"id": "resp_t", "object": "response", "status": "completed", "error": None, "output": [
            {"type": "function_call", "id": "fc_a", "call_id": "call_ph", "name": "person_history", "arguments": "{}", "status": "completed"},
            {"type": "function_call", "id": "fc_b", "call_id": "call_dup", "name": "check_duplicate", "arguments": json.dumps({"name": "日報の集計"}, ensure_ascii=False), "status": "completed"},
        ]})
        script.append(message({"reply": "なるほど。次に何をしますか？", "chips": [], "done": False, "remaining": 4, "state": state(steps=[step1], sources=["生産管理システム"])}))
        n_before = len(calls)
        r = client.post("/survey/api/intake", json={"title": "日報の集計", "messages": msgs, "state": j["state"], "shown": 0})
        j = r.get_json()
        b = calls[-1]["body"]
        outs = {i["call_id"]: json.loads(i["output"]) for i in b["input"] if i.get("type") == "function_call_output"}
        check(len(calls) == n_before + 2 and r.status_code == 200 and "氏名" not in outs["call_ph"] and outs["call_ph"]["勤務形態"] == "交代勤務" and outs["call_ph"]["月次アンケート"]["行っている業務"] == "日報の集計、出勤簿の転記" and outs["call_dup"]["二重登録"] is False, "聞き取り中に AI が本人の履歴と二重登録の有無を道具で調べ、結果を受け取って続ける")
        check(r.status_code == 200 and len(j["state"]["steps"]) == 1 and "st1" in j["diagram"] and "生産管理システム" in j["diagram"] and "flowchart LR" in j["diagram"], "2ターン目: 手順1が入り、図に手順とシステムが出る")
        check(len(j["frames"]) == 1 and j["frames"][0] == j["diagram"], "増えた手順のぶんだけコマ送りの図（frames）が返る")
        check([m.get("role") for m in b["input"] if m.get("role")] == ["assistant", "user"] and b["input"][1]["content"] == "生産管理システムから前日の実績をCSVで出す", "会話の記録をそのまま送る（道具の項目は role なしで続く）")
        msgs += [{"role": "ai", "text": j["reply"]}, {"role": "user", "text": "Excelの集計表に貼り付ける\n不良件数を手で数えて追記する"}]
        script.append(message({"reply": "集計表は誰に渡しますか？", "chips": ["課長", "生産管理課", "自分だけ", "分からない"], "done": False, "remaining": 3, "state": state(steps=[step1, step2, step3], sources=["生産管理システム"], outputs=["日次実績集計表"])}))
        r = client.post("/survey/api/intake", json={"title": "日報の集計", "messages": msgs, "state": j["state"], "shown": 1})
        j = r.get_json()
        check(r.status_code == 200 and len(j["state"]["steps"]) == 3 and j["chips"] == ["課長", "生産管理課", "自分だけ", "分からない"] and "o3" in j["diagram"], "一気に書いた2行が手順2・3になり、選択肢（chips）が返る")
        check(len(j["frames"]) == 2 and "st2" in j["frames"][0] and "st3" not in j["frames"][0] and "st3" in j["frames"][1], "一気に2手順増えたら、1手順ずつのコマが2枚返る")
        b3 = calls[-1]["body"]
        check([x["name"] for x in b3["tools"]] == ["list_known_systems", "search_similar_tasks"] and "person_history" in j["state"]["checked"], "調べ済みの道具（本人の履歴・同じ名前の確認）は次のターンでは渡さない")
        msgs += [{"role": "ai", "text": j["reply"]}, {"role": "user", "text": "手順書（写真）を送りました。読み取ってください"}]
        image = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=="
        script.append(message({"reply": "写真から4つ目の手順を読み取りました。この流れで合っていますか？", "chips": [], "done": True, "remaining": 0, "state": state(steps=[step1, step2, step3, step4], sources=["生産管理システム"], outputs=["日次実績集計表"], recipients=["課長"], frequency="毎日", duration="15〜30分")}))
        r = client.post("/survey/api/intake", json={"title": "日報の集計", "messages": msgs, "state": j["state"], "image": image, "shown": 3})
        j = r.get_json()
        b = calls[-1]["body"]
        last = b["input"][-1]["content"]
        check(r.status_code == 200 and isinstance(last, list) and last[0]["type"] == "input_text" and last[1]["type"] == "input_image" and last[1]["image_url"] == image, "写真は最後のユーザー発言に input_image として付く")
        check(j["done"] is True and len(j["state"]["steps"]) == 4 and j["state"]["frequency"] == "毎日" and j["state"]["duration"] == "15〜30分" and "課長" in j["diagram"] and len(j["frames"]) == 1, "写真を読んで手順4・頻度・時間がそろい、確認（done）になる")
        final_state = j["state"]
        msgs += [{"role": "ai", "text": j["reply"]}, {"role": "user", "text": "はい、合っています"}]

        r = client.post("/survey/api/intake", json={"title": "x", "messages": [], "state": {}, "image": "data:text/plain;base64,AAAA"})
        check(r.status_code == 400 and "写真" in r.get_json()["message"], "画像以外のデータは受け付けない")
        script.append(message({"reply": "アンケートで「日報の集計」をなくしたいと答えてくださいましたね。最初に何をしますか？", "chips": [], "done": False, "remaining": 5, "state": state()}))
        r = client.post("/survey/api/intake", json={"title": "日報の集計", "messages": [], "state": {}, "origin": "monthly"})
        check(r.status_code == 200 and "月次アンケートで「なくしたい・自動化したい」と答えた業務" in calls[-1]["body"]["input"][0]["content"], "アンケートから来たときは、AI がそのことにふれてから聞き始める")

        r = client.post("/survey/api/tasks", json={"name": "", "state": {}, "messages": [], "frequency": "", "duration": "", "wish": ""})
        e = r.get_json()["errors"]
        check(r.status_code == 400 and "name" in e and "steps" in e and "frequency" in e and "duration" in e and "wish" in e, "登録: 業務名・手順・頻度・時間・希望がないとエラー")
        r = client.post("/survey/api/tasks", json={"state": final_state, "messages": msgs, "frequency": "毎日", "duration": "15〜30分", "wish": "自動化したい"})
        check(r.status_code == 400 and "wish" in r.get_json()["errors"], "希望は3つの選択肢以外を受け付けない")

        finalize = {
            "name": "日報の実績を集計表にまとめて課長に送る", "purpose": "前日の生産実績を朝会で確認するため", "output": "日次実績集計表", "output_users": ["上司"], "output_users_note": "朝会で使う",
            "inputs": ["生産管理システム", "他のExcel"], "inputs_note": "生産管理システムの実績CSV、Excelの集計表", "info_from": "生産管理システム", "actors": "課長", "prev_task": "", "next_task": "週報の作成",
            "tools": ["Excel", "メール", "社内システム"], "nature": ["転記・コピー", "集計・グラフ作成", "メール・配布"], "pains": ["時間がかかる", "同じ情報を何度も入力する"], "pains_note": "手で数えている", "headcount": 2,
            "domain": "製造実行・管理", "related": ["週報の作成"], "summary": "日次実績を集計表にまとめて課長に送る日次業務。", "ecrs": {"eliminate": None, "combine": "不良件数は品質システムから取れる可能性", "rearrange": None, "simplify": "集計表の自動更新"},
            "automation": "Power Query で CSV を取り込み、メールは自動送信", "flags": ["自動化候補", "重複入力候補"], "priority": "高", "questions": ["不良件数は品質システムに入っていますか？"],
            "tobe_steps": ["CSV を Power Query で取り込む", "集計表が自動で更新される", "課長はダッシュボードで確認"], "saving_minutes_per_month": 400, "first_impression": "手で数えている不良件数は、システムから取れそうです", "hearing_questions": [],
        }
        r = client.post("/survey/api/tasks", json={"state": final_state, "messages": msgs, "frequency": "毎日", "duration": "15〜30分", "wish": "減らしたい"})
        check(r.status_code == 400 and "shift" in r.get_json()["errors"], "交代勤務の人は、いつやるか（昼勤／夜勤／両方）がないと登録できない")
        script.append(message(finalize))
        r = client.post("/survey/api/tasks", json={"state": final_state, "messages": msgs, "frequency": "毎日", "duration": "15〜30分", "wish": "減らしたい", "self_improve": True, "skills": ["Excel関数"], "learn": ["Python"], "shift": "夜勤"})
        j = r.get_json()
        b = calls[-1]["body"]
        check(r.status_code == 200 and j["success"] and j["first_impression"] == finalize["first_impression"] and len(j["tasks"]) == 1 and j["suggested"] == ["出勤簿の転記"], "登録: AI の一言が返り、候補から登録した業務が消える")
        fin_in = b["input"][0]["content"]
        check("業務分析アシスタント" in b["instructions"] and "聞き取りの結果" in fin_in and "減らしたい" in fin_in and "製造実行・管理" in fin_in and "## 勤務\n夜勤" in fin_in and "temperature" not in calls[-1]["raw"] and [x["name"] for x in b["tools"]] == ["search_similar_tasks", "duplicate_outputs"], "登録時の分析: state・希望・勤務・業務領域リストと2つの道具を送る。temperature なし")
        with app.app_context():
            t = Task.query.filter_by(name="日報の集計").first()
            check(t is not None and len(t.steps) == 4 and t.steps[3]["to"] == "課長" and t.frequency == "毎日" and t.duration == "15〜30分" and t.wish == "減らしたい" and t.self_improve == "スキルを身につけながら、一緒にやってみたい" and t.shift == "夜勤" and t.monthly_minutes > 0, "案件に手順・頻度・時間・希望・一緒にやってみたい・勤務（夜勤）が保存される")
            check(t.purpose == finalize["purpose"] and t.output_users_list == ["上司"] and t.inputs_list == ["生産管理システム", "他のExcel"] and t.tools_list == ["Excel", "メール", "社内システム"] and t.pains_list == ["時間がかかる", "同じ情報を何度も入力する"] and t.domain.name == "製造実行・管理" and t.next_task == "週報の作成", "AI が埋めた項目が案件に入る")
            check(t.headcount == 1 and t.ai_extra_dict.get("headcount_guess") == 2 and t.monthly_minutes == 20 * 22, f"AI が推定した人数は掛けず（担当人数は1のまま）、参考として残す: {t.headcount}人 {t.monthly_minutes}分")
            check(set(t.ai_fields_list) >= {"purpose", "output", "inputs", "tools", "domain"} and "headcount" not in t.ai_fields_list and not t.ai_extra_dict.get("rule_fields") and t.ai_source == "openai" and t.ai_priority == "高" and "自動化候補" in t.ai_flags_list and t.ai_extra_dict["saving_minutes_per_month"] == 400 and t.ai_extra_dict["tobe_steps"][0].startswith("CSV"), "AI推定の印・分析結果・改善後の流れが保存される")
            check(len(t.intake) == len(msgs) and t.intake[-1]["text"] == "はい、合っています", "聞き取りの記録が保存される")
            r_ = Respondent.query.filter_by(emp_id="T001").first()
            check(r_.c2_skills == "Excel関数" and r_.c3_learn == "Python" and r_.wants_to_improve, "一緒にやってみたい人の技術が残り、自主改善希望者になる")
            tid = t.id

        r = client.post("/survey/api/tasks", json={"state": final_state, "messages": msgs, "frequency": "毎日", "duration": "15〜30分", "wish": "減らしたい", "shift": "夜勤"})
        j = r.get_json()
        check(r.status_code == 409 and j["duplicate"]["name"] == "日報の集計" and "登録済み" in j["message"], "同じ人が同じ業務名をもう一度登録しようとすると 409 で警告（AI は呼ばない）")
        script.append({"id": "resp_t", "object": "response", "status": "completed", "error": None, "output": [
            {"type": "function_call", "id": "fc_s", "call_id": "call_sim", "name": "search_similar_tasks", "arguments": json.dumps({"query": "日報の集計"}, ensure_ascii=False), "status": "completed"},
            {"type": "function_call", "id": "fc_d", "call_id": "call_do", "name": "duplicate_outputs", "arguments": "{}", "status": "completed"},
        ]})
        script.append(message(dict(finalize, name="2件目")))
        r = client.post("/survey/api/tasks", json={"state": final_state, "messages": msgs, "frequency": "毎日", "duration": "15〜30分", "wish": "減らしたい", "shift": "夜勤", "force": True})
        j = r.get_json()
        outs = {i["call_id"]: json.loads(i["output"]) for i in calls[-1]["body"]["input"] if i.get("type") == "function_call_output"}
        check(r.status_code == 200 and j["success"] and len(j["tasks"]) == 2 and outs["call_sim"]["found"] >= 1 and any(row["業務名"] == "日報の集計" for row in outs["call_sim"]["業務"]) and "同じできるものを複数人が作っている組" in outs["call_do"], "force なら別の業務として登録し、分析の道具（類似業務・重複成果物）が実行される")
        r = client.delete(f"/survey/api/tasks/{j['task']['id']}")
        check(r.status_code == 200 and len(r.get_json()["tasks"]) == 1, "重複分を取り消す")
        script.append(message(dict(finalize, name="相談の業務", hearing_questions=["いまの帳票は誰が見ていますか？", "数字を使う会議はありますか？"], wish="まずは相談したい")))
        r = client.post("/survey/api/tasks", json={"name": "相談したい業務", "state": state(title="相談したい業務", steps=[step1, step2]), "messages": [], "frequency": "毎週", "duration": "15分未満", "wish": "まずは相談したい", "shift": "昼勤"})
        check(r.status_code == 200 and r.get_json()["success"], "「まずは相談したい」でも登録できる")
        with app.app_context():
            t2 = Task.query.filter_by(name="相談したい業務").first()
            check(t2 is not None and t2.is_candidate and t2.pains_list and len(t2.ai_extra_dict["hearing_questions"]) == 2, "相談の案件は困りごとがあれば候補に数え、ヒアリング質問が残る")
            check(not Task(wish="まずは相談したい", pains="特になし").is_candidate and Task(wish="まずは相談したい", pains="時間がかかる").is_candidate and Task(wish="減らしたい").is_candidate and Task(wish="自動化したい").is_candidate,
                  "案件候補は「なくしたい・減らしたい」か本当の困りごと（「まずは相談したい」だけでは候補にしない）")

        script.append(message(dict(finalize, name="", purpose="", output_users=[], inputs=[], tools=[], nature=[], pains=[], pains_note="", domain="存在しない領域", headcount=1,
                                   summary="田中 一郎さんと確認する", tobe_steps=["佐藤課長に送る"], first_impression="鈴木さんに相談します", hearing_questions=["山田さんは何に使いますか？"],
                                   ecrs={"eliminate": "佐藤さんに聞く", "combine": None, "rearrange": None, "simplify": None}, questions=["佐藤さんは使っていますか？"])))
        r = client.post("/survey/api/tasks", json={"name": "佐藤さんの日報を写す", "state": state(title="佐藤さんの日報を写す", steps=[step1, step4], sources=["生産管理システム"], recipients=["課長"], pains=["同じ数字を2回入れている"], purpose="朝会で使う"),
                                                   "messages": [{"role": "user", "text": "佐藤さんから日報をもらう"}], "frequency": "毎週", "duration": "15分未満", "wish": "減らしたい", "shift": "昼勤"})
        check(r.status_code == 200 and r.get_json()["success"], "AI が区分を空で返しても登録できる")
        with app.app_context():
            t4 = Task.query.filter_by(name="担当者の日報を写す").first()
            ex4 = t4.ai_extra_dict if t4 else {}
            check(t4 is not None and t4.intake[0]["text"] == "担当者から日報をもらう" and t4.ai_summary == "担当者と確認する" and ex4.get("tobe_steps") == ["課長に送る"] and ex4.get("first_impression") == "担当者に相談します"
                  and ex4.get("hearing_questions") == ["担当者は何に使いますか？"] and t4.ai_ecrs_dict.get("eliminate") == "担当者に聞く" and t4.ai_ecrs_dict.get("questions") == ["担当者は使っていますか？"],
                  "業務名・聞き取りの記録・AI の要約・ECRS・改善の案・ヒアリング質問の氏名は担当者に置き換えてから保存する")
            check(t4 is not None and t4.inputs_list == ["生産管理システム"] and t4.output_users_list == ["上司"] and t4.tools_list and t4.nature_list and set(ex4.get("rule_fields") or []) >= {"inputs", "output_users", "tools", "nature", "pains", "domain"},
                  f"AI が空で返した区分（利用者・入力元・ツール・作業の性質・困りごと）はルールで埋め、推定の印を付ける: {ex4.get('rule_fields')}")
            check(t4 is not None and t4.pains_list == ["同じ情報を何度も入力する"] and t4.pains_note == "同じ数字を2回入れている" and t4.purpose == "朝会で使う", "困りごとの言葉（2回入れている）を区分に分け、元の言葉は困りごとの補足に残す")
            check(t4 is not None and t4.domain is not None and t4.domain.name == "製造実行・管理" and t4.ai_domain == "製造実行・管理" and "domain" in t4.ai_fields_list, "AI が一覧にない業務領域を返したら、ルールの推定で業務領域を埋める")

        script.append(urllib.error.URLError("接続できない（テスト）"))
        r = client.post("/survey/api/tasks", json={"name": "AIなしでも登録", "state": state(title="AIなしでも登録", steps=[step1, step4], sources=["生産管理システム"], recipients=["課長"], shift="昼勤・夜勤の両方", pains=["手で入れるのでミスが起きやすい"]), "messages": [], "frequency": "毎日", "duration": "30分〜1時間", "wish": "なくしたい"})
        check(r.status_code == 200 and r.get_json()["success"], "AI が応答しなくても登録はできる（ルールベースで補う）")
        with app.app_context():
            t3 = Task.query.filter_by(name="AIなしでも登録").first()
            check(t3 is not None and t3.ai_source == "rules" and "生産管理システム" in t3.inputs_list and t3.purpose == "不明" and t3.output_users and "社内システム" in t3.tools_list, "AI なしのときは手順から入力元・ツールを推定し、目的は不明にする")
            check(t3.monthly_minutes == 900 and "月間約 15.0 時間" in (t3.ai_summary or "") and t3.ai_priority == "高", f"ルールの分析は月間時間を出してから行う（要約に月間時間、多ければ優先度 高）: {t3.ai_summary}")
            check(t3.domain is not None and t3.domain.name == "製造実行・管理" and t3.ai_domain == t3.domain.name and "domain" in t3.ai_fields_list, "ルールで推定した業務領域を業務に入れ、推定の印を付ける")
            check(t3.inputs_list == ["生産管理システム"] and t3.nature_list == ["データ取得", "メール・配布"] and "転記・コピー" not in t3.nature_list and t3.pains_list == ["ミスが起きやすい"] and t3.actors == "",
                  f"入力元は取り出し元だけ（渡し方のメールは入れない）、作業の性質は手順の言葉から（転記に決めつけない）、困りごとは言葉から区分に、渡す相手を関わる人にしない: {t3.inputs_list} {t3.nature_list} {t3.pains_list}")
            check(set(t3.ai_fields_list) >= {"inputs", "output_users", "tools", "nature", "pains", "domain"} and set(t3.ai_extra_dict.get("rule_fields") or []) == set(t3.ai_fields_list), "ルールで埋めた項目はすべて推定の印の対象にする")
            check(t3.missing_required == ["アウトプット"] and "必須項目（アウトプット）を担当者と確認してください。" in t3.ai_ecrs_dict.get("questions", []), f"「必須項目…を確認」は本当に足りない項目だけを挙げる: {t3.ai_ecrs_dict.get('questions')}")
            words = app_module._content_words("日報の実績をExcelの集計表に転記する")
            check({"日報", "実績", "excel", "集計表", "転記"} <= words and not any("の" in w or "を" in w for w in words), f"関連業務の語は、ひらがな・助詞・記号で区切った意味のある語: {sorted(words)}")
            t_week = Task.query.filter_by(name="週報の作成").first()
            check(any(o.name == "日報の実績をExcelの集計表に転記する" for o in app_module.related_tasks_for(t_week)), "関連業務の候補は、語で区切ってつなぐ（週報 ⇔ 日次実績の転記）")

        client.post("/login", data={"emp_id": "admin", "password": "adminpass"})
        code, html, r = get(f"/cases/{tid}")
        check(code == 200 and "手順（聞き取り）" in html and "生産管理システムから前日の実績をCSVで出す" in html and "AI推定" in html and "改善後の流れ（AI案）" in html and "約 400 分/月" in html and "聞き取りの記録" in html and "はい、合っています" in html and "／夜勤）" in html, "案件詳細: 手順・AI推定の印・改善後の流れ・聞き取りの記録・勤務")
        check('id="def-task"' in html and "st1" in html and "st4" in html and "週報の作成" in html, "案件詳細の流れ図は手順の連鎖（前後の業務つき）")
        code, html, r = get("/cases/")
        check(code == 200 and '<span class="badge text-bg-dark">夜勤</span>' in html and '<span class="badge text-bg-dark">昼勤・夜勤の両方</span>' in html, "改善案件の一覧に夜勤・両方のバッジが出る")
        code, html, r = get("/people/")
        check(code == 200 and "<th>勤務形態</th>" in html and "交代勤務" in html, "DX人財情報に勤務形態の列が出る")
        check("検討: スキルを身につけながら、一緒にやってみたい" in html and "割ける時間" not in html and "発表・共有" not in html and "学びたい技術を挙げた人" in html, "DX人財情報の意欲は依頼の「一緒にやってみたい」から出し、集めなくなった列は出さない")
        code, html, r = get("/bizmap/")
        check(code == 200 and "日報の集計" in html, "業務マップに新しい案件が載る")
        code, html, r = get("/cases/?wish=まずは相談したい")
        check(code == 200 and "相談したい業務" in html, "改善案件の一覧を「まずは相談したい」で絞れる")
        csv_text = client.get("/cases/export.csv").data.decode("utf-8-sig")
        check("手順（聞き取り）" in csv_text and "1. 生産管理システムから前日の実績をCSVで出す［生産管理システム］" in csv_text and "4. 課長にメールで送る［メール］（課長へ）" in csv_text and ",夜勤," in csv_text and "AI推定の項目" in csv_text, "改善案件の CSV に勤務・手順（システム・渡す相手・渡し方）・AI推定の項目が出る")
        header = csv_text.splitlines()[0].split(",")
        check(all(h in header for h in ("分からなかったこと", "改善後の流れ（AI案）", "ECRSメモ 排除（Eliminate）", "ECRSメモ 簡素化（Simplify）", "改善内容・結果", "確認者")) and "CSV を Power Query で取り込む → 集計表が自動で更新される" in csv_text,
              "改善案件の CSV に分からなかったこと・改善後の流れ（AI案）・ECRS メモ・改善内容・確認者が出る")
        code, html, r = get(f"/cases/{tid}")
        check("AI の推定では 2人" in html and "月間時間には入れていません" in html, "案件詳細: AI が推定した人数は参考として出す")
        with app.app_context():
            t4_id = Task.query.filter_by(name="担当者の日報を写す").first().id
        code, html, r = get(f"/cases/{t4_id}")
        check('<span class="badge badge-flag me-1">推定</span>' in html, "案件詳細: ルールで埋めた項目には「推定」の印が付く")
        check("勤務形態" in client.get("/monthly/export.csv").data.decode("utf-8-sig") and "勤務形態" in client.get("/admin/export/respondents.csv").data.decode("utf-8-sig"), "月次・回答者の CSV に勤務形態が出る")
        client.post("/logout")
        r = client.delete(f"/survey/api/tasks/{tid}")
        check(r.status_code == 200 and r.get_json()["success"] and "日報の集計" in r.get_json()["suggested"], "取り消すと候補に戻る")
        check(not script, "用意した応答をすべて使った")
        check("確か<br/>" not in app_module.wrap("棚の在庫数を確かめて台帳に記入する") and "まと<br/>" not in app_module.wrap("集めた数字を資料にまとめて課長に渡す"), "流れ図の改行は語の途中（確か／める、まと／める）で切らない")
    finally:
        urllib.request.urlopen = real_urlopen

    client.post("/survey/api/reset")
    r = client.post("/survey/api/respondent", json={"emp_id": "", "name": "x", "workplace_id": ""})
    check(r.status_code == 400 and "emp_id" in r.get_json()["errors"], "社員IDが空ならエラー")
    r = client.post("/survey/api/respondent", json={"emp_id": "ZZZ9", "name": "名前 だけ", "workplace_id": 1, "work_pattern": "普通勤務"})
    check(r.status_code == 400 and "見つかりません" in r.get_json()["errors"].get("emp_id", ""), "見つからない社員IDは、氏名を入れても進めない")
    r = client.post("/survey/api/intake", json={"title": "x", "messages": [], "state": {}})
    check(r.status_code == 400, "あなたのことを保存していないと聞き取りできない")
    client.post("/survey/api/respondent", json={"emp_id": "T002", "name": "試験 次郎", "workplace_id": 1, "work_pattern": "普通勤務"})
    client.post("/login", data={"emp_id": "admin", "password": "adminpass"})
    client.post("/admin/ai", data={"openai_api_url": "", "openai_api_key": "", "openai_model": ""})
    client.post("/logout")
    r = client.post("/survey/api/intake", json={"title": "x", "messages": [], "state": {}})
    check(r.status_code == 400 and "AI連携が設定されていません" in r.get_json()["message"], "AI 未設定なら聞き取りは始められず、設定を案内する")

    return fails


def smoke_writeback():
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

    g = client.get("/bizmap/api/graph?kind=info").get_json()
    nodes = {n["id"]: n for n in g["nodes"]}
    with app.app_context():
        t1 = Task.query.filter_by(name="日報の実績をExcelの集計表に転記する").first()
        t2 = Task.query.filter_by(name="週報の作成").first()
        t3 = Task.query.filter_by(name="残業予定表の提出").first()
        t1_id, t2_id, t3_id = t1.id, t2.id, t3.id
        check(f"t{t1_id}->t{t2_id}" in {e["id"] for e in g["edges"]}, "auto seq edge t1->t2 exists")
        sys_node = next(nid for nid, n in nodes.items() if n.get("type") == "info" and n["text"] in ("生産管理システム", "品質システム", "設備データ"))
        ext_node = next(nid for nid, n in nodes.items() if n.get("type") == "work" and n["cls"] == "work" and n.get("actor") == "dept" and (n.get("who") or "") not in (t3.output_users_note or ""))
        src_node = next(nid for nid, n in nodes.items() if n.get("type") == "info" and n["cls"] == "source" and n["text"] not in ("生産管理システム", "品質システム", "設備データ"))

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
        {"from": f"t{t1_id}", "to": f"t{t3_id}"},
        {"from": sys_node, "to": f"t{t3_id}"},
        {"from": src_node, "to": f"t{t3_id}"},
        {"from": f"t{t3_id}", "to": ext_node},
        {"from": f"t{t3_id}", "to": "m_note1"},
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

def smoke_ai():
    import urllib.request

    import app as app_module

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

    URL = "https://ai.example.test/api/responses"

    AI_FORM = {"openai_api_url": URL, "openai_api_key": "test-key", "openai_model": "gpt-6-test"}
    conf.override("OPENAI_REASONING_EFFORT", "low")
    conf.override("OPENAI_MAX_OUTPUT_TOKENS", 4000)
    conf.override("OPENAI_API_TIMEOUT", 90)

    html = client.get("/admin/").data.decode("utf-8")
    with app.app_context():
        check(app_module.ai_enabled() is False and app_module._api_url() == "" and app_module._model() == "", "未設定のときは URL もモデル名も空のまま（既定の URL・モデル名を使わない）")
    check("URL・API キー・モデル名 が未設定" in html and 'name="openai_api_url"' in html and 'name="openai_api_key"' in html and 'name="openai_model"' in html, "管理画面に未設定の表示と、URL・API キー・モデル名の入力欄")
    r = client.post("/admin/ai", data=dict(AI_FORM, openai_model=""), follow_redirects=True)
    html = r.data.decode("utf-8")
    with app.app_context():
        check(app_module.ai_enabled() is False and app_module._api_url() == URL and app_module._model() == "", "URL とキーだけではまだ無効。モデル名は空のまま（既定のモデル名で補わない）")
    check("モデル名 が未設定" in html and "ルールベースで動きます" in html, "モデル名が空なら『モデル名 が未設定』と出る")
    r = client.post("/admin/ai", data=dict(AI_FORM, openai_api_url="  " + URL + "  "), follow_redirects=True)
    html = r.data.decode("utf-8")
    with app.app_context():
        check(app_module.ai_enabled() is True and app_module._api_url() == URL and app_module._model() == "gpt-6-test", "3つそろえば有効（前後の空白は除く）。再起動は不要")
    check("ChatGPT API 連携 有効" in html and "gpt-6-test" in html and "Responses API" in html and 'value="gpt-6-test"' in html and 'value="test-key"' in html, "管理画面に『ChatGPT API 連携 有効（モデル名、Responses API）』と保存した値")

    calls, script = [], []

    class _Resp:
        def __init__(self, payload):
            self.status = 200
            self._body = json.dumps(payload, ensure_ascii=False).encode("utf-8")

        def read(self):
            return self._body

        def __enter__(self):
            return self

        def __exit__(self, *exc):
            return False

    def fake_urlopen(req, timeout=None):
        raw = req.data.decode("utf-8")
        calls.append({"url": req.full_url, "method": req.get_method(), "headers": {k.lower(): v for k, v in req.header_items()}, "body": json.loads(raw), "raw": raw, "timeout": timeout})
        nxt = script.pop(0)
        if isinstance(nxt, Exception):
            raise nxt
        return _Resp(nxt)

    def message(text):
        return {"id": "resp_1", "object": "response", "status": "completed", "error": None, "output": [{"type": "reasoning", "id": "rs_1", "summary": []}, {"type": "message", "id": "msg_1", "status": "completed", "role": "assistant", "content": [{"type": "output_text", "text": text, "annotations": []}]}]}

    def http_error(code, body):
        return urllib.error.HTTPError(URL, code, "error", {}, io.BytesIO(body.encode("utf-8")))

    real_urlopen = urllib.request.urlopen
    urllib.request.urlopen = fake_urlopen
    try:
        client.post("/admin/demo/load")
        with app.app_context():
            task = Task.query.filter_by(name="週報の作成").first()
            domains = app_module.Domain.query.filter_by(parent_id=None).order_by(app_module.Domain.sort).all()
            cand_names = {t.name for t in app_module.related_tasks_for(task, limit=40)}
            other = next(t for t in Task.query.order_by(Task.id).all() if t.id != task.id and t.name not in cand_names)
            ai_json = json.dumps({"domain": domains[0].name, "related": [other.name, "存在しない業務"], "summary": "テスト要約（佐藤さんに確認）", "ecrs": {"eliminate": None, "combine": "結合のテスト", "rearrange": None, "simplify": None},
                                  "automation": "Python で自動化", "flags": ["自動化候補"], "priority": "高", "questions": ["確認1"], "tobe_steps": ["新しい流れ1"], "saving_minutes_per_month": 120, "headcount": 3}, ensure_ascii=False)
            minutes_before = task.monthly_minutes

            def refresh():
                app_module._finalize_task(task, app_module.task_state(task), task.intake, refresh=True)

            script.append(message("```json\n" + ai_json + "\n```"))
            refresh()
            c = calls[-1]
            check(len(calls) == 1 and c["url"] == URL and c["method"] == "POST", "config の URL に POST する")
            check(c["headers"].get("authorization") == "Bearer test-key" and c["headers"].get("content-type") == "application/json", "ヘッダーは Content-Type: application/json と Authorization: Bearer <キー>")
            b = c["body"]
            check(b["model"] == "gpt-6-test" and b["instructions"] and isinstance(b["input"], list) and task.name in b["input"][0]["content"] and b["max_output_tokens"] == 4000, "本文は model / instructions / input / max_output_tokens（Responses API）")
            check("業務分析アシスタント" in b["instructions"] and [x["name"] for x in b["tools"]] == ["search_similar_tasks", "duplicate_outputs"] and "登録済みの内容（再分析）" in b["input"][0]["content"] and "業務領域（推定）" in b["input"][0]["content"] and "回答者の選択" not in b["input"][0]["content"],
                  "再分析は登録と同じ分析（FINALIZE の指示と2つの道具）で、登録済みの内容（業務領域は「推定」）も送る")
            check(b.get("reasoning") == {"effort": "low"}, "reasoning.effort は config の OPENAI_REASONING_EFFORT")
            check("temperature" not in b and "top_p" not in b and "temperature" not in c["raw"] and "top_p" not in c["raw"] and "messages" not in b and "max_tokens" not in b, "temperature / top_p / max_tokens / messages は送らない")
            check(c["timeout"] == 90.0, "タイムアウトは config の OPENAI_API_TIMEOUT（既定 90 秒）")
            check(task.ai_source == "openai" and task.ai_domain == domains[0].name and task.ai_summary == "テスト要約（担当者に確認）" and task.ai_priority == "高" and task.ai_flags_list == ["自動化候補"] and task.ai_ecrs_dict["combine"] == "結合のテスト",
                  f"応答の JSON（```json 付き）を読んで結果に反映し、氏名は担当者に置き換える: {task.ai_summary}")
            check(task.ai_related_list == [other.name], f"関連業務は道具で見つかった実在の業務名も受け付け、ない名前は捨てる: {task.ai_related_list}")
            extra = task.ai_extra_dict
            check(extra.get("tobe_steps") == ["新しい流れ1"] and extra.get("saving_minutes_per_month") == 120 and extra.get("headcount_guess") == 3 and task.headcount == 1 and task.monthly_minutes == minutes_before,
                  "再分析で改善後の流れ・削減見込みを更新し、人数の推定は参考に残す（月間時間には掛けない）")
            check(task.domain.name == domains[0].name and "domain" in task.ai_fields_list and "domain" not in (extra.get("rule_fields") or []), "推定のままの業務領域は再分析の推定（AI）に置き換わる")

            script.append(http_error(400, '{"error":{"message":"Unsupported parameter: reasoning is not supported with this model.","param":"reasoning"}}'))
            script.append(message(ai_json))
            refresh()
            check(len(calls) == 3 and "reasoning" in calls[1]["body"] and "reasoning" not in calls[2]["body"] and task.ai_source == "openai", "reasoning が 400 で拒否されたら、外して再試行して成功")

            script.append(http_error(401, '{"error":{"message":"invalid api key"}}'))
            refresh()
            check(task.ai_source == "rules" and len(calls) == 4, "401 のときは再試行せずルールベースで補う")
            html = client.get("/admin/").data.decode("utf-8")
            check("直近のエラー" in html and "401" in html, "管理画面の AI連携に直近のエラー（HTTP 401）が出る")
            script.append({"id": "resp_r", "object": "response", "status": "completed", "error": None, "output": [{"type": "message", "id": "m", "status": "completed", "role": "assistant", "content": [{"type": "refusal", "refusal": "お答えできません"}]}]})
            refresh()
            check(task.ai_source == "rules", "拒否（refusal）のときもルールベースで補う")
            script.append(urllib.error.URLError("接続できない（テスト）"))
            refresh()
            check(task.ai_source == "rules", "接続エラーのときもルールベースで補う")
            script.append(message("これは JSON ではありません"))
            refresh()
            check(task.ai_source == "rules" and task.ai_extra_dict.get("tobe_steps") == ["新しい流れ1"], "JSON でない応答のときもルールベースで補う（前の改善案は消さない）")
            task_id = task.id

        script.append(message(ai_json))
        r = client.post(f"/cases/{task_id}/reanalyze", follow_redirects=True)
        html = r.data.decode("utf-8")
        check(r.status_code == 200 and "AI分析を更新しました（ChatGPT API）" in html and ">ChatGPT API</span>" in html and "テスト要約" in html, "案件詳細の再分析: ChatGPT API のバッジと結果")
        check("改善後の流れ（AI案）" in html and "新しい流れ1" in html and "AI の推定では 3人" in html, "再分析の改善後の流れと、人数の推定（参考）が案件詳細に出る")
        html = client.get("/admin/").data.decode("utf-8")
        check("直近のエラー" not in html, "成功すると直近のエラーは消える")
        script.append(message(json.dumps([{"kind": "調査", "title": "テスト提案のタイトル", "detail": "根拠 12 件"}], ensure_ascii=False)))
        r = client.post("/ai/suggest", follow_redirects=True)
        html = r.data.decode("utf-8")
        check(r.status_code == 200 and "テスト提案のタイトル" in html, "ダッシュボードの『AIで提案を生成』も Responses API 経由で出る")
        check("temperature" not in calls[-1]["raw"] and calls[-1]["body"].get("reasoning") == {"effort": "low"}, "提案生成も temperature なし・reasoning あり")
        def tool_call_message(name, args, call_id="call_1"):
            return {"id": "resp_t", "object": "response", "status": "completed", "error": None, "output": [{"type": "function_call", "id": "fc_1", "call_id": call_id, "name": name, "arguments": json.dumps(args, ensure_ascii=False), "status": "completed"}]}

        got = []
        tools = [{"type": "function", "function": {"name": "echo", "description": "test", "parameters": {"type": "object", "properties": {"q": {"type": "string"}}, "required": ["q"]}}}]
        handlers = {"echo": lambda args: got.append(args) or {"rows": [{"a": 1}]}}
        with app.app_context():
            script.append(tool_call_message("echo", {"q": "x"}))
            script.append(message("done"))
            text = app_module._call_ai("sys", "hello", tools=tools, handlers=handlers)
            b1, b2 = calls[-2]["body"], calls[-1]["body"]
            check(text == "done" and got == [{"q": "x"}], "道具の依頼（function_call）を実行して最終回答を得る")
            check(b1["tools"] == [{"type": "function", "name": "echo", "description": "test", "parameters": tools[0]["function"]["parameters"]}] and b1["tool_choice"] == "auto" and b1["store"] is False and b1["input"] == [{"role": "user", "content": "hello"}] and "temperature" not in calls[-2]["raw"], "tools は平らな形、tool_choice は auto、store は false、temperature なし（aiagent_rag_bundle と同じ送り方）")
            check(b2["input"][-2] == {"type": "function_call", "call_id": "call_1", "name": "echo", "arguments": json.dumps({"q": "x"}, ensure_ascii=False)} and b2["input"][-1]["type"] == "function_call_output" and b2["input"][-1]["call_id"] == "call_1" and json.loads(b2["input"][-1]["output"]) == {"rows": [{"a": 1}]}, "2往復目の input に function_call と function_call_output（JSON 文字列）を積む")
            got.clear()
            script.append(tool_call_message("echo", {"q": "x"}, "c1"))
            script.append(tool_call_message("echo", {"q": "x"}, "c2"))
            script.append(tool_call_message("echo", {"q": "x"}, "c3"))
            script.append(message("final"))
            text = app_module._call_ai("sys", "hello", tools=tools, handlers=handlers)
            outs = [i for i in calls[-1]["body"]["input"] if i.get("type") == "function_call_output"]
            check(text == "final" and got == [{"q": "x"}] and len(outs) == 3 and "同じ呼び出し" in outs[1]["output"] and "tools" not in calls[-1]["body"] and calls[-1]["body"]["input"][-1]["role"] == "user", "同じ呼び出しは2回目から実行せずに差し戻し、続くなら道具なしで最終回答をまとめる")
            script.append(tool_call_message("nope", {}, "c9"))
            script.append(message("ok"))
            text = app_module._call_ai("sys", "hello", tools=tools, handlers=handlers)
            check(text == "ok" and "未知の道具" in calls[-1]["body"]["input"][-1]["output"], "未知の道具は error として返し、AI に続けさせる")
            script.append(tool_call_message("echo", {}, "c8"))
            script.append(message("ok2"))
            text = app_module._call_ai("sys", "hello", tools=tools, handlers={"echo": lambda args: (_ for _ in ()).throw(ValueError("q を指定してください"))})
            check(text == "ok2" and "q を指定してください" in calls[-1]["body"]["input"][-1]["output"], "道具の失敗は理由つきの error として返す")
            script.append(http_error(400, '{"error":{"message":"Unknown parameter: tools","param":"tools"}}'))
            script.append(message("plain"))
            text = app_module._call_ai("sys", "hello", tools=tools, handlers=handlers)
            check(text == "plain" and "tools" in calls[-2]["body"] and "tools" not in calls[-1]["body"] and "tool_choice" not in calls[-1]["body"], "中継が tools を拒否したら、道具なしでやり直して答えを得る")
            script.append(message("plain2"))
            n0 = len(calls)
            text = app_module._call_ai("sys", "hello", tools=tools, handlers=handlers)
            check(text == "plain2" and len(calls) == n0 + 1 and "tools" not in calls[-1]["body"], "一度拒否されたら、以後は最初から道具なしで送る（1往復で済む）")

        script.append(message("OK"))
        script.append(tool_call_message("ping", {}, "p1"))
        script.append(message("OK"))
        r = client.post("/admin/ai/test", data=AI_FORM, follow_redirects=True)
        html = r.data.decode("utf-8")
        check(r.status_code == 200 and "接続できました" in html and "道具の呼び出し（tool calling）も通りました" in html and calls[-2]["body"]["tool_choice"] == "auto" and "tools" in calls[-2]["body"] and "temperature" not in calls[-1]["raw"], "接続テスト: 保存し直すと道具を試し直し、応答と道具の呼び出しが通ることを確認する")
        script.append(http_error(404, "not found"))
        r = client.post("/admin/ai/test", data=AI_FORM, follow_redirects=True)
        html = r.data.decode("utf-8")
        check("接続できませんでした" in html and "404" in html, "接続テスト: 失敗したら理由を表示")
        script.append(message("OK"))
        script.append(message("道具は呼びません"))
        r = client.post("/admin/ai/test", data=AI_FORM, follow_redirects=True)
        html = r.data.decode("utf-8")
        check("道具の呼び出しは通りませんでした" in html, "接続テスト: 道具が呼ばれなければその旨を表示")
        r = client.post("/admin/ai", data={"openai_api_url": "", "openai_api_key": "", "openai_model": ""}, follow_redirects=True)
        html = r.data.decode("utf-8")
        with app.app_context():
            check(app_module.ai_enabled() is False and app_module._api_url() == "" and app_module._api_key() == "" and app_module._model() == "", "空で保存すると空に戻る（何も補わない）")
        check("URL・API キー・モデル名 が未設定" in html, "空に戻すと未設定の表示")
        check(not script, "用意した応答をすべて使った")
    finally:
        urllib.request.urlopen = real_urlopen

    return fails

TESTS = [("smoke_bizmap_edit", smoke_bizmap_edit), ("smoke_board", smoke_board), ("smoke_errors", smoke_errors), ("smoke_login_monthly", smoke_login_monthly), ("smoke_request", smoke_request), ("smoke_writeback", smoke_writeback), ("smoke_ai", smoke_ai)]


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
