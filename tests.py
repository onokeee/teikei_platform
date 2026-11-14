"""スモークテスト。python tests.py で全部、python tests.py smoke_board のように名前を付けると1本。一時ファイルの SQLite を使う"""
import csv
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

    # ---- 業務マップと他の機能のつながり（相手・受け取り元・相手の箱の ID・保留と廃止・案件の流れ図・図の下の表・月次の業務名） ----
    with app.test_request_context():
        by_name = {t.name: t for t in Task.query.all()}
        t_day, t_week, t_asm, t_def = by_name["日報の実績をExcelの集計表に転記する"], by_name["週報の作成"], by_name["組立実績の日次入力（生産管理システム）"], by_name["不良品発生報告のメール配信"]
        wp1 = t_day.workplace_id
        check(app_module._names("オペレーターの故障連絡（口頭・メモ）", dots=True) == ["オペレーターの故障連絡（口頭・メモ）"] and app_module._names("週報・生産管理システム", dots=True) == ["週報", "生産管理システム"]
              and any(s_["name"] == "オペレーターの故障連絡（口頭・メモ）" for s_ in app_module.task_sources(by_name["設備故障記録の台帳入力"])),
              "かっこの中の「・」では分けない（「オペレーターの故障連絡（口頭・メモ）」は1つの受け取り元）、かっこの外は分ける")
        tr = Task(name="相手の確かめ用の業務", respondent_id=t_day.respondent_id, workplace_id=wp1, purpose="朝会で確認するため", output_users_note="検査・品質保証課 課長、製造1課 課長",
                  steps_json=json.dumps([{"no": 1, "action": "集計表を課長に渡す", "output": "集計表", "to": "課長"}], ensure_ascii=False))
        db.session.add(tr)
        db.session.flush()
        labels = [r_["who"]["label"] for r_ in app_module.task_recipients(tr)]
        check(labels == ["製造1課 課長", "検査・品質保証課 課長"], f"渡す相手はだれが同じものだけまとめる（「課長」は「製造1課 課長」と同じ、「検査・品質保証課 課長」は別）: {labels}")
        boss_id = app_module.nid("p", f"{tr.id}:課長")
        d1 = app_module.build_info_flow([tr])
        raw_vocab = app_module.Setting.get(app_module.MAP_VOCAB_KEY, "")
        v = app_module.vocab_public()
        for row_ in v["actors"]:
            if row_["key"] == "boss":
                row_["words"] = "上司、部長"
        app_module.Setting.set(app_module.MAP_VOCAB_KEY, json.dumps(v, ensure_ascii=False))
        tr.purpose, tr.output_users_note = "目的を書き直した", ""
        d2 = app_module.build_info_flow([tr])
        check(boss_id in d1.nodes and boss_id in d2.nodes and d1.nodes[boss_id]["who"] == "製造1課 課長" and d2.nodes[boss_id]["who"] == "課長" and d2.nodes[boss_id]["text"] == "目的を書き直した",
              f"相手の箱の ID は業務と答えた名前で決まる（区分の語・目的・相手の数を変えても同じ箱で、手修正が残る）: {d2.nodes.get(boss_id)}")
        app_module.Setting.set(app_module.MAP_VOCAB_KEY, raw_vocab)
        db.session.rollback()

        tasks = Task.query.all()
        flow = app_module.build_info_flow(tasks)
        src = flow.render()
        legend = app_module.map_legend_html(flow)
        check(flow.nodes[f"t{t_asm.id}"]["cls"] == "hold" and flow.nodes[f"t{t_def.id}"]["cls"] == "abolished" and "（保留）" in src and "（廃止済み）" in src
              and "classDef hold" in src and "stroke-dasharray:5 4" in src and "保留中の業務" in legend and "廃止した業務" in legend and flow.nodes[f"t{t_asm.id}"]["text"] == t_asm.name,
              "業務マップ: 保留の業務は点線の枠に「（保留）」、廃止して完了した業務は薄い色に「（廃止済み）」（名前そのものは変えない）")
        fa, fd = app_module.build_task_flow(t_asm), app_module.build_task_flow(t_def)
        fd_nodes = [n for i, n in fd.nodes.items() if i.startswith("st") or i == f"t{t_def.id}"]
        check(all(n["cls"] == "hold" for i, n in fa.nodes.items() if i.startswith("st")) and "（保留）" in fa.render() and fd_nodes and all(n["cls"] == "abolished" for n in fd_nodes) and "（廃止済み）" in fd.render(),
              "案件の流れ図でも、保留は点線の枠に「（保留）」、廃止済みは薄い色に「（廃止済み）」")
        f_day = app_module.build_task_flow(t_day)
        roles = {n["text"]: n["cls"] for n in f_day.nodes.values() if n["type"] == "info"}
        check(roles.get("Excel（日次実績集計表）") == "output" and roles.get("紙の日報") == "source", f"案件の流れ図も業務マップと同じ色の決まり（手順が書き込む情報はできるもの）: {roles}")
        tn = Task(name="前後の確かめ用の業務", respondent_id=t_day.respondent_id, workplace_id=wp1, prev_task="日報の実績をExcelの集計表に転記する、各ラインの日報記入\n各ラインの日報記入（紙）",
                  next_task="週報の作成、生産実績の月報作成", steps_json=json.dumps([{"no": 1, "action": "数字を確かめる"}], ensure_ascii=False))
        db.session.add(tn)
        db.session.flush()
        fn = app_module.build_task_flow(tn)
        prev_boxes = sorted(fn.nodes[a]["text"] for (a, b), e in fn.edges.items() if b == "st1" and e["type"] == "work")
        nexts = {fn.nodes[b]["text"]: e["handoff"] for (a, b), e in fn.edges.items() if a == "st1" and e["type"] == "work"}
        monthly = by_name["生産実績の月報作成"]
        check(prev_boxes == sorted(["日報の実績をExcelの集計表に転記する", "各ラインの日報記入"]) and nexts.get("週報の作成") == "none"
              and nexts.get("生産実績の月報作成") == app_module._task_handoff(tn, monthly, fn),
              f"案件の流れ図: 前後の業務は1業務1つの箱（同じもの＝same_thing はまとめる）、引き継ぎ方は業務マップと同じ決まり（同じ人なら引き継ぎなし）: {prev_boxes} {nexts}")
        db.session.rollback()

        rows = app_module.handoff_rows(tasks)
        pairs = {(r_["src"], r_["dst"]): r_ for r_ in rows}
        wp1_name = t_day.workplace.name
        check(("オペレーターの故障連絡（口頭・メモ）", "設備保全課") in pairs and ("受注一覧（Excel）", "生産管理課") in pairs and ("生産管理システム", "生産管理課") in pairs
              and (wp1_name, "総務") in pairs and (wp1_name, "課長") not in pairs and pairs[(wp1_name, f"{wp1_name} 課長")]["info"] >= 1,
              "図の下の表も業務マップと同じ決まり（受け取り元は1つずつ、「他部署」は目的の部署名、「課長」は「製造1課 課長」にまとめる）")

        db.session.add(MonthlyResponse(period=current_period(), emp_id="P201", name="", workplace_id=wp1,
                                       answers_json=json.dumps({"has": "ある", "task_names": "月報の作成\n週報の作成、日報の集計・確認\n佐藤さんの日報まとめ\n週報の作成"}, ensure_ascii=False)))
        db.session.add(MonthlyResponse(period=current_period(), emp_id="P205", name="", workplace_id=wp1, answers_json=json.dumps({"has": "ある", "task_names": "日報の集計・確認"}, ensure_ascii=False)))
        db.session.add(Employee(emp_id="P202", name="対象外の人", active=False, source="answer"))
        db.session.add(MonthlyResponse(period=current_period(), emp_id="P202", name="", workplace_id=wp1, answers_json=json.dumps({"has": "ある", "task_names": "対象外の人の業務"}, ensure_ascii=False)))
        db.session.add(MonthlyResponse(period=shift_period(current_period(), -6), emp_id="P203", name="", workplace_id=wp1, answers_json=json.dumps({"has": "ある", "task_names": "7か月前の業務"}, ensure_ascii=False)))
        other_wp = by_name["組立実績の日次入力（生産管理システム）"].workplace_id
        db.session.add(MonthlyResponse(period=shift_period(current_period(), -1), emp_id="P204", name="", workplace_id=wp1, answers_json=json.dumps({"has": "ある", "task_names": "異動前の業務"}, ensure_ascii=False)))
        db.session.add(MonthlyResponse(period=current_period(), emp_id="P204", name="", workplace_id=other_wp, answers_json=json.dumps({"has": "ある", "task_names": "異動後の業務"}, ensure_ascii=False)))
        db.session.commit()
        pend = {p["name"]: len(p["people"]) for p in app_module.pending_task_names(wp1)}
        pend_other = [p["name"] for p in app_module.pending_task_names(other_wp)]
        check("月報の作成" in pend and "週報の作成" not in pend and "日報の集計・確認" in pend and "担当者の日報まとめ" in pend and not any("佐藤" in p for p in pend),
              f"月次の業務名: 1行・「、」ごとに分け（「・」では分けない）、「月報の作成」と「週報の作成」は別、氏名は担当者に置き換える: {sorted(pend)}")
        check("対象外の人の業務" not in pend and "7か月前の業務" not in pend and "異動前の業務" not in pend and "異動後の業務" in pend_other
              and app_module.name_key("ＡＢＣ点検") == app_module.name_key("abc点検"),
              f"月次の業務名: 各人の最新の回答（6か月以内）を職場で絞り、対象外の人は除く。全角・半角はそろえる: {sorted(pend)} / {pend_other}")
        flow = app_module.add_pending_tasks(app_module.build_info_flow(tasks), wp1, tasks)
        u_month = app_module.nid("u", f"{wp1}:月報の作成")
        check((f"t{t_week.id}", u_month) in flow.edges and (f"t{t_day.id}", u_month) not in flow.edges,
              "未登録の業務への点線も、前後の業務の名前が同じときだけ（「週報の作成」は「月報の作成」につながない）")

    g_wp = client.get(f"/bizmap/api/graph?kind=info&workplace_id={wp1}").get_json()
    u_node = next(n for n in g_wp["nodes"] if n["cls"] == "pending" and n["text"] == "日報の集計・確認（2人）")
    data = {"nodes": [dict(n) for n in g_wp["nodes"]] + [{"id": "m_who", "text": "佐藤課長に確認", "type": "work", "cls": "work", "who": "山田さん", "who_edited": True, "actor": "dept", "actor_locked": True, "source": "manual"}],
            "edges": [dict(e) for e in g_wp["edges"]] + [{"id": f"t{t_day.id}->{u_node['id']}", "from": f"t{t_day.id}", "to": u_node["id"], "type": "work", "source": "manual"},
                                                         {"id": f"t{t_day.id}->m_who", "from": f"t{t_day.id}", "to": "m_who", "type": "work", "source": "manual", "label": "鈴木さんに電話", "label_edited": True}],
            "suppressed_nodes": [], "suppressed_edges": [], "view": None}
    r = client.post("/bizmap/api/graph", json={"kind": "info", "workplace_id": wp1, "data": data})
    with app.app_context():
        stored = json.loads(FlowMap.query.filter_by(scope=f"wp:{wp1}|dom:0").first().data)
        m_who = next(n for n in stored["nodes"] if n["id"] == "m_who")
        e_who = next(e for e in stored["edges"] if e["to"] == "m_who")
    check(r.status_code == 200 and m_who["who"] == "担当者" and m_who["text"] == "課長に確認" and e_who["label"] == "担当者に電話",
          f"保存するときにサーバーで氏名を置き換える（だれ・箱の名前・線のラベル。だれの区分を手で選んでいても）: {m_who['who']} / {m_who['text']} / {e_who['label']}")
    with app.app_context():
        db.session.add(Task(name="日報の集計・確認", respondent_id=Respondent.query.filter_by(emp_id="DEMO-001").first().id, workplace_id=wp1, frequency="毎日", duration="15分未満"))
        db.session.commit()
        new_id = Task.query.filter_by(name="日報の集計・確認").first().id
    g_wp = client.get(f"/bizmap/api/graph?kind=info&workplace_id={wp1}").get_json()
    eids = {e["id"] for e in g_wp["edges"]}
    new_box = next((n for n in g_wp["nodes"] if n["id"] == f"t{new_id}"), {})
    check(u_node["id"] not in {n["id"] for n in g_wp["nodes"]} and f"t{t_day.id}->t{new_id}" in eids and new_box.get("text") == "日報の集計・確認" and not new_box.get("text_edited"),
          "未登録の箱（2人）に引いた手の線は、同じ職場に同じ名前の業務が登録されたら、その業務の箱につなぎ直す（業務名は変えない）")
    with app.app_context():
        steps_only = app_module.Workplace(name="手順だけの職場", sort=99)
        db.session.add(steps_only)
        db.session.flush()
        db.session.add(Task(name="手順だけの業務", respondent_id=Respondent.query.filter_by(emp_id="DEMO-001").first().id, workplace_id=steps_only.id,
                            steps_json=json.dumps([{"no": 1, "action": "集計する"}], ensure_ascii=False)))
        db.session.commit()
        steps_wp = steps_only.id
    html = client.get(f"/bizmap/?workplace_id={steps_wp}").data.decode("utf-8")
    check("図に載せた業務 1 件（関わる人・情報の受け取り元の入力あり 1 件）" in html, "見出しの件数: 聞き取りの手順がある業務は入力ありに数える")
    html = client.get(f"/bizmap/?workplace_id={wp1}").data.decode("utf-8")
    check(html.count('border align-middle">全体</span>') >= 4, "絞り込みに関係なく全体で数える表には「全体」と書く")
    html = client.get("/bizmap/edit?kind=info").data.decode("utf-8")
    fe = client.get("/static/flowedit.js").data.decode("utf-8")
    check('"status_marks"' in html and "fe-state" in fe and "else if (it.who && it.who !== n.who)" in fe and 'if (n.actor_locked) renderAll();\n    classify([n]);' in fe,
          "編集: 保留・廃止済みの印を出し、だれの区分を手で選んでいても名前はサーバーで氏名を置き換える")

    # ---- 案件の流れ図を業務マップにそろえる（ほかの業務が書いた前後の業務・同じファイル・完了の色・名前のない取り出し元） ----
    with app.test_request_context():
        by_name = {t.name: t for t in Task.query.all()}
        t_day, t_week, t_cnt, t_month = by_name["日報の実績をExcelの集計表に転記する"], by_name["週報の作成"], by_name["出勤簿の集計"], by_name["生産実績の月報作成"]
        full = app_module.build_info_flow(list(by_name.values()))
        f_week = app_module.build_task_flow(t_week)
        last = f"st{t_week.steps[-1]['no']}"
        after = {f_week.nodes[b]["text"]: e["inferred"] for (a, b), e in f_week.edges.items() if a == last and e["type"] == "work"}
        before = {f_week.nodes[a]["text"]: e["inferred"] for (a, b), e in f_week.edges.items() if b == "st1" and e["type"] == "work"}
        in_full = {full.nodes[b]["text"] for (a, b), e in full.edges.items() if a == f"t{t_week.id}" and e["type"] == "work"}
        check({t_month.name, "課内会議資料の作成"} <= in_full <= set(after) and after[t_month.name] is True and after["課内会議資料の作成"] is True
              and before.get(t_day.name) is True and after.get("月報の作成") is False,
              f"案件の流れ図: ほかの業務が前後の業務にこの業務を書いていれば、業務マップと同じくその業務の箱と点線でつなぐ（自分の前後の業務は実線のまま）: {after} / {before}")
        f_day, f_mon = app_module.build_task_flow(t_day), app_module.build_task_flow(t_month)
        day_files = [i for i, n in f_day.nodes.items() if n["type"] == "info" and "日次実績集計表" in n["text"]]
        writers = sorted(a for (a, b) in f_day.edges if day_files and b == day_files[0])
        mon_files = sorted(n["text"] for n in f_mon.nodes.values() if n["type"] == "info" and "月報" in n["text"])
        bare = app_module.build_steps_flow([{"no": 1, "action": "受注一覧から数字を拾う", "system": "Excel（受注一覧）"}, {"no": 2, "action": "数字をExcelに入力する", "system": "Excel", "output": "集計表"}])
        bare_writes = sorted(bare.nodes[b]["text"] for (a, b) in bare.edges if a == "st2" and bare.nodes[b]["type"] == "info")
        check(len(day_files) == 1 and f_day.nodes[day_files[0]]["text"] == "Excel（日次実績集計表）" and f_day.nodes[day_files[0]]["cls"] == "output" and writers == ["st2", "st3"]
              and mon_files == sorted(["Excel（月報）", "Excel（月報集計表）"]) and bare_writes == sorted(["Excel", "集計表"]),
              f"案件の流れ図: 手順が書き込むファイルと同じできるもの（「Excel（日次実績集計表）」と「日次実績集計表」）は1つの箱、「Excel（月報）」と「Excel（月報集計表）」・媒体名だけの「Excel」は別: {[f_day.nodes[i]['text'] for i in day_files]} {writers} {mon_files} {bare_writes}")
        f_cnt = app_module.build_task_flow(t_cnt)
        check(t_cnt.status == "完了" and full.nodes[f"t{t_cnt.id}"]["cls"] == "done" and {n["cls"] for i, n in f_cnt.nodes.items() if i.startswith("st")} == {"done"}
              and ":::done" in f_cnt.render() and "完了した業務" in app_module.map_legend_html(f_cnt),
              "案件の流れ図: 完了した業務の手順は業務マップと同じ緑（完了した業務）で描く")
        r_ = Respondent.query.filter_by(emp_id="DEMO-007").first()
        mk = lambda name, *steps: Task(name=name, respondent_id=r_.id, workplace_id=r_.workplace_id, inputs="人から聞く", steps_json=json.dumps(list(steps), ensure_ascii=False))
        ledger = {"no": 2, "action": "不具合台帳に入力する", "system": "Excel（不具合台帳）", "output": "不具合台帳"}
        k_tie = mk("PHSの連絡を台帳に残す", {"no": 1, "action": "オペレーターから不具合の連絡を受ける", "how_from": "PHSで聞く"}, ledger)
        k_none = mk("台帳に入力するだけの業務", dict(ledger, no=1))
        k_has = mk("PHSで内容を聞いて台帳に残す", {"no": 1, "action": "オペレーターから不具合の内容を聞く", "input": "不具合の内容", "how_from": "PHSで聞く"}, ledger)
        db.session.add_all([k_tie, k_none, k_has])
        db.session.flush()

        def unnamed(t):
            f = app_module.build_task_flow(t)
            return [(b, e["method"]) for (a, b), e in f.edges.items() if f.nodes[a]["text"].endswith("（名前未確認）")]
        check(unnamed(k_tie) == [("st1", "phs")] and unnamed(k_none) == [] and unnamed(k_has) == [],
              f"案件の流れ図: 区分しか分からない取り出し元は、受け取り方が合う手順（PHSで聞く）につなぎ、合う手順がない・その手順に取り出し元があるときは描かない: {unnamed(k_tie)} {unnamed(k_none)} {unnamed(k_has)}")
        full_k = app_module.build_info_flow([k_tie, k_none, k_has])
        unnamed_full = sorted((b, e["method"]) for (a, b), e in full_k.edges.items() if full_k.nodes[a]["text"].endswith("（名前未確認）"))
        rows_k = {r_["name"]: [x.name for x in r_["tasks"]] for r_ in app_module.source_rows([k_tie, k_none, k_has])}
        check(unnamed_full == [(f"t{k_tie.id}", "phs")] and rows_k.get("人から聞く内容（名前未確認）") == [k_tie.name]
              and [s_["name"] for s_ in app_module.task_sources(k_none)] == [] and [(s_["name"], s_["method"]) for s_ in app_module.task_sources(k_tie)] == [("人から聞く内容（名前未確認）", "phs")],
              f"R3-5: 業務マップ全体・システム・情報源の表も案件の流れ図と同じ: 手順のある業務の区分だけの取り出し元は、受け取り方が合う手順があるときだけ（その受け取り方で）描き、ないときは描かない: {unnamed_full} {rows_k}")
        eq_steps = [{"no": 1, "action": "設備モニタで稼働状況を見る", "system": "設備モニタ"}, {"no": 2, "action": "稼働日報をExcelで作る", "system": "Excel（稼働日報）", "output": "稼働日報"}]
        k_eq = Task(name="設備の稼働日報の作成", respondent_id=r_.id, workplace_id=r_.workplace_id, inputs="設備データ", output="稼働日報", steps_json=json.dumps(eq_steps, ensure_ascii=False))
        k_old = Task(name="設備モニタの値を書き写す", respondent_id=r_.id, workplace_id=r_.workplace_id, inputs="設備データ|生産管理システム", inputs_note="設備モニタ", steps_json="[]")
        db.session.add_all([k_eq, k_old])
        db.session.flush()
        infos = lambda f: sorted(n["text"] for n in f.nodes.values() if n["type"] == "info" and n["cls"] == "source")
        rows_eq = sorted(r_["name"] for r_ in app_module.source_rows([k_eq, k_old]))
        check(infos(app_module.build_info_flow([k_eq])) == ["設備モニタ"] and infos(app_module.build_task_flow(k_eq)) == ["設備モニタ"]
              and infos(app_module.build_info_flow([k_old])) == ["生産管理システム", "設備モニタ"] and infos(app_module.build_task_flow(k_old)) == ["生産管理システム", "設備モニタ"]
              and rows_eq == ["生産管理システム", "設備モニタ"],
              f"R3-5: 名前のある取り出し元から分類した区分（設備モニタ → 設備データ）は、業務マップ・案件の流れ図・システム・情報源の表で別の箱にしない（ほかの区分は描く）: {rows_eq}")
        db.session.rollback()

    # ---- 職場名を変えても、前の名前で書いた相手・前後の業務はその職場の枠とだれになる ----
    with app.app_context():
        wp_pm = app_module.Workplace.query.filter_by(name="生産管理課").first().id
    html_wp = client.post(f"/admin/workplace/{wp_pm}/update", data={"name": "生産計画課"}, follow_redirects=True).data.decode("utf-8")
    with app.test_request_context():
        t_asm = Task.query.filter_by(name="組立実績の日次入力（生産管理システム）").first()
        old_who = app_module.who_of("生産管理課 課長", "製造1課")
        rec = app_module.build_info_flow([t_asm]).nodes.get(app_module.nid("p", f"{t_asm.id}:生産管理課"), {})
        case = app_module.build_task_flow(t_asm)
        merged = [n["text"] for n in case.nodes.values() if n["cls"] == "work" and n["who"] == "生産計画課"]
        rows = {(r_["src"], r_["dst"]) for r_ in app_module.handoff_rows([t_asm])}
        check("前の名前で書かれた" in html_wp and (old_who["actor"], old_who["label"]) == ("dept", "生産計画課 課長") and rec.get("who") == "生産計画課" and rec.get("actor") == "dept"
              and rec.get("group") == f"wp{wp_pm}" and merged == ["進捗確認"] and ("組立課", "生産計画課") in rows,
              f"職場名の変更: 前の名前（生産管理課）で書いた相手はいまの職場の枠とだれ、後の業務「生産管理課の進捗確認」も相手の箱にまとまる: {old_who} / {rec.get('who')} {rec.get('group')} / {merged}")
    client.post(f"/admin/workplace/{wp_pm}/update", data={"name": "生産管理課"})
    client.post("/admin/workplace/add", data={"name": "生産計画課"})
    with app.test_request_context():
        new_wp = app_module.Workplace.query.filter_by(name="生産計画課").first().id
        names_now = dict(app_module._workplace_names())
        stored = app_module.workplace_aliases().get(str(wp_pm))
    client.post(f"/admin/workplace/{new_wp}/delete")
    with app.test_request_context():
        names_after = dict(app_module._workplace_names())
    check(stored == ["生産計画課"] and names_now.get("生産計画課") == "生産計画課" and names_now.get("生産管理課") == "生産管理課" and names_after.get("生産計画課") == "生産管理課",
          f"職場名の変更: 名前を戻すと戻した名前は前の名前から外れ、前の名前と同じ職場がほかにあればいまの職場名を優先する: {stored} {names_now.get('生産計画課')} {names_after.get('生産計画課')}")

    # ---- 職場名の変更: 前の名前は、名前として独立して出てくるときだけいまの名前に直す（「第二組立課」は直さない） ----
    with app.app_context():
        wp_asm = app_module.Workplace.query.filter_by(name="組立課").first().id
    client.post(f"/admin/workplace/{wp_asm}/update", data={"name": "組立1課"})
    with app.test_request_context():
        texts = ("第二組立課", "組立課 課長", "組立課長", "組立課の進捗確認", "製造1課から組立課へ", "組立課題の整理", "組立課二係")
        swapped = {s: app_module.current_wp_text(s) for s in texts}
        other = app_module.who_of("第二組立課", "製造1課")
        mine = app_module.who_of("組立課 課長", "製造1課")
        d_ = app_module.Diagram()
        frame_other = d_.nodes[app_module.recipient_node(d_, "第二組立課", "製造1課", frame=("wp1", "製造1課"))[0]]["group"]
        frame_mine = d_.nodes[app_module.recipient_node(d_, "組立課の班長", "製造1課", frame=("wp1", "製造1課"))[0]]["group"]
        dept_other, dept_mine = app_module.dept_in_text("第二組立課に伝えるため", "製造1課"), app_module.dept_in_text("組立課に伝えるため", "製造1課")
    client.post(f"/admin/workplace/{wp_asm}/update", data={"name": "組立課"})
    check(swapped == {"第二組立課": "第二組立課", "組立課 課長": "組立1課 課長", "組立課長": "組立1課長", "組立課の進捗確認": "組立1課の進捗確認", "製造1課から組立課へ": "製造1課から組立1課へ",
                      "組立課題の整理": "組立課題の整理", "組立課二係": "組立課二係"}
          and other["label"] == "第二組立課" and frame_other != f"wp{wp_asm}" and (mine["actor"], mine["label"]) == ("dept", "組立1課 課長") and frame_mine == f"wp{wp_asm}"
          and dept_other == "第二組立課" and dept_mine == "組立1課",
          f"R3-4: 職場名の変更（組立課 → 組立1課）: 前の名前は名前として独立して出てくるときだけ（後ろの役職はよい）いまの名前に直し、「第二組立課」のような長い名前の一部は直さず、その職場の枠にも入れない: {swapped} / {other} {frame_other}")

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
        own_survey = app_module.get_or_create_survey(current_period()).id
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
        sample = by["日報の実績をExcelの集計表に転記する"]
        sample_id, sample_note = sample.id, sample.ai_plan_note
    code, html = page(f"/cases/{irregular_id}")
    check("不定期（月1回として計算）" in html, "案件詳細: 不定期は「月1回として計算」と出す")
    code, html = page(f"/cases/{sample_id}")
    check(sample_note == "デモ" and '<span class="badge bg-warning-subtle text-warning-emphasis">デモ</span> デモ用の見本の案です' in html and "改善後の流れ（AI案）" in html and ">ルールベース</span>" in html,
          "デモ: ルールで分析した業務に付けた見本の改善案（AI案）には「デモ」の印を付ける")

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
        wk_summary, wk_at, wk_edited = wk.ai_summary or "", wk.ai_at, wk.ai_extra_dict.get("info_edited_at")
        dom_q = app_module.Domain.query.filter_by(name="品質管理").first().id
    html_up = client.post(f"/cases/{wk_id}/update", data=dict(form, prev_task="日報の実績をExcelの集計表に転記する", next_task="月報の作成（佐藤さんへ）", actors="佐藤課長、生産管理課", domain_id=str(dom_q), headcount="3"),
                          follow_redirects=True).data.decode("utf-8")
    with app.app_context():
        wk = db.session.get(Task, wk_id)
        check(wk_marked and wk.prev_task == "日報の実績をExcelの集計表に転記する" and wk.next_task == "月報の作成（担当者へ）" and wk.actors == "課長、生産管理課" and wk.domain_id == dom_q and "domain" not in wk.ai_fields_list
              and wk.headcount == 3 and wk.monthly_minutes == wk_minutes * 3,
              "案件詳細で前の業務・後の業務・関わる人・業務領域・担当人数を直せる（氏名は置き換え、直した領域は推定の印を外し、人数は月間時間に掛ける）")
        hours_now = f"月間約 {wk.monthly_minutes / 60:.1f} 時間"
        check(not wk_edited and wk.ai_extra_dict.get("info_edited_at") and wk.ai_source == "rules" and wk.ai_at > wk_at and hours_now in (wk.ai_summary or "") and hours_now not in wk_summary
              and wk.ai_priority == app_module.rule_analyze(wk, app_module._domains(), [])["priority"] and "ルールベースの分析もやり直しました" in html_up and "業務情報を直す前の AI 分析" not in html_up,
              f"R3-10・R3-11: 業務情報を直して保存すると、前回の分析がルールならルールで分析し直し（新しい月間時間の要約・優先度）、改善チームが手を付けた印を残す: {wk.ai_summary}")
    client.post(f"/cases/{wk_id}/update", data=dict(form, domain_id=""))
    code, html = page(f"/cases/{wk_id}")
    check('name="adopt_ai_domain"' in html and 'name="prev_task"' in html and 'name="next_task"' in html and 'name="actors"' in html and 'name="domain_id"' in html and 'name="headcount"' in html,
          "業務領域が空でも「AI推定にする」を選べ、更新フォームに業務情報の欄がある")
    check(all(f'name="{k}"' in html for k in ("purpose", "output", "output_users", "output_users_note", "inputs", "inputs_note", "frequency", "duration"))
          and all(f'id="f-{k}"' in html for k in ("domain_id", "purpose", "output", "output_users", "inputs", "frequency", "duration")),
          "案件詳細の更新フォームに必須項目（目的・アウトプット・利用者と補足・入力元と補足・頻度・1回の時間）の欄がある")
    with app.app_context():
        wk = db.session.get(Task, wk_id)
        marks_before = set(wk.ai_fields_list)
    client.post(f"/cases/{wk_id}/update", data=dict(form, purpose="週の生産状況を佐藤さんと課長会議で報告するため", output="", output_users=["", "上司", "他部署"], output_users_note="課長、生産管理課",
                                                      inputs=["", "他のExcel", "紙の帳票"], inputs_note="日次実績集計表、紙の日報", frequency="毎日", duration="15分未満"))
    with app.app_context():
        wk = db.session.get(Task, wk_id)
        check(wk.purpose == "週の生産状況を担当者と課長会議で報告するため" and wk.output == "" and wk.output_users_list == ["上司", "他部署"] and wk.output_users_note == "課長、生産管理課"
              and wk.inputs_list == ["紙の帳票", "他のExcel"] and wk.inputs_note == "日次実績集計表、紙の日報" and (wk.frequency, wk.duration) == ("毎日", "15分未満") and wk.monthly_minutes == 20 * 10 * 3,
              f"案件詳細で必須項目を直せる（氏名は置き換え、頻度・1回の時間を変えると月間時間を計算し直す）: {wk.monthly_minutes}")
        check({"output_users", "inputs"} <= marks_before and not {"output_users", "inputs", "purpose", "output"} & set(wk.ai_fields_list) and not {"output_users", "inputs"} & set(wk.ai_extra_dict.get("rule_fields") or [])
              and "tools" in wk.ai_fields_list, f"保存した項目だけ推定の印が外れる（直していない項目の印は残る）: {wk.ai_fields_list}")
    code, html = page(f"/cases/{wk_id}")
    check('必須項目（<a href="#f-domain_id" class="alert-link">業務領域</a>・<a href="#f-output" class="alert-link">アウトプット</a>）が未入力です' in html and 'href="#case-info"' in html and 'id="case-info"' in html
          and 'アウトプット（最終的にできるもの） <span class="badge text-bg-warning">未入力</span>' in html,
          "必須項目が足りないときの案内は、案件管理の業務情報の欄（リンク）を指し、空の欄に「未入力」が付く")
    client.post(f"/cases/{wk_id}/update", data=dict(form))
    with app.app_context():
        wk = db.session.get(Task, wk_id)
        check(wk.output_users_list == ["上司", "他部署"] and wk.inputs_list == ["紙の帳票", "他のExcel"] and wk.purpose.startswith("週の生産状況") and wk.frequency == "毎日",
              "業務情報の欄を送らないフォーム（段階・方針だけの更新）では、必須項目を変えない")
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
    code, html = page("/admin/employees")
    check(html.count("デモの社員IDで本物の業務あり") == 1, "デモを入れている間も、デモの社員IDで本物の業務を登録した人には対象者一覧で印が付く")
    r = client.post("/admin/demo/clear", follow_redirects=True)
    html = r.data.decode("utf-8")
    with app.app_context():
        real = Task.query.filter_by(name="本物の業務").first()
        check(real is not None and real.respondent is not None and real.respondent.emp_id == "DEMO-001" and not real.respondent.is_demo and Task.query.count() == 1 and Respondent.query.count() == 1
              and app_module.Employee.query.filter_by(emp_id="DEMO-001").first() is not None and app_module.Employee.query.filter_by(emp_id="DEMO-002").first() is None,
              "デモ削除: 本物の業務がある回答者は残してデモの印を外し、業務を付け替えない。ほかのデモの人は消す")
    check("デモデータを削除しました" in html and "消さずに残しました: DEMO-001 田中 一郎（業務 1 件）" in html and "対象外にしてください" in html,
          "デモ削除で残した人（デモの社員IDで本物の業務がある人）を、社員IDと氏名つきで知らせる")
    code, html = page("/admin/employees")
    row = next((x for x in html.split("<tr") if "DEMO-001" in x), "")
    check(html.count("デモの社員IDで本物の業務あり") == 1 and "デモの社員IDで本物の業務あり" in row and "業務 1 件" in row,
          "対象者一覧: デモ削除で残した人の行に「デモの社員IDで本物の業務あり」の印")

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
    check('id="mq-other-id"' in html and "別の社員IDで回答する" in html and html.index('id="mq-other-id"') < html.index('id="mq-form"') and 'getElementById("mq-other-id").addEventListener' in html and "empInput.focus()" in html,
          "P5: 回答済みの社員IDを入れても「別の社員IDで回答する」で社員IDの欄を空にしてフォームに戻れる")
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
    with app.app_context():
        MonthlyResponse.query.filter_by(period=period, emp_id="DEMO-001").delete()
        r001 = Respondent.query.filter_by(emp_id="DEMO-001").first()
        new_wp = next(w for w in (1, 2, 3) if w != r001.workplace_id)
        new_pat = "交代勤務" if r001.work_pattern != "交代勤務" else "普通勤務"
        db.session.commit()
    r = client.post("/monthly/submit", json=dict(payload, emp_id="DEMO-001", workplace_id=str(new_wp), work_pattern=new_pat))
    with app.app_context():
        r001 = Respondent.query.filter_by(emp_id="DEMO-001").first()
        e001 = Employee.query.filter_by(emp_id="DEMO-001").first()
        check(r.status_code == 200 and (r001.workplace_id, r001.work_pattern) == (new_wp, new_pat) == (e001.workplace_id, e001.work_pattern),
              "P11: 月次で選んだ職場・勤務形態で、「自動化を検討」の回答者と対象者一覧の値も更新する（新しい選択が勝つ）")

    # ---- 設問の編集（要ログイン） ----
    r = client.post("/login", data={"emp_id": "admin", "password": "adminpass"})
    code, html, r = get(f"/monthly/admin?period={period}")
    check(code == 200 and "qe-list" in html and "基本設問に戻す" in html, "設問エディタ")
    r = client.get("/monthly/admin/questions?period=default")
    check(r.get_json()["questions"][0]["key"] == "has", "基本設問の取得")
    base_qs = r.get_json()["questions"]
    far = "2031-05"
    code, html, r = get(f"/monthly/admin?period={far}")
    j_q = client.get(f"/monthly/admin/questions?period={far}").get_json()
    with app.app_context():
        none_yet = MonthlySurvey.query.filter_by(period=far).first() is None
    check(code == 200 and none_yet and j_q["questions"] is None and "まだありません" in html and 'id="qe-active" checked' in html,
          "P14: 設問の編集画面や設問の取得を開いただけでは、その月度の設問セットを作らない")
    r = client.post("/monthly/admin/save", json={"period": far, "questions": base_qs, "is_active": True})
    with app.app_context():
        made = MonthlySurvey.query.filter_by(period=far).first()
        check(r.status_code == 200 and made is not None and len(made.questions) == len(base_qs), "P14: 保存したときに設問セットを作る")
        db.session.delete(made)
        db.session.commit()
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
    with app.app_context():
        x111 = MonthlyResponse.query.filter_by(period=period, emp_id="X111").first().id
        old11 = MonthlyResponse(period=prev, emp_id="X111", name="", workplace_id=2, answers_json=json.dumps({"has": "ある", "wants": "ある", "wants_tasks": "前月の業務", "stance": "改善チームに依頼したい"}, ensure_ascii=False))
        db.session.add(old11)
        db.session.commit()
        old11 = old11.id
    with client.session_transaction() as sess:
        handed = "monthly_last" in sess
    check(f"m={x111}" in j11["register_url"] and f"m={x111}" in r.headers["Location"] and not handed,
          "P4: 月次の送信後のリンクに回答の番号（m）を付けて引き継ぎ、引き継いだら控え（monthly_last）を消す")

    def shown(page, text):
        return text in page or json.dumps(text)[1:-1] in page

    html_old = client.get("/survey/", query_string={"emp": "ｘ１１１", "src": "monthly", "m": old11}).data.decode("utf-8")
    html_bad = client.get("/survey/", query_string={"emp": "DEMO-009", "src": "monthly", "m": old11}).data.decode("utf-8")
    check(f'"period": "{prev}"' in html_old and shown(html_old, "前月の業務") and '"self_improve": false' in html_old and shown(html_old, '"x111"') and "monthly: null" in html_bad,
          "P4: 引き継ぐのは m の回答そのもの（社員IDは全角・小文字でも同じ人。別の人の回答番号なら引き継がない）")
    with client.session_transaction() as sess:
        sess["monthly_last"] = x111
    code, html_d, r_d = get(f"/monthly/done?period={prev}")
    check(code == 200 and "前月の業務" in html_d and "週報の作成" not in html_d and f"m={old11}" in html_d, "P4: 完了ページは URL の月度の回答を出す（直前に送った別の月度の回答ではなく）")
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
        from datetime import datetime as _dt
        from sqlalchemy import text as _sql
        db.session.expunge_all()
        db.session.add(Employee(emp_id="C1", name="追加 一", source="manual", created_at=_dt(2003, 2, 10)))
        db.session.add(Employee(emp_id="C2", name="追加 二", source="answer", created_at=_dt(2003, 3, 20)))
        db.session.add(MonthlyResponse(period="2003-01", emp_id="C2", name="", workplace_id=1, answers_json="{}", created_at=_dt(2003, 1, 25)))
        db.session.commit()
        ps = ["2003-01", "2003-02", "2003-03", "2003-04"]
        pick = lambda: [x for x in app_module.member_rows() if x["emp"].emp_id in ("C1", "C2")]
        before = app_module.member_stats(pick(), ps)
        db.session.add(MonthlyResponse(period="2003-04", emp_id="C1", name="", workplace_id=1, answers_json="{}"))
        db.session.commit()
        after = app_module.member_stats(pick(), ps)
        check([before[p]["members"] for p in ps] == [1, 2, 2, 2] and [after[p]["members"] for p in ps] == [1, 2, 2, 2] and after["2003-04"]["answered"] == 1,
              f"P8: 分母に入るのは「対象者一覧に載った月」と「初めて回答した月」の早いほうから。あとで初めて回答しても過去の月は変わらない: {[before[p]['members'] for p in ps]} → {[after[p]['members'] for p in ps]}")
        db.session.execute(_sql("UPDATE employees SET created_at = NULL WHERE emp_id = 'C2'"))
        db.session.commit()
        app_module._backfill_member_since()
        db.session.commit()
        check(Employee.query.filter_by(emp_id="C2").first().created_at == _dt(2003, 1, 25), "P8: 載った日が空の行は、最初の月次の回答・最初の業務・更新日のうち最も早い日で埋める")
        MonthlyResponse.query.filter(MonthlyResponse.emp_id.in_(["C1", "C2"])).delete(synchronize_session=False)
        Employee.query.filter(Employee.emp_id.in_(["C1", "C2"])).delete(synchronize_session=False)
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
    with app.app_context():
        answered_n = len(app_module.answered_people())
    check(f"回答した人 {answered_n} 人（業務を1件以上登録したか、月次アンケートに回答した人。対象外は除く）" in html and f"回答した人 {answered_n} 人／業務" in html and "一度でも回答した人）" not in html,
          f"P9: 管理のトップの人数は「回答した人」（対象外とまだ回答のない人を除く）: {answered_n} 人")

    # ---- 対象外の人は回答できず、回答率と同じ数え方で回答数を出す ----
    EXCL = "この社員IDは対象外になっています。管理者に連絡してください"
    with app.app_context():
        e011 = Employee.query.filter_by(emp_id="DEMO-011").first().id
    client.post(f"/admin/employees/{e011}/toggle", json={"excluded": True})
    j = client.get(f"/monthly/api/employee/demo-011?period={period}").get_json()
    r_m = client.post("/monthly/submit", json=dict(full, emp_id="DEMO-011", workplace_id="1"))
    r_s = client.post("/survey/api/respondent", json={"emp_id": "DEMO-011", "workplace_id": 1, "work_pattern": "普通勤務"})
    check(j["found"] and j.get("excluded") and j["message"] == EXCL and r_m.status_code == 400 and r_m.get_json()["errors"]["emp_id"] == EXCL and r_s.status_code == 400 and r_s.get_json()["errors"]["emp_id"] == EXCL,
          "P7: 対象外の社員IDは回答できない（照会は見つかった＋対象外、月次の送信・あなたのことの保存は対象外のメッセージ）")
    code, html_f, r = get(f"/monthly/?period={period}&emp=E013")
    js_sv = client.get("/static/survey.js").data.decode("utf-8")
    check("j.excluded" in html_f and (EXCL in html_f or json.dumps(EXCL)[1:-1] in html_f) and "j.excluded" in js_sv and "setIdentity(j.errors.emp_id, true)" in js_sv,
          "P7: 月次アンケートと「自動化を検討」の両方の画面で対象外のメッセージを出す")
    with app.app_context():
        st_now = app_module.member_stats(app_module.member_rows(), [period])[period]
        row_now = next(x for x in app_module.monthly_summary() if x["period"] == period)
        excluded_answers = MonthlyResponse.query.filter_by(period=period, emp_id="DEMO-011").count()
    code, html_r, r = get("/monthly/results")
    code, html_i, r = get("/")
    check(excluded_answers == 1 and row_now["member_answered"] == st_now["answered"] and row_now["n"] == row_now["member_answered"] + 1 and "うち対象者" in html_r and f"（うち対象者 {row_now['member_answered']} 人）" in html_i,
          f"P7: 回答率の隣の回答数は対象者だけで数え、対象外の人の回答があれば「回答 N 人（うち対象者 n 人）」と出す: {row_now['n']} / {row_now['member_answered']}")
    client.post(f"/admin/employees/{e011}/toggle", json={"excluded": False})

    # ---- 職場の削除は使われていないときだけ。名前は変えられる ----
    with app.app_context():
        wp_used = app_module.Workplace.query.filter_by(name="製造1課").first().id
    html_w = client.post(f"/admin/workplace/{wp_used}/delete", follow_redirects=True).data.decode("utf-8")
    with app.app_context():
        kept = db.session.get(app_module.Workplace, wp_used) is not None
    check(kept and "使われているため削除できません" in html_w and "業務 " in html_w and "月次の回答 " in html_w and "対象者 " in html_w, "P13: 業務・月次の回答・対象者・回答者が使っている職場は削除せず、何が使っているかを出す")
    client.post(f"/admin/workplace/{wp_used}/update", data={"name": "製造1課（第一）"})
    html_w = client.post(f"/admin/workplace/{wp_used}/update", data={"name": "製造2課"}, follow_redirects=True).data.decode("utf-8")
    with app.app_context():
        renamed = db.session.get(app_module.Workplace, wp_used).name
    client.post(f"/admin/workplace/{wp_used}/update", data={"name": "製造1課"})
    client.post("/admin/workplace/add", data={"name": "試験の職場"})
    with app.app_context():
        wp_new = app_module.Workplace.query.filter_by(name="試験の職場").first().id
    client.post(f"/admin/workplace/{wp_new}/delete")
    code, html_a, r = get("/admin/")
    with app.app_context():
        check(renamed == "製造1課（第一）" and "同じ名前の職場があります" in html_w and db.session.get(app_module.Workplace, wp_new) is None and f"/admin/workplace/{wp_used}/update" in html_a,
              "P13: 管理で職場名を変えられ（同じ名前は不可）、使われていない職場は削除できる")

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

    # ---- 社員IDの表記ゆれ（全角・小文字）は同じ人。回答のある人は削除せず対象外に ----
    with app.app_context():
        n_before = Employee.query.count()
    r = client.post("/admin/employees/add", data={"emp_id": "ｄｅｍｏ－００９", "name": "小林 直樹", "workplace_id": "1", "work_pattern": ""}, follow_redirects=True)
    csv_body = "社員ID,氏名\n" + "ｄｅｍｏ－０１０ ,加藤 恵\n" + "n001,新規 社員\n"
    r2 = client.post("/admin/employees/import", data={"file": (io.BytesIO(csv_body.encode("utf-8")), "emp.csv")}, content_type="multipart/form-data", follow_redirects=True)
    with app.app_context():
        check(Employee.query.count() == n_before and "社員情報を更新しました" in r.data.decode("utf-8") and "追加 0 人、更新 2 人" in r2.data.decode("utf-8") and Employee.query.filter_by(emp_id="ｄｅｍｏ－００９").first() is None,
              "P6: 手動追加・CSV取込の社員IDは全角を半角にし、大文字・小文字の違いも同じ人として上書きする")
        n001 = Employee.query.filter_by(emp_id="N001").first().id
        db.session.add(Employee(emp_id="R001", name="保存だけ", source="manual"))
        db.session.add(Respondent(emp_id="r001", name="", workplace_id=1, work_pattern="普通勤務"))
        db.session.commit()
        r001 = Employee.query.filter_by(emp_id="R001").first().id
    code, html, r = get("/admin/employees")
    check(f"/admin/employees/{n001}/delete" in html and f"/admin/employees/{e009}/delete" not in html and f"/admin/employees/{r001}/delete" not in html, "P10: 対象者一覧の削除ボタンは、まだ回答も登録もない人だけに出す")
    html_x = client.post(f"/admin/employees/{e009}/delete", follow_redirects=True).data.decode("utf-8")
    html_y = client.post(f"/admin/employees/{r001}/delete", follow_redirects=True).data.decode("utf-8")
    client.post(f"/admin/employees/{n001}/delete")
    with app.app_context():
        check(db.session.get(Employee, e009) is not None and db.session.get(Employee, r001) is not None and "削除できません" in html_x and "「対象外」にチェック" in html_x and "削除できません" in html_y and db.session.get(Employee, n001) is None,
              "P10: 月次の回答・業務・「あなたのこと」の記録がある人は削除せず「対象外」を案内する。まだ何もない人は削除できる")

    _people_checks(check)
    _people_round2_checks(check)
    _migration_checks(check)
    return fails


def _people_checks(check):
    import app as app_module
    from datetime import datetime, timedelta

    tmpdir = tempfile.mkdtemp()
    dbfile = os.path.join(tmpdir, "people.sqlite3").replace("\\", "/")
    app = create_app({"SQLALCHEMY_DATABASE_URI": "sqlite:///" + dbfile, "TESTING": True})
    client = app.test_client()
    client.post("/login", data={"emp_id": "admin", "password": "adminpass"})
    cur = current_period()
    prev = shift_period(cur, -1)
    now = datetime.now()
    old, mid, new = now - timedelta(days=60), now - timedelta(days=30), now - timedelta(days=1)
    with app.app_context():
        w1, w2, w3 = [w.id for w in app_module.Workplace.query.order_by(app_module.Workplace.sort, app_module.Workplace.id).limit(3).all()]
        for e in ("P001", "P002", "P003", "P004", "P005", "P006", "P007", "P008"):
            db.session.add(Employee(emp_id=e, name=f"氏名 {e}", source="answer", active=(e != "P006")))

        def resp(e, wp, skills="", learn=""):
            r = Respondent(emp_id=e, name=f"氏名 {e}", workplace_id=wp, work_pattern="普通勤務", c2_skills=skills, c3_learn=learn)
            db.session.add(r)
            db.session.flush()
            return r

        def task(r, name, at, checked=False, assignment="未定"):
            db.session.add(Task(respondent_id=r.id, workplace_id=r.workplace_id, name=name, frequency="毎週", duration="15分未満", wish="減らしたい",
                                self_improve=app_module.SELF_IMPROVE_YES if checked else app_module.SELF_IMPROVE_NO, assignment=assignment, created_at=at))

        def monthly(e, period, wp, at, **answers):
            row = MonthlyResponse(period=period, emp_id=e, name=f"氏名 {e}", workplace_id=wp, answers_json=json.dumps(answers, ensure_ascii=False), created_at=at)
            db.session.add(row)
            db.session.flush()
            return row.id

        task(resp("P001", w1, skills="特になし", learn="Python|特になし"), "一の業務", old, checked=True)
        monthly("P001", prev, w1, mid, has="ある", wants="ある", wants_tasks="一の業務", stance="改善チームに依頼したい")
        monthly("P002", prev, w1, mid, has="ある", wants="ある", wants_tasks="二の業務", stance="学びながら改善チームと一緒に取り組みたい")
        task(resp("P002", w1), "二の業務", new)
        task(resp("P003", w2), "三の業務", mid)
        want_id = monthly("P004", prev, w2, mid, has="ある", wants="ある", wants_tasks="検査記録の転記\n日報の集計", stance="自分で取り組みたい")
        monthly("P004", cur, w2, new, has="ある", wants="ない")
        resp("P005", w3)
        monthly("P006", cur, w3, new, has="ない")
        task(resp("P007", w2), "七の業務", mid, assignment="自主改善")
        db.session.commit()
    client.post("/logout")
    client.post("/survey/api/respondent", json={"emp_id": "P008", "workplace_id": w3, "work_pattern": "普通勤務"})
    client.post("/login", data={"emp_id": "admin", "password": "adminpass"})
    with app.test_request_context():
        people = {p["emp_id"]: p for p in app_module.dx_people()}
        k = app_module.kpis()
        wps = {w["workplace"].id: w for w in app_module.workplace_stats()}
        quiet = next((s for s in app_module.suggest_actions(limit=50) if s["title"].startswith("まだ回答のない職場")), None)
        all_details = " ".join(s["detail"] for s in app_module.suggest_actions(limit=50))
        have, learn = app_module.skills_summary()
        w3_name = db.session.get(app_module.Workplace, w3).name
        w1_name = db.session.get(app_module.Workplace, w1).name
    check(set(people) == {"P001", "P002", "P003", "P004", "P007"},
          f"P9: 回答した人は業務を1件以上登録したか月次に回答した人だけ（あなたのことの保存だけの人・対象外の人は入らない）: {sorted(people)}")
    check(k["respondent_count"] == 5 and k["self_improver_count"] == 3 and {e for e, p in people.items() if p["positive"]} == {"P002", "P004", "P007"},
          f"P9・P17・P21・R2-2: 自主改善者と分母は同じ「回答した人」から（{k['self_improver_count']}／{k['respondent_count']}）。チェックと月次の進め方は新しいほうで決め、チェックなしは否定に数えない。担当区分「自主改善」で伴走中の人も数える")
    check((wps[w1]["respondents"], wps[w1]["self_improvers"], wps[w2]["respondents"], wps[w2]["self_improvers"], wps[w3]["respondents"]) == (2, 1, 3, 2, 0)
          and all(w["self_improvers"] <= w["respondents"] for w in wps.values()),
          f"P9・P17: 職場別の「回答した人」と自主改善は同じ人の集まりで数える: {[(w['respondents'], w['self_improvers']) for w in wps.values()]}")
    check(quiet is not None and w3_name in quiet["title"] and w1_name not in quiet["title"] and "記入会" not in all_details,
          f"P9: 「まだ回答のない職場」は月次の回答も業務もない職場（対象外の人の回答は数えない）。「記入会」という言葉は使わない: {quiet and quiet['title']}")
    p1, p2, p3, p4, p7 = (people[e] for e in ("P001", "P002", "P003", "P004", "P007"))
    check(p3["wish"] == "チェックなし" and p2["wish"] == "チェックなし" and p1["wish"] == "スキルを身につけながら、一緒にやってみたい" and not p1["positive"] and p2["positive"],
          "P21: 「一緒にやってみたい」にチェックしなかった業務は「チェックなし」（改善チームに任せたいとはみなさない）")
    check(p4["stance_latest"] == "自分で取り組みたい" and p4["stance_label"] == period_label(prev) and p4["self_positive_count"] == 1 and p1["self_positive_count"] == 0,
          "P16: 進め方は空でない直近の回答とその月度で出し、自分で／一緒にの回数も数える")
    check("特になし" not in have and "特になし" not in learn and learn.get("Python") == 1 and p1["skills"] == [] and p1["skills_none"],
          f"P20: 「特になし」は技術の人数・グラフ・学びたい人・教育テーマに数えない: {dict(have)} {dict(learn)}")
    check([x["label"] for x in p7["target_tasks"]] == ["伴走"] and [x["label"] for x in p1["target_tasks"]] == ["希望"], "P23: 担当区分が「自主改善」の案件は業務欄に「伴走」と出す")
    check(p4["monthly_wants"] == ["検査記録の転記", "日報の集計"] and p4["monthly_wants_id"] == want_id, "P24: 業務0件の人は、直近の月次で挙げたなくしたい業務を出す")
    html = client.get("/people/").data.decode("utf-8")
    check("チェックなし" in html and "改善チームに任せたい" not in html and ">伴走<" in html and f"/survey/start?emp=P004&amp;src=monthly&amp;m={want_id}" in html and "検査記録の転記・日報の集計" in html
          and "自分で／一緒に 1回（回答 2 回中）" in html and "自分で／一緒に 0回（回答 1 回中）" in html and period_label(prev) in html and ">特になし<" in html,
          "P16・P20・P21・P23・P24: DX人財情報の一覧（チェックなし・伴走・月次の進め方と回数・特になし・業務0件の人の自動化を検討へのリンク）")
    check(f">{len(people)}<span" in html and "業務を1件以上登録したか、月次アンケートに回答した人（対象外は除く）" in html, "P9: DX人財情報の「回答した人」も同じ決まり")
    rows = list(csv.reader(io.StringIO(client.get("/admin/export/respondents.csv").data.decode("utf-8-sig"))))
    by_id = {x[0]: x for x in rows[1:]}
    check(set(by_id) == set(people) and "P004" in by_id and "自分で取り組みたい" in by_id["P004"] and period_label(prev) in by_id["P004"] and by_id["P003"][5] == "チェックなし",
          f"P19: 回答者CSVは DX人財情報と同じ人・同じ値（月次だけの人も入る）: {sorted(by_id)}")
    html = client.get("/").data.decode("utf-8")
    check("自ら改善に取り組みたい人（回答した人 5 人中）" in html and "<th class=\"text-end\">回答した人</th>" in html and "（うち対象者 1 人）" in html,
          "P9・P17・P7: ダッシュボードの自主改善者の分母・職場別の「回答した人」・月次カードの回答数（うち対象者）")
    html = client.get("/admin/").data.decode("utf-8")
    check("回答した人 5 人（" in html, "P9: 管理のトップも同じ人数")
    rid = None
    with app.app_context():
        rid = Respondent.query.filter_by(emp_id="P008").first()
        check(rid is not None and "P008" not in {p["emp_id"] for p in app_module.dx_people()}, "P9: 「あなたのこと」を保存しただけでは回答した人にならない")


def _people_round2_checks(check):
    import app as app_module
    from datetime import datetime, timedelta

    tmpdir = tempfile.mkdtemp()
    dbfile = os.path.join(tmpdir, "people2.sqlite3").replace("\\", "/")
    app = create_app({"SQLALCHEMY_DATABASE_URI": "sqlite:///" + dbfile, "TESTING": True})
    client = app.test_client()
    client.post("/login", data={"emp_id": "admin", "password": "adminpass"})
    cur = current_period()
    prev = shift_period(cur, -1)
    stale = shift_period(cur, -8)
    now = datetime.now()
    later = "その後の月次で依頼したい"
    with app.app_context():
        w1, w2, w3 = [w.id for w in app_module.Workplace.query.order_by(app_module.Workplace.sort, app_module.Workplace.id).limit(3).all()]
        for e in ("A001", "A002", "A003", "A004", "A005", "A006", "A007", "X001"):
            db.session.add(Employee(emp_id=e, name=f"氏名 {e}", source="answer", created_at=now - timedelta(days=800)))

        def task(e, wp, name, at, checked=False, assignment="未定", **extra):
            r = Respondent.query.filter_by(emp_id=e).first()
            if r is None:
                r = Respondent(emp_id=e, name=f"氏名 {e}", workplace_id=wp, work_pattern="普通勤務")
                db.session.add(r)
                db.session.flush()
            db.session.add(Task(respondent_id=r.id, workplace_id=wp, name=name, frequency="毎週", duration="15分未満", wish="減らしたい",
                                self_improve=app_module.SELF_IMPROVE_YES if checked else "", assignment=assignment, created_at=at, **extra))

        def monthly(e, period, wp, at, **answers):
            db.session.add(MonthlyResponse(period=period, emp_id=e, name=f"氏名 {e}", workplace_id=wp, answers_json=json.dumps(answers, ensure_ascii=False), created_at=at))

        ask = dict(has="ある", wants="ある", wants_tasks="何かの業務", stance="改善チームに依頼したい")
        task("A001", w1, "A1の業務", now - timedelta(days=40), checked=True, assignment="自主改善")
        monthly("A001", prev, w1, now - timedelta(days=20), **ask)
        task("A002", w1, "A2の業務", now - timedelta(days=40), checked=True, assignment="自主改善", stage="request", status="対応中")
        monthly("A002", prev, w1, now - timedelta(days=20), **ask)
        task("A003", w1, "A3の業務", now - timedelta(days=40), checked=True, assignment="自主改善", assignee="改善チーム")
        task("A004", w1, "A4の業務", now - timedelta(days=30))
        monthly("A004", cur, w2, now - timedelta(days=1), has="ない")
        monthly("A005", stale, w3, now - timedelta(days=240), has="ある", wants="ある", wants_tasks="古い業務", stance="自分で取り組みたい")
        monthly("A006", prev, w3, now - timedelta(days=20), has="ある", wants="ある", wants_tasks="新しい業務", stance="自分で取り組みたい")
        task("A007", w3, "A7の業務", now - timedelta(days=730), checked=True)
        monthly("X001", prev, w1, now - timedelta(days=20), has="ない")
        monthly("X001", cur, w1, now - timedelta(days=1), has="ない")
        db.session.commit()
    with app.test_request_context():
        people = {p["emp_id"]: p for p in app_module.dx_people()}
        wps = {w["workplace"].id: w for w in app_module.workplace_stats()}
        in_w1 = {p["emp_id"] for p in app_module.dx_people(workplace_id=w1)}
        in_w2 = {p["emp_id"] for p in app_module.dx_people(workplace_id=w2)}
        quiet = next((s for s in app_module.suggest_actions(limit=50) if s["title"].startswith("まだ回答のない職場")), None)
        names = {w.id: w.name for w in app_module.Workplace.query.all()}
    a1, a2, a3, a4, a5, a6, a7 = (people[e] for e in ("A001", "A002", "A003", "A004", "A005", "A006", "A007"))
    check([x["label"] for x in a1["target_tasks"]] == ["希望"] and [x["label"] for x in a2["target_tasks"]] == ["伴走"] and [x["label"] for x in a3["target_tasks"]] == ["伴走"],
          "R2-1: 登録で担当区分が自主改善になっても開始のままなら「希望」。段階が進んだか改善チームが手を付けた（担当者など）ら「伴走」")
    check(not a1["positive"] and a2["positive"] and a3["positive"], "R2-2: 伴走している「自主改善」の案件がある人は自分で改善したい人に数える（登録で自動で付いた担当区分だけでは数えない）")
    check(a1["wish_note"] == later and not a1["wish_positive"] and a2["wish_note"] == later and a3["wish_note"] == "",
          "R2-5: チェックのあとの月次で依頼したいと答えた人は、意欲のバッジを灰色にして「その後の月次で依頼したい」")
    check(not a5["positive"] and a6["positive"] and a7["positive"] and a5["stance_latest"] == "自分で取り組みたい" and a5["stance_label"] == period_label(stale),
          "R2-6: 月次の進め方で自分で改善したいとみなすのは直近6か月の回答だけ（チェックは古くても数える。表示は空でない直近の答えのまま）")
    check(a4["workplace_id"] == w2 and in_w2 == {"A004"} and in_w1 == {"A001", "A002", "A003", "X001"},
          f"R2-4: 職場はいちばん新しい回答（月次か自動化を検討）で選んだ職場。DX人財情報の職場の絞り込みも同じ: {sorted(in_w1)} {sorted(in_w2)}")
    got = (wps[w1]["respondents"], wps[w1]["self_improvers"], wps[w2]["respondents"], wps[w2]["self_improvers"], wps[w3]["respondents"], wps[w3]["self_improvers"])
    check(got == (4, 2, 1, 0, 3, 2) and wps[w1]["tasks"] == 4 and sum(w["respondents"] for w in wps.values()) == len(people),
          f"R2-4: 職場別の回答した人・自主改善は1人1回、いまの職場で数える（業務は登録した職場）: {got}")
    check(quiet is not None and all(names[w] not in quiet["title"] for w in (w1, w2, w3)), f"R2-4: 「まだ回答のない職場」には前に回答があった職場を出さない: {quiet and quiet['title']}")
    html = client.get("/people/").data.decode("utf-8")
    check(later in html and 'text-bg-secondary">スキルを身につけながら、一緒にやってみたい' in html and "直近6か月" in html, "R2-5・R2-6: DX人財情報に灰色のバッジと「その後の月次で依頼したい」、判断の説明")
    html = client.get("/").data.decode("utf-8")
    check("いまの職場" in html and "1人1回" in html, "R2-4: ダッシュボードの職場別の注記")
    rows = list(csv.reader(io.StringIO(client.get("/admin/export/respondents.csv").data.decode("utf-8-sig"))))
    check("月次の進め方（空でない直近の答え）" in rows[0] and "月次の進め方（直近の回答）" not in rows[0], "R2-7: 回答者CSVの見出しは「月次の進め方（空でない直近の答え）」")
    review = "担当区分の見直しを"
    check(a5["stance_old"] and not a6["stance_old"] and a2["wish_hint"] == review and a2["positive"] and a1["wish_hint"] == "" and a3["wish_hint"] == "",
          "R3-1・R3-2: 6か月より前の進め方には印を付け、伴走している自主改善の案件がある人がその後の月次で依頼したいと答えたら、伴走・自分で改善したいのまま「担当区分の見直しを」")
    html = client.get("/people/").data.decode("utf-8")
    row_of = lambda e: next((x for x in html.split("<tr") if f">{e}<" in x), "")
    check('text-bg-secondary">自分で取り組みたい</span>' in row_of("A005") and f"{period_label(stale)}（6か月より前）" in row_of("A005") and "6か月より前" not in row_of("A006")
          and 'text-bg-success">自分で取り組みたい</span>' in row_of("A006") and f"{later} <span" in row_of("A002") and review in row_of("A002") and review not in row_of("A001") and ">伴走<" in row_of("A002"),
          "R3-1・R3-2: DX人財情報の進め方の列は、6か月より前の答えを灰色にして「6か月より前」。意欲の列は「その後の月次で依頼したい」の横に「担当区分の見直しを」")
    by_id = {x[0]: x for x in rows[1:]}
    check(by_id["A002"][5] == f"スキルを身につけながら、一緒にやってみたい（{later}・{review}）" and by_id["A001"][5] == f"スキルを身につけながら、一緒にやってみたい（{later}）"
          and by_id["A003"][5] == "スキルを身につけながら、一緒にやってみたい" and by_id["A005"][7] == f"{period_label(stale)}（6か月より前）" and by_id["A006"][7] == period_label(prev),
          f"R3-3: 回答者CSVの意欲の列にも同じ注記（その後の月次で依頼したい・担当区分の見直しを）、月度の列に「6か月より前」: {by_id['A002'][5]} / {by_id['A005'][7]}")

    captured = []
    real_ai = app_module.suggest_actions_ai
    app_module.suggest_actions_ai = lambda ctx: captured.append(ctx) or []
    try:
        client.post("/admin/ai", data={"openai_api_url": "https://ai.example.test/api/responses", "openai_api_key": "k", "openai_model": "m"})
        client.post("/ai/suggest")
    finally:
        app_module.suggest_actions_ai = real_ai
        client.post("/admin/ai", data={"openai_api_url": "", "openai_api_key": "", "openai_model": ""})
    ctx = captured[0] if captured else {}
    check(ctx.get("回答した人") == len(people) and "アンケート対象者数（一度でも回答した人）" not in ctx and "回答者数" not in ctx and ctx.get("職場別") and all("回答した人" in w for w in ctx["職場別"]),
          f"R2-11: AI に渡す集計の人数は「回答した人」（{ctx.get('回答した人')} 人）")

    later_at = now + timedelta(minutes=5)
    with app.app_context():
        for e in ("A008", "A009", "A010"):
            db.session.add(Employee(emp_id=e, name=f"氏名 {e}", source="answer", created_at=now - timedelta(days=800)))
        task("A008", w1, "A8の業務", now - timedelta(days=40), assignment="自主改善", assignee="改善チーム")
        monthly("A008", prev, w1, now - timedelta(days=20), **ask)
        monthly("A009", cur, w2, now, has="ない")
        monthly("A009", prev, w3, later_at, has="ない")
        task("A010", w1, "A10の業務", now)
        monthly("A010", prev, w3, later_at, has="ない")
        db.session.commit()
    with app.test_request_context():
        more = {p["emp_id"]: p for p in app_module.dx_people() if p["emp_id"] in ("A008", "A009", "A010")}
    a8, a9, a10 = (more[e] for e in ("A008", "A009", "A010"))
    check(a8["wish"] == "チェックなし" and a8["wish_note"] == later and a8["wish_hint"] == review and a8["positive"] and [x["label"] for x in a8["target_tasks"]] == ["伴走"],
          "R3-2: 改善チームが自主改善にして伴走している案件（チェックなし）も、その後の月次で依頼したいと答えたら「担当区分の見直しを」（伴走・自分で改善したいのまま）")
    check(a9["workplace_id"] == w2 and a10["workplace_id"] == w1,
          f"R3-7: いまの職場は新しい月度（月次の月度・業務を登録した月）から決め、同じ月度なら回答日時で決める（前の月度にあとから答えても、新しい月度の職場は変わらない）: {a9['workplace_id']} {a10['workplace_id']}")

    with app.app_context():
        x1 = Employee.query.filter_by(emp_id="X001").first().id
    j = client.post(f"/admin/employees/{x1}/toggle", json={"excluded": True}).get_json()
    via_route = {s["period"]: s for s in j["stats"]}
    with app.app_context():
        at = db.session.get(Employee, x1).excluded_at
        st = app_module.member_stats([r for r in app_module.member_rows() if r["emp"].emp_id == "X001"], [prev, cur])
    check(at is not None and (st[prev]["members"], st[prev]["answered"], st[cur]["members"]) == (1, 1, 0) and via_route[prev]["answered"] >= 1,
          f"R2-10: 対象外にした日を残し、それより前の月は回答率の分子・分母にそのまま入れる（その月から外す）: {st}")
    client.post(f"/admin/employees/{x1}/toggle", json={"excluded": False})
    month_start = datetime.strptime(cur + "-01", "%Y-%m-%d")
    with app.app_context():
        e = db.session.get(Employee, x1)
        cleared = e.excluded_at is None and e.active
        e.active, e.updated_at = False, month_start
        db.session.commit()
    restarted = create_app({"SQLALCHEMY_DATABASE_URI": "sqlite:///" + dbfile, "TESTING": True})
    with restarted.app_context():
        filled = db.session.get(Employee, x1).excluded_at
        st = app_module.member_stats([r for r in app_module.member_rows() if r["emp"].emp_id == "X001"], [prev, cur])
    check(cleared and filled == month_start and (st[prev]["members"], st[prev]["answered"], st[cur]["members"]) == (1, 1, 0),
          f"R2-10・R3-6: 対象に戻すと日付は消える。日付のない対象外（以前のデータ）は起動時に更新日で埋め、それより前の月は分子・分母に入れたまま: {filled} {st}")

    with app.test_request_context():
        pre_old, pre_new = app_module._stance_json("A005"), app_module._stance_json("A006")
        old_m = MonthlyResponse.query.filter_by(emp_id="A005").first().id
        new_m = MonthlyResponse.query.filter_by(emp_id="A006").first().id
    j5 = client.post("/survey/api/respondent", json={"emp_id": "A005", "workplace_id": w3, "work_pattern": "普通勤務"}).get_json()
    j6 = client.post("/survey/api/respondent", json={"emp_id": "A006", "workplace_id": w3, "work_pattern": "普通勤務"}).get_json()
    h5 = client.get("/survey/", query_string={"emp": "A005", "src": "monthly", "m": old_m}).data.decode("utf-8")
    h6 = client.get("/survey/", query_string={"emp": "A006", "src": "monthly", "m": new_m}).data.decode("utf-8")
    check(pre_old is None and j5["stance"] is None and pre_new["self_improve"] is True and j6["stance"]["self_improve"] is True and '"self_improve": false' in h5 and '"self_improve": true' in h6,
          f"R3-1: 「一緒にやってみたい」の事前チェックの元（月次の進め方）も直近6か月の回答だけ。6か月より前の「自分で取り組みたい」ではチェックを入れない: {j5.get('stance')} {j6.get('stance')}")


def _migration_checks(check):
    import sqlite3
    from datetime import datetime

    tmpdir = tempfile.mkdtemp()
    path = os.path.join(tmpdir, "old.sqlite3")
    uri = "sqlite:///" + path.replace("\\", "/")
    app1 = create_app({"SQLALCHEMY_DATABASE_URI": uri, "TESTING": True})
    with app1.app_context():
        db.session.add(MonthlyResponse(period="2004-03", emp_id="Q001", name="", workplace_id=1, answers_json="{}", created_at=datetime(2004, 3, 9)))
        r_low = Respondent(emp_id="q001", name="", workplace_id=1, work_pattern="普通勤務")
        r_up = Respondent(emp_id="Q001", name="", workplace_id=2, work_pattern="", c2_skills="Excel関数")
        r_q3 = Respondent(emp_id="q003", name="", workplace_id=1, work_pattern="普通勤務")
        db.session.add_all([r_low, r_up, r_q3])
        db.session.flush()
        db.session.add(MonthlyResponse(period="2004-04", emp_id="q001", name="", workplace_id=1, answers_json="{}", created_at=datetime(2004, 4, 2), respondent_id=r_low.id))
        db.session.add(MonthlyResponse(period="2004-04", emp_id="Ｑ００１", name="重複 太郎", workplace_id=1, answers_json="{}", created_at=datetime(2004, 4, 20)))
        db.session.add(MonthlyResponse(period="2004-05", emp_id="q003", name="", workplace_id=1, answers_json="{}", created_at=datetime(2004, 5, 2)))
        steps = json.dumps([{"no": 1, "action": "日報を転記する"}], ensure_ascii=False)
        db.session.add(Task(respondent_id=r_low.id, workplace_id=1, name="Q1の業務", frequency="毎週", duration="15分未満", steps_json=steps, self_improve="改善チームに任せたい"))
        db.session.add(Task(respondent_id=r_up.id, workplace_id=2, name="Q1の別業務", frequency="毎週", duration="15分未満", steps_json="[]", self_improve="改善チームに任せたい"))
        db.session.add(Task(respondent_id=r_q3.id, workplace_id=1, name="Q3の業務", frequency="毎週", duration="15分未満", steps_json=steps, self_improve="スキルを身につけながら、一緒にやってみたい"))
        db.session.commit()
        db.engine.dispose()
    con = sqlite3.connect(path)
    con.executescript(
        "CREATE TABLE employees_old (id INTEGER PRIMARY KEY, emp_id VARCHAR(50) NOT NULL UNIQUE, name VARCHAR(100) NOT NULL, department VARCHAR(100), work_pattern VARCHAR(20) DEFAULT '', "
        "workplace_id INTEGER, email VARCHAR(200), active BOOLEAN, source VARCHAR(20), is_demo BOOLEAN, updated_at DATETIME);"
        "DROP TABLE employees; ALTER TABLE employees_old RENAME TO employees;"
    )
    task_cols = [row[1] for row in con.execute("PRAGMA table_info(tasks)").fetchall() if row[1] != "monthly_name"]
    con.executescript(f"CREATE TABLE tasks_old AS SELECT {', '.join(task_cols)} FROM tasks; DROP TABLE tasks; ALTER TABLE tasks_old RENAME TO tasks;")
    con.executemany("INSERT INTO employees (emp_id, name, department, work_pattern, email, active, source, is_demo, updated_at) VALUES (?, ?, '', '', '', ?, ?, 0, ?)", [
        ("Q001", "", 1, "answer", "2004-05-01 00:00:00.000000"),
        ("ｑ００１", "正しい 名前", 0, "manual", "2004-04-01 00:00:00.000000"),
        ("q001", "", 1, "answer", "2004-06-01 00:00:00.000000"),
        ("Ｑ００２", "全角 だけ", 1, "csv", "2004-07-01 00:00:00.000000"),
        ("Q003", "三番 目", 1, "manual", "2004-08-01 00:00:00.000000"),
    ])
    con.commit()
    con.close()
    import logging

    import app as app_module

    logged = []
    catcher = logging.Handler(logging.WARNING)
    catcher.emit = logged.append
    app_module.log.addHandler(catcher)
    try:
        app2 = create_app({"SQLALCHEMY_DATABASE_URI": uri, "TESTING": True})
    finally:
        app_module.log.removeHandler(catcher)
    merged_logs = [x.getMessage() for x in logged if "同じ月度の月次の回答" in x.getMessage()]
    check(len(merged_logs) == 1 and "月度 2004-04" in merged_logs[0] and "社員ID Ｑ００１" in merged_logs[0] and "2004-04-20" in merged_logs[0] and "重複" not in merged_logs[0],
          f"R3-8: 起動時のまとめで同じ月度の月次の回答を消したら、月度と社員ID（氏名は書かない）を警告としてログに残す: {merged_logs}")
    with app2.app_context():
        from sqlalchemy import inspect
        cols = {c["name"] for c in inspect(db.engine).get_columns("employees")}
        task_cols = {c["name"] for c in inspect(db.engine).get_columns("tasks")}
        rows = Employee.query.order_by(Employee.emp_id).all()
        check("created_at" in cols and [e.emp_id for e in rows] == ["Q001", "Q002", "q003"],
              f"P6・R2-8: 起動時に、全角・大文字小文字だけが違う社員IDの行を1行にまとめ、半角にそろえる（回答で使われている書き方を残す）: {[e.emp_id for e in rows]}")
        check(rows[0].active is False and rows[0].name == "正しい 名前", "P6: まとめるとき、どれかが対象外なら対象外のまま、氏名はいちばん確かなものを残す")
        st_q = app_module.member_stats([r for r in app_module.member_rows() if r["emp"].emp_id == "Q001"], ["2004-03", "2004-04"])
        check(rows[0].excluded_at == datetime(2004, 4, 1) and (st_q["2004-03"]["members"], st_q["2004-03"]["answered"], st_q["2004-04"]["members"]) == (1, 1, 0),
              f"R3-6: 日付のない対象外は起動時に更新日（なければ起動した日）で対象外にした日を埋め、それより前の月は回答率に入れたまま: {rows[0].excluded_at} {st_q}")
        check(rows[0].created_at == datetime(2004, 3, 9) and rows[1].created_at == datetime(2004, 7, 1),
              f"P8: 起動時に employees へ載った日（created_at）を足し、最初の月次の回答・最初の業務・更新日の早いほうで埋める: {[e.created_at for e in rows]}")
        resp = Respondent.query.filter(db.func.lower(Respondent.emp_id) == "q001").all()
        mrs = MonthlyResponse.query.filter(db.func.lower(MonthlyResponse.emp_id) == "q001").order_by(MonthlyResponse.period).all()
        kept = resp[0] if len(resp) == 1 else None
        check(kept is not None and kept.emp_id == "Q001" and sorted(t.name for t in kept.tasks) == ["Q1の別業務", "Q1の業務"] and kept.c2_skills == "Excel関数" and kept.work_pattern == "普通勤務"
              and [(m.period, m.emp_id) for m in mrs] == [("2004-03", "Q001"), ("2004-04", "Q001")] and mrs[1].created_at == datetime(2004, 4, 2) and mrs[1].respondent_id == kept.id,
              f"R2-8: 起動時に月次の回答と「自動化を検討」の回答者の社員IDもそろえ、同じ人の回答者は1行にまとめて業務を移す（同じ月度の重複した回答は先のものを残す）: {[(m.period, m.emp_id) for m in mrs]}")
        q3 = Respondent.query.filter(db.func.lower(Respondent.emp_id) == "q003").first()
        check(q3.emp_id == "q003" and MonthlyResponse.query.filter_by(period="2004-05").first().emp_id == "q003", "R2-8: 社員IDの書き方は、すでに回答で使われているものを残す")
        t1, t2, t3 = (Task.query.filter_by(name=n).first() for n in ("Q1の業務", "Q1の別業務", "Q3の業務"))
        check("excluded_at" in cols and "monthly_name" in task_cols and t1.self_improve == "" and t2.self_improve == "改善チームに任せたい" and t3.self_improve == "スキルを身につけながら、一緒にやってみたい" and (t1.monthly_name or "") == "",
              "R2-3・R2-10・R2-15: 起動時に対象外にした日・始めた月次の業務名の列を足し、「自動化を検討」で登録した業務（手順あり）の「改善チームに任せたい」は空（チェックなし）にする（手順のない古い業務はそのまま）")
        db.engine.dispose()

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
    check(j["stance"] and j["stance"]["value"] == "改善チームに依頼したい" and j["stance"]["self_improve"] is False and j["stance"]["period"] == current_period(),
          f"P22: あなたのことを保存すると、月次の直近の進め方（空でないもの）と月度を返す（月次から来たときに限らない）: {j.get('stance')}")
    code, html_m, r = get("/survey/?emp=T001&src=monthly")
    check(code == 200 and '"self_improve": false' in html_m and ('"wants_tasks": ["\\u65e5' in html_m or '"wants_tasks": ["日報の集計"]' in html_m), "アンケートから来ると、なくしたい業務と進め方（依頼したい＝チェックなし）を渡す")
    check('"respondent": null' in html_m, "開いたときは前の回答者を戻さず、空で始める")
    check('"workplace_id": 1' in html_m, "月次アンケートから来たときは、その回答で選んだ職場を引き継ぐ")
    client.post("/survey/api/respondent", json={"emp_id": "T001", "workplace_id": 1, "work_pattern": "交代勤務"})
    r1 = client.post("/survey/api/tasks", json={"emp_id": "T002", "name": "別人の業務", "state": {}, "messages": []})
    r2 = client.post("/survey/api/intake", json={"emp_id": "T002", "title": "x", "messages": [], "state": {}})
    r3 = client.delete("/survey/api/tasks/999999", json={"emp_id": "T002"})
    r4 = client.post("/survey/api/tasks", json={"emp_id": "", "name": "x", "state": {}, "messages": []})
    check(all(x.status_code == 409 and "あなたのこと" in x.get_json()["message"] and x.get_json()["errors"]["_"] == "respondent" for x in (r1, r2, r3, r4)),
          "P1: 画面の社員ID（emp_id）が保存した回答者と違えば、聞き取り・登録・取り消しは 409 で「あなたのこと」を入れ直すよう返す")
    r = client.post("/survey/api/tasks", json={"emp_id": "ｔ００１", "name": "", "state": {}, "messages": []})
    check(r.status_code == 400 and "name" in r.get_json()["errors"], "P1: 全角・小文字でも同じ社員IDなら同じ人として通す（入力チェックへ進む）")
    js_rq = client.get("/static/survey.js").data.decode("utf-8")
    check("function forgetPerson()" in js_rq and "clearSelects();" in js_rq and "api(\"POST\", CFG.api.reset)" in js_rq and js_rq.count("emp_id: empNow()") >= 4 and "respondentLost(j)" in js_rq,
          "P1: 社員IDを打ち替えた・見つからないときは回答者・登録一覧・候補・職場・勤務形態を消してリセットし、聞き取り・登録・取り消しに社員IDを付けて送る")
    check("confirmedId && confirmedId !== key" in js_rq and js_rq.index("forgetPerson();\n      confirmedId = key;") < js_rq.index("if (v.workplace_id && v.work_pattern) debounce"),
          "P2: 別の人の社員IDが見つかったときは、前の人の職場・勤務形態を消してから（自動保存はしない）")
    check("resetPage();" in js_rq and "function resetPage()" in js_rq and "history.replaceState" in js_rq and js_rq.index("await api(\"POST\", CFG.api.reset);\n    resetPage();") > 0,
          "P3: 送信して完了で回答者をリセットし、お礼のあと画面を空に戻す（URL の社員IDも消す）")
    check("if (j.respondent) S.respondent = j.respondent;" in js_rq, "P18: 登録のあと S.respondent を新しい値にして、技術の欄の事前選択に使う")
    check("O.skill_none" in js_rq and "S.stance" in js_rq and "conv.origin === \"monthly\" && CFG.monthly && CFG.monthly.self_improve" not in js_rq,
          "P20・P22: 技術の「特になし」は他と同時に選べない。「一緒にやってみたい」の事前チェックは月次からに限らず直近の進め方から")
    check("t.can_delete" in js_rq and "t.stage" in js_rq and "j.duplicate.can_delete" in js_rq, "P15: 登録一覧に段階を出し、取り消せないときは「取り消す」を出さない")

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
        with app.app_context():
            e_t1 = Employee.query.filter_by(emp_id="T001").first()
            e_t1.workplace_id, e_t1.work_pattern = 2, "普通勤務"
            db.session.commit()
        r = client.post("/survey/api/tasks", json={"emp_id": "T001", "state": final_state, "messages": msgs, "frequency": "毎日", "duration": "15〜30分", "wish": "減らしたい", "self_improve": True, "skills": ["Excel関数", "特になし"], "learn": ["Python"], "shift": "夜勤"})
        j = r.get_json()
        b = calls[-1]["body"]
        check(r.status_code == 200 and j["success"] and j["first_impression"] == finalize["first_impression"] and len(j["tasks"]) == 1 and j["suggested"] == ["出勤簿の転記"], "登録: AI の一言が返り、候補から登録した業務が消える")
        check(j["respondent"]["c2_skills"] == ["Excel関数"] and j["respondent"]["c3_learn"] == ["Python"], f"P18・P20: 登録の応答に新しい技術が入り、「特になし」は他の技術と一緒には残らない: {j['respondent']['c2_skills']}")
        check(j["tasks"][0]["stage"] == "開始" and j["tasks"][0]["can_delete"] is True, f"P15: 登録した業務に段階と取り消せるかが付く: {j['tasks'][0].get('stage')}")
        with app.app_context():
            e_t1 = Employee.query.filter_by(emp_id="T001").first()
            t_new = Task.query.filter_by(name="日報の集計").first()
            check((e_t1.workplace_id, e_t1.work_pattern) == (2, "普通勤務") and t_new.workplace_id == 1,
                  "P11: 業務の登録で、対象者一覧の職場・勤務形態を回答者の古い値で上書きしない（業務の職場は「あなたのこと」の選択）")
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
            t2_id = t2.id
        refused = []
        for change in ({"decision": "廃止"}, {"status": "保留"}, {"stage": "request", "status": "対応中"}):
            with app.app_context():
                t2 = db.session.get(Task, t2_id)
                t2.decision, t2.status, t2.stage = "未定", "受付", "start"
                for k_, v_ in change.items():
                    setattr(t2, k_, v_)
                db.session.commit()
            r = client.delete(f"/survey/api/tasks/{t2_id}", json={"emp_id": "T001"})
            refused.append(r.status_code == 409 and r.get_json()["message"] == "改善チームが対応を始めているため取り消せません")
        r = client.post("/survey/api/tasks", json={"emp_id": "T001", "name": "相談したい業務", "state": state(title="相談したい業務", steps=[step1]), "messages": [], "frequency": "毎週", "duration": "15分未満", "wish": "まずは相談したい", "shift": "昼勤"})
        dup = r.get_json()
        with app.app_context():
            st2 = next(x for x in app_module._state(db.session.get(Task, t2_id).respondent)["tasks"] if x["id"] == t2_id)
            check(all(refused) and db.session.get(Task, t2_id) is not None and r.status_code == 409 and dup["duplicate"]["can_delete"] is False and st2["can_delete"] is False and st2["stage"] == "要求定義",
                  f"P15: 方針を決めた・保留・段階が進んだ案件は、本人は取り消せない（409）。二重登録の警告でも「前の登録を取り消す」は出さない: {refused}")
            t2 = db.session.get(Task, t2_id)
            t2.stage, t2.status = "start", "受付"
            db.session.commit()

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
            kinds_a, kinds_b = app_module._word_kinds("設備の稼働状況を課長に報告する"), app_module._word_kinds("顧客クレームの対応状況を共有する")
            check(kinds_a.get("状況") == "part" and app_module._pair_weight("状況", kinds_a, kinds_b) == 1 and app_module._pair_weight("excel", {"excel": "whole"}, {"excel": "whole"}) == 1
                  and app_module._pair_weight("日次実績集計表", {"日次実績集計表": "whole"}, {"日次実績集計表": "whole"}) == 3,
                  "関連業務の語の重み: 長い語から切り出した2文字どうし（状況）やツール・媒体の語（Excel）は1、3文字以上の語は3")
            demo7 = Respondent.query.filter_by(emp_id="DEMO-007").first()
            t_eq = Task(respondent=demo7, workplace_id=demo7.workplace_id, name="設備の稼働日報を作って課長に送る", purpose="設備の稼働状況を課長に報告するため", frequency="毎日", duration="15〜30分", wish="減らしたい",
                        steps_json=json.dumps([{"no": 1, "action": "設備モニタで稼働状況を見る", "system": "設備モニタ"}, {"no": 2, "action": "稼働日報をExcelで作る", "system": "Excel（稼働日報）", "output": "稼働日報"},
                                               {"no": 3, "action": "課長にメールで送る", "system": "メール", "to": "課長"}], ensure_ascii=False))
            db.session.add(t_eq)
            db.session.flush()
            app_module._finalize_task(t_eq, app_module.task_state(t_eq), [], use_ai=False)
            rel_eq = [o.name for o in app_module.related_tasks_for(t_eq)]
            check("クレーム対応の進捗報告" not in rel_eq and "週報の作成" not in rel_eq and "月次保全報告書の作成" in rel_eq and t_eq.ai_related_list == rel_eq[:5],
                  f"関連業務（ルール）: 2文字の切れ端（状況・報告）だけの重なりではつながず、設備・稼働の重なる業務だけ残す: {rel_eq}")
            check(t_eq.inputs_list == ["設備データ"], f"ルールの分類: システムの語がなくても、設備モニタは情報の取り出し元（設備データ）にする: {t_eq.inputs_list}")
            db.session.rollback()
            by_name = {t.name: t for t in Task.query.all()}
            check("生産実績の月報作成" in [o.name for o in app_module.related_tasks_for(by_name["生産性指標の報告（部へ）"])]
                  and "保全部品の発注依頼書作成" not in [o.name for o in app_module.related_tasks_for(by_name["外注先への納期確認メール"])]
                  and "週次生産計画の作成と配布" not in [o.name for o in app_module.related_tasks_for(t_week)],
                  "関連業務: 前後の業務の名前（「の」の有無は問わない）でつながる業務は残し、短い語（発注）や頻度の語（週次）だけではつながない")
            src = {n: app_module._source_category(n) for n in ("設備モニタ", "画面", "端末", "モニター", "品質の画面", "ファイルサーバ")}
            check(src == {"設備モニタ": "設備データ", "画面": "その他", "端末": "その他", "モニター": "設備データ", "品質の画面": "品質システム", "ファイルサーバ": "他のExcel"},
                  f"ルールの分類: モニタ・画面・端末はシステムの取り出し元として分類し、共有フォルダ・ファイルサーバはファイル（他のExcel）にする: {src}")
            nat = {a: app_module._step_natures({"action": a}) for a in ("共有フォルダに保存する", "ファイルサーバに格納する", "共有フォルダに保存してメールで知らせる", "メールで共有する", "日報を課長と共有する")}
            check(nat["共有フォルダに保存する"] == ["台帳更新"] and nat["ファイルサーバに格納する"] == ["台帳更新"] and "メール・配布" in nat["共有フォルダに保存してメールで知らせる"]
                  and nat["メールで共有する"] == ["メール・配布"] and "メール・配布" not in nat["日報を課長と共有する"],
                  f"ルールの分類: 共有フォルダへの保存はメール・配布にせず台帳更新（ファイリング）にし、メール・配布は「共有」とメール・送る・配るがそろったときだけ: {nat}")
            cats = app_module.rule_categories(Task(name="x"), app_module._clean_state({"steps": [
                {"no": 1, "action": "端末の画面で前日の不良数を見る", "system": "端末"}, {"no": 2, "action": "不良数を集計する", "system": "Excel（不良集計表）", "output": "不良集計表"},
                {"no": 3, "action": "不良集計表を共有フォルダに保存する", "system": "共有フォルダ"}]}))
            check(cats["inputs"] == ["その他"] and "メール・配布" not in cats["nature"] and "台帳更新" in cats["nature"],
                  f"ルールの分類: 端末も取り出し元に数え、保存先の共有フォルダは取り出し元・メール・配布にしない: {cats}")

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
        with client.session_transaction() as sess:
            left = dict(sess)
        r = client.delete(f"/survey/api/tasks/{tid}", json={"emp_id": "T001"})
        check(r.status_code == 400 and "user" not in left and "respondent_id" not in left, f"P12: ログアウトでセッションを丸ごと消す（回答者も残らない）: {sorted(left)}")
        client.post("/survey/api/respondent", json={"emp_id": "T001", "workplace_id": 1, "work_pattern": "交代勤務"})
        r = client.delete(f"/survey/api/tasks/{tid}", json={"emp_id": "T001"})
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
    client.post("/login", data={"emp_id": "admin", "password": "adminpass"})
    client.post("/admin/ai", data={"openai_api_url": "", "openai_api_key": "", "openai_model": ""})
    client.post("/logout")
    client.post("/survey/api/respondent", json={"emp_id": "T002", "name": "試験 次郎", "workplace_id": 1, "work_pattern": "普通勤務"})
    r = client.post("/survey/api/intake", json={"emp_id": "T002", "title": "x", "messages": [], "state": {}})
    check(r.status_code == 400 and "AI連携が設定されていません" in r.get_json()["message"], "AI 未設定なら聞き取りは始められず、設定を案内する")

    _survey_round2_checks(check)
    return fails


def _survey_round2_checks(check):
    import app as app_module
    from flask import session as flask_session

    tmpdir = tempfile.mkdtemp()
    dbfile = os.path.join(tmpdir, "survey2.sqlite3").replace("\\", "/")
    app = create_app({"SQLALCHEMY_DATABASE_URI": "sqlite:///" + dbfile, "TESTING": True})
    client = app.test_client()
    cur = current_period()
    prev = shift_period(cur, -1)
    excl = "この社員IDは対象外になっています。管理者に連絡してください"
    step = {"no": 1, "action": "日報の数字をExcelに転記する", "system": "Excel", "input": "日報", "output": "集計表", "to": ""}

    def shown(page, text):
        return text in page or json.dumps(text)[1:-1] in page

    def answers(**kw):
        base = {"has": "ある", "hours": "15分未満", "task_names": ["日報の集計"], "trend": "変わらない", "wants": "ある", "wants_tasks": ["日報の集計"], "stance": "改善チームに依頼したい"}
        base.update(kw)
        return base

    def register(emp, name, **kw):
        body = {"emp_id": emp, "name": name, "state": {"title": name, "steps": [step]}, "messages": [], "frequency": "毎週", "duration": "15分未満", "wish": "減らしたい"}
        body.update(kw)
        return client.post("/survey/api/tasks", json=body)

    with app.app_context():
        for e in ("S001", "S002", "S003"):
            db.session.add(Employee(emp_id=e, name=f"試験 {e}", source="manual"))
        survey = app_module.get_or_create_survey(cur)
        db.session.add(MonthlySurvey(period=prev, is_active=True, questions_json=survey.questions_json))
        db.session.add(Respondent(emp_id="s001", name="", workplace_id=1, work_pattern="普通勤務"))
        old = MonthlyResponse(period=prev, emp_id="s001", name="", workplace_id=1, answers_json=json.dumps(
            {"has": "ある", "task_names": "前月の業務", "wants": "ある", "wants_tasks": "前月の業務\n日報・週報の集計\n佐藤さんへの報告", "stance": "改善チームに依頼したい"}, ensure_ascii=False))
        new = MonthlyResponse(period=cur, emp_id="S001", name="", workplace_id=1, answers_json=json.dumps(
            {"has": "ある", "task_names": "今月の業務", "wants": "ある", "wants_tasks": "今月の業務", "stance": "自分で取り組みたい"}, ensure_ascii=False))
        db.session.add_all([old, new])
        db.session.commit()
        old_id, new_id = old.id, new.id

    j = client.get(f"/monthly/api/employee/S001?period={prev}").get_json()
    html = client.get(f"/monthly/?period={prev}&emp=S001").data.decode("utf-8")
    r = client.post("/monthly/submit", json={"period": prev, "emp_id": "S001", "workplace_id": "1", "work_pattern": "普通勤務", "answers": answers()})
    check(j["answered"] is True and 'id="mq-form" class="card mq-card d-none"' in html and r.status_code == 400 and "回答済み" in r.get_json()["message"],
          "R2-8: 月次の回答済みの判定（照会・フォーム・送信の二重チェック）は社員IDの大文字・小文字の違いを同じ人とみなす")
    r = client.post("/survey/api/respondent", json={"emp_id": "S001", "workplace_id": 1, "work_pattern": "普通勤務"})
    with app.app_context():
        n_resp = Respondent.query.filter(db.func.lower(Respondent.emp_id) == "s001").count()
    check(r.status_code == 200 and r.get_json()["resumed"] is True and n_resp == 1, "R2-8: 「あなたのこと」の回答者も大文字・小文字の違いで別の人にしない")
    check(r.get_json()["suggested"] == ["今月の業務"], f"R2-12: 引き継ぎがなければ候補は直近の月次の回答から: {r.get_json()['suggested']}")

    html = client.get("/survey/", query_string={"emp": "S001", "src": "monthly", "m": old_id}).data.decode("utf-8")
    with client.session_transaction() as sess:
        handed = sess.get("monthly_handoff")
    r = client.post("/survey/api/respondent", json={"emp_id": "S001", "workplace_id": 1, "work_pattern": "普通勤務"})
    sug = r.get_json()["suggested"]
    check(handed == old_id and sug == ["前月の業務", "日報・週報の集計", "担当者への報告"] and r.get_json()["suggested_wants"] == sug,
          f"R2-12・R2-14: 月次から引き継いだ回答（m）を候補に使い、業務名は「・」で分けず、氏名は担当者に置き換える: {sug}")
    check(shown(html, "日報・週報の集計") and shown(html, "担当者への報告") and not shown(html, "佐藤さん"), "R2-14: 月次から来たときの業務名（引き継ぎの帯）も「・」で分けず、氏名を置き換える")
    with app.test_request_context():
        resp = Respondent.query.filter(db.func.lower(Respondent.emp_id) == "s001").first()
        flask_session["monthly_handoff"] = old_id
        hist_m = app_module.intake_handlers(resp)["person_history"]({})["月次アンケート"]
        flask_session.pop("monthly_handoff")
        hist_latest = app_module.intake_handlers(resp)["person_history"]({})["月次アンケート"]
    check(hist_m["月度"] == prev and hist_latest["月度"] == cur, "R2-12: AI の道具 person_history も引き継いだ月次の回答を優先し、なければ直近の回答（大文字・小文字の違いも同じ人）")

    r = register("S001", "報告書をまとめる", monthly_name="佐藤さんへの報告")
    j = r.get_json()
    with app.app_context():
        t = Task.query.filter_by(name="報告書をまとめる").first()
        mn = t.monthly_name if t else None
    check(r.status_code == 200 and mn == "担当者への報告" and j["task"]["monthly_name"] == "担当者への報告" and "担当者への報告" not in j["suggested"] and "日報・週報の集計" in j["suggested"],
          f"R2-15: 月次の業務名から始めた登録はその名前を残し（氏名は置き換え）、業務名が変わっても候補から外す: {j.get('suggested')}")
    js = client.get("/static/survey.js").data.decode("utf-8")
    check("monthly_name: conv.monthlyName" in js and "function doneNames()" in js and "startIntake(false, \"monthly\", next)" in js and "monthlySource(" in js,
          "R2-15: 画面は月次の候補・引き継ぎから始めたときに業務名を送り、登録済みの判定にも使う")

    r = client.post("/monthly/submit", json={"period": cur, "emp_id": "S002", "workplace_id": "9999", "work_pattern": "普通勤務", "answers": answers()})
    check(r.status_code == 400 and r.get_json()["errors"].get("workplace_id") == "職場を選んでください", "R2-13: 月次の送信は職場マスタにない職場を受け付けない")
    r = client.post("/survey/api/respondent", json={"emp_id": "S002", "workplace_id": 9999, "work_pattern": "普通勤務"})
    check(r.status_code == 400 and r.get_json()["errors"].get("workplace_id") == "職場を選んでください", "R2-13: 「あなたのこと」も職場マスタにない職場を受け付けない")
    with client.session_transaction() as sess:
        sess["monthly_last"] = old_id
    r = client.post("/monthly/submit", json={"period": cur, "emp_id": "S002", "workplace_id": "1", "work_pattern": "普通勤務",
                                             "answers": answers(wants_tasks=["日報・週報の集計", "佐藤さんへの報告"])})
    check(r.status_code == 200 and r.get_json()["wants_tasks"] == ["日報・週報の集計", "担当者への報告"], f"R2-14: 送信後に返す業務名も「・」で分けず、氏名を置き換える: {r.get_json().get('wants_tasks')}")
    with client.session_transaction() as sess:
        kept_last = "monthly_last" in sess
    html_after_js = client.get(f"/monthly/done?period={cur}").data.decode("utf-8")
    check(not kept_last and "自動化を検討する" not in html_after_js and "今月の業務" not in html_after_js,
          "R3-9: 画面（JavaScript）からの月次の送信では完了ページの控え（monthly_last）を残さず、前の人の控えも消す（完了ページに前の人のリンクが出ない）")
    with app.app_context():
        pending = {p["name"] for p in app_module.pending_task_names(1)}
    check("日報・週報の集計" in pending and "担当者への報告" not in pending, f"R2-15: 業務マップの未登録の業務も、始めた月次の業務名を登録済みとみなす: {sorted(pending)}")

    form = {"period": cur, "emp_id": "S003", "workplace_id": "1", "work_pattern": "普通勤務", "comment": ""}
    for k, v in answers(task_names=["週報の作成"], wants_tasks=["週報・月報の作成", "鈴木さんへの連絡"]).items():
        form["ans-" + k] = v
    r = client.post("/monthly/submit", data=form)
    html1 = client.get(f"/monthly/done?period={cur}").data.decode("utf-8")
    with client.session_transaction() as sess:
        left = "monthly_last" in sess
    html2 = client.get(f"/monthly/done?period={cur}").data.decode("utf-8")
    check(r.status_code == 302 and "週報・月報の作成" in html1 and "担当者への連絡" in html1 and "鈴木さん" not in html1 and "自動化を検討する" in html1,
          "R2-14: 完了ページの業務名も「・」で分けず、氏名を置き換える")
    check(not left and "自動化を検討する" not in html2, "R2-12: 完了ページは控え（monthly_last）を一度見せたら消す")

    client.post("/survey/api/respondent", json={"emp_id": "S003", "workplace_id": 1, "work_pattern": "普通勤務"})
    r = register("S003", "取り消しの確認")
    tid = r.get_json()["task"]["id"]
    with app.app_context():
        t = db.session.get(Task, tid)
        payload = app_module.task_payload(t)
        stored = t.self_improve
    check(r.status_code == 200 and stored == "" and payload["自分で改善"] == "チェックなし" and r.get_json()["task"]["can_delete"] is True,
          "R2-3: 「一緒にやってみたい」にチェックしなければ空で保存し（改善チームに任せたいとは書かない）、AI に渡すときは「チェックなし」")
    client.post("/login", data={"emp_id": "admin", "password": "adminpass"})
    html_d = client.get(f"/cases/{tid}").data.decode("utf-8")
    html_l = client.get("/cases/").data.decode("utf-8")
    rows = list(csv.reader(io.StringIO(client.get("/cases/export.csv").data.decode("utf-8-sig"))))
    col = rows[0].index("自分で改善")
    row = next(x for x in rows[1:] if x[0] == str(tid))
    check("<th>自分で改善</th><td>チェックなし</td>" in html_d and ">チェックなし</td>" in html_l and row[col] == "チェックなし" and "改善チームに任せたい" not in html_d,
          "R2-3: 案件詳細・案件一覧・案件CSV では空を「チェックなし」と出す")
    client.post("/logout")
    client.post("/survey/api/respondent", json={"emp_id": "S003", "workplace_id": 1, "work_pattern": "普通勤務"})
    touches = [("verified", True), ("assignee", "改善チーム"), ("ecrs_combine", "まとめる"), ("result_note", "済み"), ("progress_note", "確認中"),
               ("after_minutes", 10.0), ("steps_done", "1"), ("assignment", "改善チーム主導"), ("stage_log", None)]
    refused = []
    for field, value in touches:
        with app.app_context():
            t = db.session.get(Task, tid)
            if field == "stage_log":
                t.set_stage("request")
                t.set_stage("start")
            else:
                setattr(t, field, value)
            db.session.commit()
        r = client.delete(f"/survey/api/tasks/{tid}", json={"emp_id": "S003"})
        with app.app_context():
            t = db.session.get(Task, tid)
            refused.append((field, r.status_code == 409 and r.get_json()["message"] == "改善チームが対応を始めているため取り消せません" and app_module._task_json(t)["can_delete"] is False))
            t.verified, t.assignee, t.ecrs_combine, t.result_note, t.progress_note, t.after_minutes, t.steps_done, t.assignment = False, "", "", "", "", None, "", "未定"
            StageLog.query.filter_by(task_id=tid).delete()
            db.session.commit()
    client.post("/login", data={"emp_id": "admin", "password": "adminpass"})
    client.post(f"/cases/{tid}/update", data={"purpose": "日報の数字を課長会議で使うため"})
    client.post("/logout")
    client.post("/survey/api/respondent", json={"emp_id": "S003", "workplace_id": 1, "work_pattern": "普通勤務"})
    r = client.delete(f"/survey/api/tasks/{tid}", json={"emp_id": "S003"})
    with app.app_context():
        t = db.session.get(Task, tid)
        info_refused = (r.status_code == 409 and t is not None and t.purpose == "日報の数字を課長会議で使うため" and bool(t.ai_extra_dict.get("info_edited_at"))
                        and app_module._task_json(t)["can_delete"] is False and t.stage_key == "start" and not t.stage_logs)
        t.ai_extra = json.dumps({k: v for k, v in t.ai_extra_dict.items() if k != "info_edited_at"}, ensure_ascii=False)
        db.session.commit()
    check(info_refused, "R3-11: 改善チームが案件詳細の業務情報（目的など）を直して保存した案件は、開始のままでも本人は取り消せない")
    r = client.delete(f"/survey/api/tasks/{tid}", json={"emp_id": "S003"})
    check(all(ok for _, ok in refused) and r.status_code == 200, f"R2-16: 開始の段階でも改善チームが手を付けた案件（確認・担当者・ECRSメモ・改善内容・ひとこと・段階の記録など）は本人が取り消せない: {refused}")

    client.post("/survey/api/respondent", json={"emp_id": "S002", "workplace_id": 1, "work_pattern": "普通勤務"})
    with app.app_context():
        e = Employee.query.filter_by(emp_id="S002").first()
        e.active = False
        db.session.commit()
    r = register("S002", "対象外になったあと")
    with client.session_transaction() as sess:
        left = "respondent_id" in sess
    j = r.get_json()
    check(r.status_code == 400 and j["message"] == excl and j["errors"]["_"] == "respondent" and j["errors"]["emp_id"] == excl and not left,
          "R2-9: 開いたままの画面でも、対象外になった人は登録できない（対象外のメッセージ）")
    check("if (j.errors.emp_id) { found = false; idProblem = j.errors.emp_id; setIdentity(j.errors.emp_id, true); }" in js, "R2-9: 画面は対象外のメッセージを社員IDの欄に出す")


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
    fe = client.get("/static/flowedit.js").data.decode("utf-8")
    check(all(c.get("checked") is True for c in cands) and "cb.checked = c.checked !== false" in fe, "候補はふつう初めから選ばれている（選ばない候補だけ checked: false）")

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
    with app.app_context():
        t4_id = Task.query.filter_by(name="生産実績の月報作成").first().id
    by_id2 = {e["id"]: e for e in g2["edges"]}
    check(auto_seq["inferred"] is False and by_id2[f"t{t2_id}->t{t4_id}"]["inferred"] is True,
          "反映後: 書き戻した業務どうしの線は実線（名前から推定した点線にしない）。書き戻していない名前の一致は点線のまま")
    check(f"t{t1_id}->t{t2_id}" not in eids, "反映後: 外した前後の矢印は再生成されない")
    r = client.get("/bizmap/api/writeback?kind=info")
    check(len(r.get_json()["candidates"]) <= 2, f"再反映の候補は残らない（{len(r.get_json()['candidates'])}件）")

    r = client.get("/bizmap/edit?kind=info")
    check(r.status_code == 200 and "fe-writeback" in r.data.decode("utf-8"), "editor has writeback button")
    r = client.get("/bizmap/api/writeback?kind=info&domain_id=99")
    check(r.status_code == 200 and r.get_json()["candidates"] == [], "writeback without saved map -> empty")

    # ---- 書き戻しと再生成（手で選んだだれの区分・書き戻した受け取り元・業務名の変更・点線の取り消し） ----
    import app as app_module
    with app.app_context():
        t3 = db.session.get(Task, t3_id)
        t3.output_users_note = "検査・品質保証課 課長"
        t_cnt = Task.query.filter_by(name="出勤簿の集計").first()
        t_cnt.prev_task = "残業予定の連絡"
        db.session.get(Task, t2_id).next_task = "クレーム対応の進捗"
        t1 = db.session.get(Task, t1_id)
        t1.next_task = (t1.next_task or "") + "、点検結果のまとめ"
        monthly_id, cnt_id, wp3, wp1, wp3_name = Task.query.filter_by(name="クレーム対応の進捗報告").first().id, t_cnt.id, t3.workplace_id, t1.workplace_id, t3.workplace.name
        db.session.add(MonthlyResponse(period=current_period(), emp_id="P301", name="", workplace_id=wp3, answers_json=json.dumps({"has": "ある", "task_names": "残業予定の連絡"}, ensure_ascii=False)))
        db.session.add(MonthlyResponse(period=current_period(), emp_id="P302", name="", workplace_id=wp1, answers_json=json.dumps({"has": "ある", "task_names": "点検結果のまとめ"}, ensure_ascii=False)))
        db.session.commit()
    g3 = client.get("/bizmap/api/graph?kind=info").get_json()
    u_check = next(n["id"] for n in g3["nodes"] if n["cls"] == "pending" and n["text"].startswith("点検結果のまとめ"))
    eids3 = {e["id"] for e in g3["edges"]}
    inferred3 = {e["id"]: e["inferred"] for e in g3["edges"]}
    check(f"t{t2_id}->t{monthly_id}" in eids3 and f"t{t1_id}->{u_check}" in eids3 and inferred3[f"t{t2_id}->t{monthly_id}"] is True and inferred3[f"t{t1_id}->{u_check}"] is True,
          "点線: 「クレーム対応の進捗」（業務名の一部）と未登録の業務への前後の業務は名前からつながる")
    data = {"nodes": [], "edges": [dict(e) for e in g3["edges"]], "suppressed_nodes": [], "suppressed_edges": [f"t{t2_id}->t{monthly_id}", f"t{t1_id}->{u_check}"], "view": None}
    for n in g3["nodes"]:
        m = dict(n)
        if n["id"] == f"t{t3_id}":
            m["text"], m["text_edited"] = "残業連絡（Excel）", True
        data["nodes"].append(m)
    data["nodes"] += [
        {"id": "m_ext", "text": "新しい作業", "type": "work", "cls": "work", "who": "山田製作所", "who_edited": True, "actor": "external", "actor_locked": True, "who_state": "ok", "source": "manual"},
        {"id": "m_boss", "text": "新しい作業", "type": "work", "cls": "work", "who": "課長", "who_edited": True, "source": "manual"},
        {"id": "m_src", "text": "検査チェックシート", "type": "info", "cls": "source", "source": "manual"},
    ]
    data["edges"] = [e for e in data["edges"] if e["id"] not in data["suppressed_edges"]] + [
        {"id": f"t{t3_id}->m_ext", "from": f"t{t3_id}", "to": "m_ext", "type": "work", "source": "manual"},
        {"id": f"t{t3_id}->m_boss", "from": f"t{t3_id}", "to": "m_boss", "type": "work", "source": "manual"},
        {"id": f"m_src->t{t1_id}", "from": "m_src", "to": f"t{t1_id}", "type": "info", "source": "manual"},
    ]
    client.post("/bizmap/api/graph", json={"kind": "info", "data": data})
    cands = client.get("/bizmap/api/writeback?kind=info").get_json()["candidates"]
    texts = [c["text"] for c in cands]
    check(any("利用者に「山田製作所（顧客・社外）」を追加" in t for t in texts) and any("利用者に「課長」を追加" in t for t in texts) and any("受け取り元に「検査チェックシート」を追加" in t for t in texts),
          f"書き戻し: 手で選んだだれの区分は名前と一緒に残し、「課長」は「検査・品質保証課 課長」に吸い込まれない: {texts}")
    check(any("「残業連絡（Excel）」に変更" in t and "「出勤簿の集計」の前後の業務に書いた「残業予定の連絡」も直す" in t for t in texts)
          and any("「週報の作成」の後の業務から「クレーム対応の進捗報告」" in t for t in texts) and any("後の業務から「点検結果のまとめ」を外す" in t for t in texts),
          f"書き戻し: 業務名の変更は前後の業務に書いた名前も直す。点線（名前の一部・未登録の業務）を消すと外す候補が出る: {texts}")
    r = client.post("/bizmap/api/writeback", json={"kind": "info", "ids": [c["id"] for c in cands]})
    with app.test_request_context():
        t1, t2, t3, t_cnt = (db.session.get(Task, i) for i in (t1_id, t2_id, t3_id, cnt_id))
        pend3 = [p["name"] for p in app_module.pending_task_names(wp3)]
        check(t3.name == "残業連絡（Excel）" and t_cnt.prev_task == "残業連絡（Excel）" and "残業予定の連絡" not in pend3 and "クレーム対応の進捗" not in (t2.next_task or "")
              and "点検結果のまとめ" not in (t1.next_task or "") and "残業連絡（Excel）" in (t1.next_task or ""),
              f"反映後: 前後の業務の名前も新しい業務名になり、古い名前は未登録の業務に出ない。消した点線の名前は外れる: {t_cnt.prev_task} / {t1.next_task} / {t2.next_task}")
        check("山田製作所（顧客・社外）" in (t3.output_users_note or "") and "課長" in [x.strip() for x in t3.output_users_note.split("、")] and "顧客・社外" in t3.output_users_list
              and "検査チェックシート" in (t1.info_from or ""), f"反映後の業務データ: {t3.output_users_note} / {t3.output_users_list} / {t1.info_from}")
        case = app_module.build_task_flow(t1)
        check(any(n["text"] == "検査チェックシート" for n in case.nodes.values()), "書き戻した受け取り元は、同じ媒体の取り出し元があっても案件の流れ図に必ず描く")
    g4 = client.get("/bizmap/api/graph?kind=info").get_json()
    nodes4 = {n["id"]: n for n in g4["nodes"]}
    edges4 = {(e["from"], e["to"]): e for e in g4["edges"]}
    ext = next((n for n in g4["nodes"] if n.get("who") == "山田製作所" and n["source"] == "auto"), {})
    boss = nodes4.get(app_module.nid("p", f"{t3_id}:課長"), {})
    src_box = next((n for n in g4["nodes"] if n["text"] == "検査チェックシート" and n["source"] == "auto"), {})
    check("m_ext" not in nodes4 and ext.get("actor") == "external" and ext.get("who_state") == "ok" and (f"t{t3_id}", ext.get("id")) in edges4 and (f"o{t3_id}", ext.get("id")) in edges4
          and app_module.nid("p", f"{t3_id}:顧客・社外") not in nodes4 and not any(n.get("who") == "山田製作所" and n.get("who_state") == "unconfirmed" for n in g4["nodes"]),
          f"再生成: 手で選んだ区分の相手は同じ自動の箱（だれ未確認・区分だけの箱は増えない）にまとまり、手の線もその箱につながる: {ext}")
    check("m_boss" not in nodes4 and boss.get("who") == f"{wp3_name} 課長" and (f"o{t3_id}", boss.get("id")) in edges4 and "m_src" not in nodes4 and (src_box.get("id"), f"t{t1_id}") in edges4,
          f"再生成: 書き戻した「課長」と受け取り元は自動の箱になる: {boss.get('who')} / {src_box.get('id')}")
    left = client.get("/bizmap/api/writeback?kind=info").get_json()["candidates"]
    check(not any("山田製作所" in c["text"] or "検査チェックシート" in c["text"] or "残業連絡" in c["text"] for c in left), f"反映した内容は候補に残らない: {[c['text'] for c in left]}")
    check(edges4[(f"t{t1_id}", f"t{t3_id}")]["inferred"] is False, "業務名を変えても、書き戻した業務どうしの線は実線のまま")

    # ---- 書き戻した未登録の業務との線も実線。点線を消して外すと、書き戻した印も消える ----
    data = {"nodes": [dict(n) for n in g4["nodes"]], "edges": [dict(e) for e in g4["edges"]] + [{"id": f"t{t1_id}->{u_check}", "from": f"t{t1_id}", "to": u_check, "type": "work", "source": "manual"}],
            "suppressed_nodes": [], "suppressed_edges": [], "view": None}
    client.post("/bizmap/api/graph", json={"kind": "info", "data": data})
    cands = client.get("/bizmap/api/writeback?kind=info").get_json()["candidates"]
    client.post("/bizmap/api/writeback", json={"kind": "info", "ids": [c["id"] for c in cands if c["id"] == f"next:{t1_id}:{u_check}"]})
    g5 = client.get("/bizmap/api/graph?kind=info").get_json()
    e5 = {e["id"]: e for e in g5["edges"]}
    data = {"nodes": [dict(n) for n in g5["nodes"]], "edges": [dict(e) for e in g5["edges"] if e["id"] != f"t{t1_id}->{u_check}"], "suppressed_nodes": [], "suppressed_edges": [f"t{t1_id}->{u_check}"], "view": None}
    client.post("/bizmap/api/graph", json={"kind": "info", "data": data})
    cands = client.get("/bizmap/api/writeback?kind=info").get_json()["candidates"]
    client.post("/bizmap/api/writeback", json={"kind": "info", "ids": [c["id"] for c in cands if c["id"] == f"unnext:{t1_id}:{u_check}"]})
    with app.test_request_context():
        links = app_module.written_links()
        t1_next = db.session.get(Task, t1_id).next_task or ""
    check(e5.get(f"t{t1_id}->{u_check}", {}).get("inferred") is False and e5.get(f"t{t1_id}->{u_check}", {}).get("source") == "auto"
          and f"t{t1_id}->{u_check}" not in links and f"t{t1_id}->t{t3_id}" in links and "点検結果のまとめ" not in t1_next,
          f"書き戻した未登録の業務との線も実線。点線を消して外すと書き戻した印も消える: {sorted(links)}")

    # ---- 業務名の書き戻し: 同じ名前だけ自動で直し、名前の一部だけが同じものは別の候補（初めは選ばない） ----
    with app.app_context():
        r2 = Respondent.query.filter_by(emp_id="DEMO-002").first()
        p_task = Task(name="部長会議資料の作成", respondent_id=r2.id, workplace_id=r2.workplace_id)
        q_task = Task(name="部長会議の準備", respondent_id=r2.id, workplace_id=r2.workplace_id, next_task="部長会議資料の作成", prev_task="部長会議資料の作成の下書き")
        db.session.add_all([p_task, q_task])
        db.session.commit()
        p_id, q_id = p_task.id, q_task.id
    data = {"nodes": [{"id": f"t{p_id}", "text": "部長会議資料の作成（月次）", "text_edited": True, "type": "work", "cls": "task", "source": "auto"}], "edges": [], "suppressed_nodes": [], "suppressed_edges": [], "view": None}
    client.post("/bizmap/api/graph", json={"kind": "info", "data": data})
    cands = client.get("/bizmap/api/writeback?kind=info").get_json()["candidates"]
    main = next((c for c in cands if c["id"] == f"rename:{p_id}"), {})
    partial = [c for c in cands if c["checked"] is False]
    with app.app_context():
        q_next0 = db.session.get(Task, q_id).next_task
    check(main.get("checked") is True and "「部長会議の準備」の前後の業務に書いた「部長会議資料の作成」も直す" in main.get("text", "") and "生産実績の月報作成" not in main.get("text", "")
          and any("「生産実績の月報作成」の後の業務に書いた「部長会議」を「部長会議資料の作成（月次）」に直す" in c["text"] for c in partial)
          and any("「部長会議の準備」の前の業務に書いた「部長会議資料の作成の下書き」を「部長会議資料の作成（月次）の下書き」に直す" in c["text"] for c in partial)
          and len(partial) == 2 and q_next0 == "部長会議資料の作成",
          f"業務名の書き戻し: 前後の業務は同じ名前だけ一緒に直し、名前の一部だけが同じもの（「部長会議」など）は直し方を書いた別の候補（初めは選ばない）: {[c['text'] for c in partial]}")
    pick = [c["id"] for c in cands if c["checked"]] + [c["id"] for c in partial if "の下書き" in c["text"]]
    client.post("/bizmap/api/writeback", json={"kind": "info", "ids": pick})
    with app.app_context():
        p_name, t4_next, q_next, q_prev = db.session.get(Task, p_id).name, db.session.get(Task, t4_id).next_task, db.session.get(Task, q_id).next_task, db.session.get(Task, q_id).prev_task
    check(p_name == "部長会議資料の作成（月次）" and t4_next == "部長会議" and q_next == "部長会議資料の作成（月次）" and q_prev == "部長会議資料の作成（月次）の下書き",
          f"反映後: 同じ名前は新しい業務名に、選んだ名前の一部の候補は書いたとおりに直り、選ばなかった「部長会議」はそのまま: {p_name} / {t4_next} / {q_next} / {q_prev}")

    with app.test_request_context():
        tasks_by_id = {t.id: t for t in Task.query.all()}
        saved = {"schema": 2, "nodes": [{"id": f"t{t2_id}", "text": "佐藤さんの週報", "text_edited": True, "type": "work", "cls": "task", "source": "auto"},
                                        {"id": "m_p", "text": "鈴木課長の確認", "type": "work", "cls": "work", "who": "", "source": "manual"}],
                 "edges": [{"id": "x", "from": f"t{t2_id}", "to": "m_p", "type": "work", "source": "manual"}]}
        cands = app_module.writeback_candidates(saved, tasks_by_id)
        app_module.apply_writeback(cands, {c["id"] for c in cands}, tasks_by_id)
        t2 = tasks_by_id[t2_id]
        check(t2.name == "担当者の週報" and "課長の確認" in (t2.next_task or "") and "鈴木" not in (t2.next_task or ""), f"書き戻す値も氏名を置き換える: {t2.name} / {t2.next_task}")
        db.session.rollback()

    # ---- 番号の使い回しをしない（業務・職場・領域。消した番号の手修正が新しい行に付かない） ----
    with app.app_context():
        r0 = Respondent.query.first()
        t = Task(name="番号の確かめ1", respondent_id=r0.id, workplace_id=r0.workplace_id)
        db.session.add(t)
        db.session.commit()
        first_id = t.id
        db.session.delete(t)
        db.session.commit()
        t = Task(name="番号の確かめ2", respondent_id=r0.id, workplace_id=r0.workplace_id)
        db.session.add(t)
        db.session.commit()
        second_id = t.id
    client.post("/admin/workplace/add", data={"name": "番号確認の職場"})
    client.post("/admin/domain/add", data={"name": "番号確認の領域"})
    with app.app_context():
        w_id = app_module.Workplace.query.filter_by(name="番号確認の職場").first().id
        d_id = app_module.Domain.query.filter_by(name="番号確認の領域").first().id
    client.post(f"/admin/workplace/{w_id}/delete")
    client.post(f"/admin/domain/{d_id}/delete")
    client.post("/admin/workplace/add", data={"name": "番号確認の職場2"})
    client.post("/admin/domain/add", data={"name": "番号確認の領域2"})
    with app.app_context():
        w2 = app_module.Workplace.query.filter_by(name="番号確認の職場2").first().id
        d2 = app_module.Domain.query.filter_by(name="番号確認の領域2").first().id
        big = (db.session.query(db.func.max(Task.id)).scalar() or 0) + 50
        db.session.add(FlowMap(kind="info", scope="wp:0|dom:77", data=json.dumps({"nodes": [{"id": f"t{big}", "group": "wp66"}], "edges": [], "suppressed_edges": [f"t1->o{big + 5}"]})))
        db.session.commit()
    check(second_id > first_id and w2 > w_id and d2 > d_id, f"消した業務・職場・領域の番号は使い回さない: 業務 {first_id}→{second_id} 職場 {w_id}→{w2} 領域 {d_id}→{d2}")
    app2 = create_app({"SQLALCHEMY_DATABASE_URI": "sqlite:///" + dbfile, "TESTING": True})
    with app2.app_context():
        t = Task(name="番号の確かめ3", respondent_id=Respondent.query.first().id)
        w = app_module.Workplace(name="番号確認の職場3")
        dm = app_module.Domain(name="番号確認の領域3")
        db.session.add_all([t, w, dm])
        db.session.commit()
        check(t.id > big + 5 and w.id > 66 and dm.id > 77, f"起動時に手修正が指している番号より大きい番号から始める: 業務 {t.id} 職場 {w.id} 領域 {dm.id}")

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
        with app.app_context():
            plan_at = db.session.get(Task, task_id).ai_extra_dict.get("ai_plan_at") or ""
        script.append(http_error(401, '{"error":{"message":"invalid api key"}}'))
        r = client.post(f"/cases/{task_id}/reanalyze", follow_redirects=True)
        html = r.data.decode("utf-8")
        with app.app_context():
            t_fb = db.session.get(Task, task_id)
            note, src = t_fb.ai_plan_note, t_fb.ai_source
        check(plan_at.startswith("20") and src == "rules" and note == f"前回の AI 分析（{plan_at}）" and "AI から使える応答がなかったため、ルールベースで分析しました（HTTP 401" in html
              and ">ルールベース</span>" in html and f'<span class="badge text-bg-light border">{note}</span> 前回の AI 分析で出た案です' in html and "新しい流れ1" in html
              and '<span class="badge badge-flag">推定</span>' in html and '<span class="badge badge-flag">AI</span>' not in html,
              f"再分析がルールに落ちたら、そのことを知らせ、前回の AI の改善案は日付つきで分けて出す（いまの分析のように見せない。関連業務もルールの推定と出す）: {note}")
        csv_text = client.get("/cases/export.csv").data.decode("utf-8-sig")
        check(f"［{note}］新しい流れ1" in csv_text, "CSV の改善後の流れ（AI案）にも、前回の AI 分析の案であることと日付を付ける")
        script.append(message(ai_json))
        r = client.post(f"/cases/{task_id}/reanalyze", follow_redirects=True)
        html = r.data.decode("utf-8")
        check("AI分析を更新しました（ChatGPT API）" in html and "前回の AI 分析" not in html and "改善後の流れ（AI案）" in html and '<span class="badge badge-flag">AI</span>' in html,
              "AI で再分析し直すと、改善案はいまの分析の案に戻る（前回の印は消え、関連業務は AI の印）")
        with app.app_context():
            t_ai = db.session.get(Task, task_id)
            before = (t_ai.ai_summary, t_ai.ai_at, t_ai.ai_priority, t_ai.ai_source)
            stamp = t_ai.ai_at.strftime("%Y/%m/%d %H:%M")
        n_calls = len(calls)
        r = client.post(f"/cases/{task_id}/update", data={"purpose": "週の生産状況を部長会議で共有するため（テスト）"}, follow_redirects=True)
        html = r.data.decode("utf-8")
        with app.app_context():
            t_ai = db.session.get(Task, task_id)
            after = (t_ai.ai_summary, t_ai.ai_at, t_ai.ai_priority, t_ai.ai_source)
            stale, edited = t_ai.ai_stale_note, t_ai.ai_extra_dict.get("info_edited_at")
        check(len(calls) == n_calls and after == before and stale == f"業務情報を直す前の AI 分析（{stamp}）" and edited
              and "業務情報を直しました。AI 分析を新しくするには「再分析」を押してください" in html and f'<span class="badge text-bg-light border">{stale}</span>' in html and ">ChatGPT API</span>" in html,
              f"R3-10: 業務情報を直して保存しても AI の分析はそのまま残し（AI は呼ばない）、「業務情報を直す前の AI 分析（日時）」の印と「再分析」を促す知らせを出す: {stale}")
        csv_text = client.get("/cases/export.csv").data.decode("utf-8-sig")
        check(f"［{stale}］テスト要約" in csv_text, "R3-10: 案件CSVの AI要約にも「業務情報を直す前の AI 分析（日時）」の印を付ける")
        script.append(message(ai_json))
        r = client.post(f"/cases/{task_id}/reanalyze", follow_redirects=True)
        html = r.data.decode("utf-8")
        with app.app_context():
            t_ai = db.session.get(Task, task_id)
            still_edited, stale_after = t_ai.ai_extra_dict.get("info_edited_at"), t_ai.ai_stale_note
        check("AI分析を更新しました（ChatGPT API）" in html and "業務情報を直す前の AI 分析" not in html and stale_after == "" and still_edited == edited,
              "R3-10・R3-11: 再分析すると印は消える（改善チームが業務情報を直した印は残す）")
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
