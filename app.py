"""PC定型業務ゼロ化プロジェクト ポータル（Flask + SQLite）。アプリ本体はこの1ファイル

    flask --app app run --host 0.0.0.0 --port 5000                   開発用サーバー（ポートは --port で指定）
    waitress-serve --host 0.0.0.0 --port 5000 --call app:create_app   本番用サーバー（waitress）
    python app.py / python app.py serve                               上と同じ。ポートは config.py の PORT
    python tests.py        スモークテスト

設定は config.py、社員ID照会は ldap_auth.py、データは instance/portal.sqlite3。
必要なパッケージ: pip install Flask Flask-SQLAlchemy waitress（AI連携を使うときは anthropic も）

このファイルの構成（検索用の見出し）:
  [設定] [選択肢・係数] [自動化依頼のフォーム定義] [データモデル] [社員ID照会の控え] [集計・分析]
  [業務マップ] [AI連携] [初期データ・デモ・移行] [画面: ログイン / ダッシュボード / 自動化依頼 / 改善案件 / 業務マップ /
  DX人財情報 / 管理 / 月次アンケート / 活動状況] [テンプレート] [静的ファイル] [アプリ生成] [起動]
"""
import csv
import hashlib
import hmac
import io
import json
import logging
import mimetypes
import os
import re
import secrets
import sys
from collections import Counter, OrderedDict, defaultdict
from datetime import datetime, timedelta

from flask import Blueprint, Flask, Response, abort, flash, jsonify, redirect, render_template, request, session, url_for
from flask_sqlalchemy import SQLAlchemy
from jinja2 import DictLoader

import config as _config
import ldap_auth

log = logging.getLogger("portal")
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 以前はモジュールに分かれていた名前空間。1ファイルになったので、このモジュール自身を指す
C = survey_spec = sys.modules[__name__]


# ==============================================================================
# 設定（config.py の読み込み）
# ==============================================================================
class _Conf:
    """config.py の値を読む。conf.get(名前, 既定値)。テストは conf.override(名前, 値) で一時的に差し替える（config.py は書き換えない）"""

    def get(self, name, default=None):
        value = getattr(_config, name, default)
        return default if value is None else value

    def override(self, name, value):
        setattr(_config, name, value)  # ldap_auth も同じ config モジュールを見るので、そちらにも効く


conf = _Conf()


# ==============================================================================
# 選択肢・係数・しきい値（アンケート設計書に準拠）
# ==============================================================================
SEP = "|"  # 複数選択の区切り文字


def split_list(value):
    if not value:
        return []
    return [v for v in str(value).split(SEP) if v]


def join_list(values):
    return SEP.join([v for v in values if v])


# ---- パートB 業務の棚卸し ----
OUTPUT_USERS = ["自分", "職場内", "上司", "他部署", "顧客・社外", "分からない"]
INPUT_SOURCES = ["生産管理システム", "品質システム", "設備データ", "紙の帳票", "メール", "他のExcel", "人から聞く", "その他"]
FREQUENCIES = ["毎日", "週に数回", "毎週", "毎月", "四半期・年次", "不定期"]
DURATIONS = ["15分未満", "15〜30分", "30分〜1時間", "1〜2時間", "2時間以上"]
TOOLS = ["Excel", "Access", "メール", "社内システム", "紙", "VBA・マクロ", "RPA", "その他"]
NATURES = ["転記・コピー", "集計・グラフ作成", "資料作成", "確認・チェック", "データ取得", "メール・配布", "台帳更新", "判断を伴う"]
PAINS = ["時間がかかる", "ミスが起きやすい", "同じ情報を何度も入力する", "手順が人によって違う", "急ぎで割り込まれる", "何のためか分からない", "特になし"]
WISHES = ["なくしたい", "減らしたい", "自動化したい", "このままでよい", "分からない"]
# ヒト・情報の流れ（B17・B18）
SYSTEM_SOURCES = {"生産管理システム", "品質システム", "設備データ"}
SELF_IMPROVE = ["取り組みたい", "教えてもらえれば取り組みたい", "改善チームに任せたい", "分からない"]

# ---- パートC 改善への参加意欲とスキル ----
C1_OPTIONS = ["ぜひ取り組みたい", "教えてもらえれば取り組みたい", "興味はあるが時間がない", "今は考えていない"]
SKILLS = ["Excel関数", "ピボットテーブル", "Excel VBA", "Access", "Power Query", "Power BI", "Python", "RPA", "生成AI", "特になし"]
C4_HOURS = ["1時間未満", "1〜2時間", "2〜4時間", "4時間以上"]
C6_OPTIONS = ["できる", "内容による", "難しい"]

# ---- 月間時間の算出（設計書 10. 集計・活用の流れ） ----
FREQ_COEF = {
    "毎日": 20.0,
    "週に数回": 10.0,
    "毎週": 4.0,
    "毎月": 1.0,
    "四半期・年次": 0.3,
    "不定期": 0.0,  # 個別確認
}
DURATION_MINUTES = {
    "15分未満": 10.0,
    "15〜30分": 22.0,
    "30分〜1時間": 45.0,
    "1〜2時間": 90.0,
    "2時間以上": 150.0,  # 個別確認（暫定値）
}

# ---- 案件管理 ----
STATUSES = ["受付", "対応中", "完了", "保留"]
DECISIONS = ["未定", "廃止", "統合", "簡素化", "自動化", "高度化", "現状維持"]
ASSIGNMENTS = ["未定", "改善チーム主導", "自主改善"]
STEPS = [
    (1, "アウトプットの確認", "その仕事が何を生み出し、誰がどう使っているかを確認する"),
    (2, "目的の明確化", "なぜやるのかを明確にする。目的不明・未利用なら廃止する"),
    (3, "ECRSによる最適化", "排除・結合・入替・簡素化でムリ・ムダ・ムラを取り除く"),
    (4, "自動化", "最適化後に残った繰り返し作業を自動化する"),
    (5, "高度化", "人が価値を発揮する業務へ役割を移す"),
]
ECRS = [
    ("eliminate", "排除（Eliminate）", "その作業をなくせないか"),
    ("combine", "結合（Combine）", "他の作業と一緒にできないか"),
    ("rearrange", "入替（Rearrange）", "順序・担当・場所を変えられないか"),
    ("simplify", "簡素化（Simplify）", "もっと簡単にできないか"),
]

# ---- 活動状況（製造部全体に公開する、案件の進み具合）----
# key, 表示名, この段階でやること, 対応する標準手順
STAGES = [
    ("start", "開始", "依頼を受け付け、活動状況に載せる", ""),
    ("request", "要求定義", "何に困っていて、何を実現したいかを依頼者と整理する", "手順1・2（アウトプットの確認・目的の明確化）"),
    ("spec", "要件定義", "ECRSで業務を最適化し、残る作業の仕様を決める", "手順3（ECRSによる最適化）"),
    ("build", "作成", "ツール・仕組みを作る。自主改善の案件は改善チームが伴走する", "手順4・5（自動化・高度化）"),
    ("test", "テスト", "実際のデータで動作と効果を確かめる", ""),
    ("check", "確認", "依頼者・職場で結果を確認し、運用に乗せる", ""),
    ("done", "完了！", "運用開始。改善前後の結果を記録する", ""),
]
STAGE_KEYS = [s[0] for s in STAGES]
STAGE_INDEX = {k: i for i, k in enumerate(STAGE_KEYS)}
STAGE_NAMES = {s[0]: s[1] for s in STAGES}
BUILD_STAGE_SELF = "伴走"  # 自主改善の案件では「作成」を「伴走」と表示する

# 改善案件候補とみなす改善希望
CANDIDATE_WISHES = {"なくしたい", "減らしたい", "自動化したい"}
SELF_IMPROVE_POSITIVE = {"取り組みたい", "教えてもらえれば取り組みたい"}
C1_POSITIVE = {"ぜひ取り組みたい", "教えてもらえれば取り組みたい"}

# ---- 分析不足領域の判定しきい値 ----
# (最小件数, 最小充足率) を満たせば次の段階へ
ANALYSIS_LEVELS = [
    # level, 分析状態, 業務把握, 判定（件数, 充足率）
    (3, "分析済み", "十分", 8, 0.8),
    (2, "追加調査", "一部", 3, 0.5),
    (1, "情報不足", "不十分", 1, 0.0),
    (0, "未分析", "不明", 0, 0.0),
]

# ---- 学習コンテンツ ----

AI_FLAGS = ["目的不明", "廃止候補", "統合候補", "重複入力候補", "自動化候補", "判断業務", "要ヒアリング"]

# ---- 月次アンケート（voiceloop 方式: 月度ごとの設問セット。基本設問は毎月同じで KPI につなぐ） ----
SHIFTS_PER_MONTH = 20  # 1か月の出勤日数の目安（推計月間時間 ＝ 1シフトあたり時間の換算値 × これ）
MONTHLY_HOURS_VALUE = {"0時間": 0.0, "30分未満": 0.25, "30分〜1時間": 0.75, "1〜2時間": 1.5, "2〜4時間": 3.0, "4時間以上": 5.0}  # 1シフトあたりの換算値（時間）
MONTHLY_CORE_QUESTIONS = [
    {"key": "has", "type": "choice", "text": "PCでおこなう定型業務（繰り返しの手動作業）はありますか", "options": ["ある", "少しある", "ない"], "required": True},
    {"key": "hours", "type": "choice", "text": "1シフト（1出勤）あたり、PCの定型業務を行っている時間は", "options": list(MONTHLY_HOURS_VALUE), "required": True},
    {"key": "trend", "type": "choice", "text": "先月と比べて、PC定型業務は", "options": ["減った", "変わらない", "増えた"], "required": True},
    {"key": "new_tasks", "type": "text", "text": "増えたPC定型業務名", "required": False, "placeholder": "例：新しい報告フォーマットへの転記",
     "depends": {"key": "trend", "show_if": ["増えた"], "required": True}},  # M3 が「増えた」のときだけ出す（出たら必須）
    {"key": "wants", "type": "choice", "text": "なくしたい・自動化したいPC定型業務はありますか", "options": ["ある", "ない"], "required": True},
    {"key": "wants_tasks", "type": "text", "text": "あれば、その業務名（1行に1件）", "required": False, "placeholder": "例：日報の実績をExcelに転記する",
     "depends": {"key": "wants", "show_if": ["ある"], "required": True}},  # M4 が「ある」のときだけ出す（出たら必須）
    {"key": "stance", "type": "choice", "text": "どのように進めたいですか", "options": ["自分で取り組みたい", "改善チームに依頼したい"], "required": False,
     "depends": {"key": "wants", "show_if": ["ある"], "required": True}},  # M4 が「ある」のときだけ出す（出たら必須）。頼みたいタイプか自分でやりたいタイプかを見る
]
MONTHLY_CORE_KEYS = {q["key"] for q in MONTHLY_CORE_QUESTIONS}
MONTHLY_KPI_LABELS = {
    "has": "繰り返し業務あり率",
    "hours": "推計 月間時間",
    "trend": "減った率・増えた率",
    "wants": "自動化したい率",
    "wants_tasks": "挙げられた業務（本登録へ誘導）",
    "stance": "依頼したい率・自分でやりたい率",
    "new_tasks": "新規業務の件数",
}
MONTHLY_KPI = {
    "has_positive": {"ある", "少しある"},
    "has_strong": {"ある"},
    "trend_decreased": {"減った"},
    "trend_increased": {"増えた"},
    "wants_positive": {"ある"},
    "ask_positive": {"改善チームに依頼したい"},
    "self_positive": {"自分で取り組みたい"},
}


# ==============================================================================
# 自動化依頼（1画面）のフォーム定義
# ==============================================================================
# ------------------------------------------------------------
# 自動化依頼（1画面）のフォーム定義。画面はこれを読んで入力欄を描く
#   type: text / textarea / number / select / choice（1つ選ぶ）/ multi（複数選ぶ）
#   options: 選択肢のリスト、または "workplaces" / "domains"（マスタから）
# ------------------------------------------------------------
PERSON_FIELDS = [
    {"key": "emp_id", "type": "text", "label": "社員ID", "required": True, "lookup": True, "placeholder": "例：E001"},
    # 氏名・職場は社員IDから照会して表示する。見つからないときだけ入力欄を出す
    {"key": "name", "type": "text", "label": "氏名", "required": True, "fallback": True},
    {"key": "workplace_id", "type": "select", "label": "職場", "required": True, "options": "workplaces"},
]

TASK_SECTIONS = [
    {
        "title": "何を・何のために",
        "fields": [
            {"key": "name", "type": "text", "label": "業務名", "required": True, "placeholder": "例：日報の実績をExcelの集計表に転記する"},
            {"key": "domain_id", "type": "choice", "label": "業務領域", "required": True, "options": "domains"},
            {"key": "purpose", "type": "text", "label": "目的", "required": True, "placeholder": "例：日次の生産実績を朝会で確認するため（分からなければ「不明」）"},
            {"key": "output", "type": "text", "label": "できるもの", "required": True, "placeholder": "例：日次実績集計表（Excel）"},
            {"key": "output_users", "type": "multi", "label": "使う人", "required": True, "options": C.OUTPUT_USERS, "note": {"key": "output_users_note", "placeholder": "部署名や用途（任意）"}},
            {"key": "inputs", "type": "multi", "label": "情報のもと", "required": True, "options": C.INPUT_SOURCES, "note": {"key": "inputs_note", "placeholder": "システム名やファイル名（任意）"}},
        ],
    },
    {
        "title": "どれくらい",
        "fields": [
            {"key": "frequency", "type": "choice", "label": "頻度", "required": True, "options": C.FREQUENCIES},
            {"key": "duration", "type": "choice", "label": "1回の時間", "required": True, "options": C.DURATIONS},
            {"key": "headcount", "type": "number", "label": "同じ作業をしている人数", "required": False, "min": 1, "default": 1, "unit": "人"},
        ],
    },
    {
        "title": "やり方と困りごと",
        "fields": [
            {"key": "tools", "type": "multi", "label": "使うツール", "required": True, "options": C.TOOLS},
            {"key": "nature", "type": "multi", "label": "作業の種類", "required": True, "options": C.NATURES},
            {"key": "pains", "type": "multi", "label": "困りごと", "required": True, "options": C.PAINS, "note": {"key": "pains_note", "placeholder": "具体的に（任意）"}},
            {"key": "wish", "type": "choice", "label": "どうしたい", "required": True, "options": C.WISHES},
            {"key": "self_improve", "type": "choice", "label": "自分でやってみたい？", "required": True, "options": C.SELF_IMPROVE},
            {"key": "note", "type": "textarea", "label": "補足", "required": False, "placeholder": "改善案や気づいたこと（任意）"},
        ],
    },
    {
        "title": "つながり（前後の業務・関わる人・情報の受け取り元）",
        "optional": True,
        "collapsed": True,
        "fields": [
            {"key": "prev_task", "type": "text", "label": "前の業務", "placeholder": "例：各ラインの日報記入"},
            {"key": "next_task", "type": "text", "label": "後の業務", "placeholder": "例：週報の作成"},
            {"key": "actors", "type": "text", "label": "関わる人・部署", "placeholder": "例：各ライン作業者、課長、生産管理課"},
            {"key": "info_from", "type": "text", "label": "情報の受け取り元", "placeholder": "例：営業からの受注メール、品質システムの検査データ"},
        ],
    },
]

C_FIELDS = [
    {"key": "c1_wish", "type": "choice", "label": "自分で改善に取り組みたい？", "required": True, "options": C.C1_OPTIONS},
    {"key": "c2_skills", "type": "multi", "label": "使える技術", "required": True, "options": C.SKILLS},
    {"key": "c3_learn", "type": "multi", "label": "学びたい技術", "options": C.SKILLS},
    {"key": "c4_hours", "type": "choice", "label": "改善に使える時間（週）", "options": C.C4_HOURS},
    {"key": "c5_targets", "type": "text", "label": "取り組みたい業務", "placeholder": "登録した業務名（複数可）"},
    {"key": "c6_share", "type": "choice", "label": "事例の発表・共有", "options": C.C6_OPTIONS},
]

D_FIELDS = [
    {"key": "d1", "type": "textarea", "label": "他部署とのやりとりで手間なこと", "placeholder": "例：毎月〇〇へ実績を送る、△△からの依頼を転記する"},
    {"key": "d2", "type": "textarea", "label": "同じ情報を何度も入力しているもの", "placeholder": "例：日報の数字を日次表・週報・月報に3回入れている"},
    {"key": "d3", "type": "textarea", "label": "誰が何に使うか分からない帳票・報告", "placeholder": "例：毎週出している〇〇集計表"},
    {"key": "d4", "type": "textarea", "label": "PC定型業務について感じていること", "placeholder": "自由にどうぞ"},
]


# ==============================================================================
# データモデル（SQLite）
# ==============================================================================
db = SQLAlchemy()


class Workplace(db.Model):
    """職場マスタ"""

    __tablename__ = "workplaces"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), unique=True, nullable=False)
    sort = db.Column(db.Integer, default=0)

    respondents = db.relationship("Respondent", back_populates="workplace")


class Domain(db.Model):
    """業務領域（業務体系の大分類・中分類）"""

    __tablename__ = "domains"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    description = db.Column(db.String(300), default="")
    sort = db.Column(db.Integer, default=0)
    parent_id = db.Column(db.Integer, db.ForeignKey("domains.id"), nullable=True)

    parent = db.relationship("Domain", remote_side=[id], backref="children")
    tasks = db.relationship("Task", back_populates="domain")

    @property
    def is_top(self):
        return self.parent_id is None


class Respondent(db.Model):
    """回答者（自動化依頼を出した人。あなたのこと・関わり方・全体のこと）"""

    __tablename__ = "respondents"
    id = db.Column(db.Integer, primary_key=True)
    emp_id = db.Column(db.String(50))  # 社員ID（社員マスタ・LDAP照会と紐づく）
    name = db.Column(db.String(100), nullable=False)
    workplace_id = db.Column(db.Integer, db.ForeignKey("workplaces.id"))
    created_at = db.Column(db.DateTime, default=datetime.now)
    is_demo = db.Column(db.Boolean, default=False)

    # パートC
    c1_wish = db.Column(db.String(50))
    c2_skills = db.Column(db.Text, default="")
    c3_learn = db.Column(db.Text, default="")
    c4_hours = db.Column(db.String(50))
    c5_targets = db.Column(db.Text, default="")
    c6_share = db.Column(db.String(50))
    part_c_done = db.Column(db.Boolean, default=False)

    # パートD
    d1 = db.Column(db.Text, default="")
    d2 = db.Column(db.Text, default="")
    d3 = db.Column(db.Text, default="")
    d4 = db.Column(db.Text, default="")
    part_d_done = db.Column(db.Boolean, default=False)

    workplace = db.relationship("Workplace", back_populates="respondents")
    tasks = db.relationship("Task", back_populates="respondent", cascade="all, delete-orphan")

    @property
    def skills(self):
        return C.split_list(self.c2_skills)

    @property
    def learn(self):
        return C.split_list(self.c3_learn)

    @property
    def wants_to_improve(self):
        """自主改善希望者か（C1 または いずれかの業務の B15）"""
        if self.c1_wish in C.C1_POSITIVE:
            return True
        return any(t.self_improve in C.SELF_IMPROVE_POSITIVE for t in self.tasks)

    @property
    def total_minutes(self):
        return sum(t.monthly_minutes or 0 for t in self.tasks)


class Task(db.Model):
    """PC業務（1業務＝1件）。登録された業務はそのまま改善案件の候補になる。"""

    __tablename__ = "tasks"
    id = db.Column(db.Integer, primary_key=True)
    respondent_id = db.Column(db.Integer, db.ForeignKey("respondents.id"), nullable=False)
    workplace_id = db.Column(db.Integer, db.ForeignKey("workplaces.id"))
    is_demo = db.Column(db.Boolean, default=False)

    # パートB
    name = db.Column(db.String(200), nullable=False)  # B1
    domain_id = db.Column(db.Integer, db.ForeignKey("domains.id"))  # B2
    purpose = db.Column(db.Text, default="")  # B3
    output = db.Column(db.Text, default="")  # B4
    output_users = db.Column(db.Text, default="")  # B5 複数選択
    output_users_note = db.Column(db.Text, default="")
    inputs = db.Column(db.Text, default="")  # B6 複数選択
    inputs_note = db.Column(db.Text, default="")
    prev_task = db.Column(db.Text, default="")  # B7
    next_task = db.Column(db.Text, default="")
    frequency = db.Column(db.String(30))  # B8
    duration = db.Column(db.String(30))  # B9
    headcount = db.Column(db.Integer, default=1)  # B10
    tools = db.Column(db.Text, default="")  # B11
    nature = db.Column(db.Text, default="")  # B12
    pains = db.Column(db.Text, default="")  # B13
    pains_note = db.Column(db.Text, default="")
    wish = db.Column(db.String(30))  # B14
    self_improve = db.Column(db.String(50))  # B15
    note = db.Column(db.Text, default="")  # B16

    # ヒト・情報の流れ（B17・B18）
    actors = db.Column(db.Text, default="")  # B17 関わる人・部署（「、」区切り）
    info_from = db.Column(db.Text, default="")  # B18 情報の受け取り元（部署・人・システム・帳票）

    monthly_minutes = db.Column(db.Float, default=0.0)

    # AI分析
    ai_domain = db.Column(db.String(100))
    ai_related = db.Column(db.Text)  # JSON list
    ai_summary = db.Column(db.Text)
    ai_ecrs = db.Column(db.Text)  # JSON dict
    ai_flags = db.Column(db.Text)  # JSON list
    ai_priority = db.Column(db.String(10))
    ai_source = db.Column(db.String(20))
    ai_at = db.Column(db.DateTime)

    # 案件管理
    status = db.Column(db.String(20), default="受付")
    steps_done = db.Column(db.String(50), default="")  # "1|2|3"
    decision = db.Column(db.String(20), default="未定")
    assignment = db.Column(db.String(20), default="未定")
    assignee = db.Column(db.String(100), default="")
    ecrs_eliminate = db.Column(db.Text, default="")
    ecrs_combine = db.Column(db.Text, default="")
    ecrs_rearrange = db.Column(db.Text, default="")
    ecrs_simplify = db.Column(db.Text, default="")
    after_minutes = db.Column(db.Float, nullable=True)  # 改善後の月間時間
    result_note = db.Column(db.Text, default="")
    completed_at = db.Column(db.DateTime, nullable=True)
    verified = db.Column(db.Boolean, default=False)  # 推進メンバー確認済み
    verified_by = db.Column(db.String(100), default="")

    # 活動状況（開始 → 要求定義 → 要件定義 → 作成/伴走 → テスト → 確認 → 完了）
    stage = db.Column(db.String(20), default="start")
    stage_at = db.Column(db.DateTime, nullable=True)  # 現在の段階に入った日時
    progress_note = db.Column(db.Text, default="")  # 活動状況に出すひとこと（今の状況・次の予定）

    created_at = db.Column(db.DateTime, default=datetime.now)
    updated_at = db.Column(db.DateTime, default=datetime.now, onupdate=datetime.now)

    respondent = db.relationship("Respondent", back_populates="tasks")
    domain = db.relationship("Domain", back_populates="tasks")
    workplace = db.relationship("Workplace")
    stage_logs = db.relationship("StageLog", back_populates="task", cascade="all, delete-orphan", order_by="StageLog.at")

    # ---- 派生値 ----
    def compute_minutes(self):
        coef = C.FREQ_COEF.get(self.frequency or "", 0.0)
        minutes = C.DURATION_MINUTES.get(self.duration or "", 0.0)
        head = self.headcount or 1
        self.monthly_minutes = round(coef * minutes * max(head, 1), 1)
        return self.monthly_minutes

    @property
    def output_users_list(self):
        return C.split_list(self.output_users)

    @property
    def inputs_list(self):
        return C.split_list(self.inputs)

    @property
    def tools_list(self):
        return C.split_list(self.tools)

    @property
    def nature_list(self):
        return C.split_list(self.nature)

    @property
    def pains_list(self):
        return C.split_list(self.pains)

    @property
    def steps_done_list(self):
        return [int(s) for s in C.split_list(self.steps_done) if s.isdigit()]

    @property
    def actors_list(self):
        import re

        return [a.strip() for a in re.split(r"[、,，/／\n;；]+", self.actors or "") if a.strip()]

    @property
    def has_flow_info(self):
        return bool(self.actors or self.info_from)

    @property
    def required_filled(self):
        """B2〜B9 の必須項目が埋まっているか（業務情報の充足判定）"""
        return bool(
            self.domain_id
            and (self.purpose or "").strip()
            and (self.output or "").strip()
            and self.output_users
            and self.inputs
            and self.frequency
            and self.duration
        )

    @property
    def is_sufficient(self):
        return self.required_filled and self.verified

    @property
    def purpose_unknown(self):
        p = (self.purpose or "").strip()
        return (not p) or ("不明" in p) or ("分からない" in p) or ("わからない" in p)

    @property
    def is_candidate(self):
        """改善案件候補か（改善希望あり、または困りごとあり）"""
        if self.wish in C.CANDIDATE_WISHES:
            return True
        pains = [p for p in self.pains_list if p != "特になし"]
        return bool(pains)

    @property
    def is_elimination_candidate(self):
        """廃止・統合候補か（目的不明・利用者不明・重複入力の指摘）"""
        return (
            self.purpose_unknown
            or "分からない" in self.output_users_list
            or "何のためか分からない" in self.pains_list
            or "同じ情報を何度も入力する" in self.pains_list
        )

    # ---- 進捗段階（活動状況） ----
    @property
    def stage_key(self):
        return self.stage if self.stage in C.STAGE_INDEX else "start"

    @property
    def stage_index(self):
        return C.STAGE_INDEX[self.stage_key]

    def stage_label(self, key):
        """段階の表示名。自主改善の案件では「作成」を「伴走」と表示する"""
        if key == "build" and self.assignment == "自主改善":
            return C.BUILD_STAGE_SELF
        return C.STAGE_NAMES.get(key, key)

    @property
    def stage_name(self):
        return self.stage_label(self.stage_key)

    @property
    def is_on_hold(self):
        return self.status == "保留"

    @property
    def stage_dates(self):
        """段階ごとに最後に入った日時 {key: datetime}。「開始」は登録日で補う"""
        dates = {"start": self.created_at}
        for log in self.stage_logs:
            dates[log.stage] = log.at
        if self.stage_at and self.stage_key not in dates:
            dates[self.stage_key] = self.stage_at
        return dates

    @property
    def stage_since(self):
        return self.stage_at or self.stage_dates.get(self.stage_key) or self.created_at

    def set_stage(self, key, by="", note=""):
        """段階を変える。ステータスを同期し、履歴（StageLog）を残す。変わったとき True"""
        if key not in C.STAGE_INDEX:
            return False
        changed = key != self.stage_key
        if changed:
            now = datetime.now()
            self.stage = key
            self.stage_at = now
            self.stage_logs.append(StageLog(stage=key, at=now, by=by or "", note=note or ""))
        # ステータスの同期（保留はそのまま）
        if key == "done":
            if self.status != "完了":
                self.completed_at = datetime.now()
            self.status = "完了"
        else:
            if self.status != "保留":
                self.status = "受付" if key == "start" else "対応中"
            self.completed_at = None
        return changed

    @property
    def saved_minutes(self):
        if self.status == "完了" and self.after_minutes is not None:
            return max((self.monthly_minutes or 0) - self.after_minutes, 0)
        return 0.0

    @property
    def ai_related_list(self):
        try:
            return json.loads(self.ai_related) if self.ai_related else []
        except ValueError:
            return []

    @property
    def ai_ecrs_dict(self):
        try:
            return json.loads(self.ai_ecrs) if self.ai_ecrs else {}
        except ValueError:
            return {}

    @property
    def ai_flags_list(self):
        try:
            return json.loads(self.ai_flags) if self.ai_flags else []
        except ValueError:
            return []

    def apply_ai_result(self, result):
        self.ai_domain = result.get("domain")
        self.ai_related = json.dumps(result.get("related") or [], ensure_ascii=False)
        self.ai_summary = result.get("summary") or ""
        self.ai_ecrs = json.dumps(result.get("ecrs") or {}, ensure_ascii=False)
        self.ai_flags = json.dumps(result.get("flags") or [], ensure_ascii=False)
        self.ai_priority = result.get("priority") or ""
        self.ai_source = result.get("source") or ""
        self.ai_at = datetime.now()


class StageLog(db.Model):
    """案件の段階が変わった履歴（活動状況の日付・最近の動きに使う）"""

    __tablename__ = "stage_logs"
    id = db.Column(db.Integer, primary_key=True)
    task_id = db.Column(db.Integer, db.ForeignKey("tasks.id"), nullable=False, index=True)
    stage = db.Column(db.String(20), nullable=False)
    at = db.Column(db.DateTime, default=datetime.now)
    by = db.Column(db.String(100), default="")
    note = db.Column(db.Text, default="")

    task = db.relationship("Task", back_populates="stage_logs")

    @property
    def stage_name(self):
        return self.task.stage_label(self.stage) if self.task else C.STAGE_NAMES.get(self.stage, self.stage)


class Setting(db.Model):
    """アプリ設定（キー・値）"""

    __tablename__ = "settings"
    key = db.Column(db.String(50), primary_key=True)
    value = db.Column(db.Text, default="")

    @staticmethod
    def get(key, default=""):
        row = db.session.get(Setting, key)
        return row.value if row and row.value is not None else default

    @staticmethod
    def set(key, value):
        row = db.session.get(Setting, key)
        if row is None:
            row = Setting(key=key, value=value)
            db.session.add(row)
        else:
            row.value = value

    @staticmethod
    def as_dict():
        return {s.key: s.value for s in Setting.query.all()}


class Employee(db.Model):
    """社員マスタ（LDAP API の照会結果のキャッシュ、または CSV 取込）"""

    __tablename__ = "employees"
    id = db.Column(db.Integer, primary_key=True)
    emp_id = db.Column(db.String(50), unique=True, nullable=False)
    name = db.Column(db.String(100), nullable=False)
    department = db.Column(db.String(100), default="")
    workplace_id = db.Column(db.Integer, db.ForeignKey("workplaces.id"), nullable=True)
    email = db.Column(db.String(200), default="")
    active = db.Column(db.Boolean, default=True)
    source = db.Column(db.String(20), default="csv")  # csv / ldap / manual
    is_demo = db.Column(db.Boolean, default=False)
    updated_at = db.Column(db.DateTime, default=datetime.now, onupdate=datetime.now)

    workplace = db.relationship("Workplace")


class MonthlySurvey(db.Model):
    """月度ごとの設問セット（voiceloop の surveys に相当）"""

    __tablename__ = "monthly_surveys"
    id = db.Column(db.Integer, primary_key=True)
    period = db.Column(db.String(7), unique=True, nullable=False)  # "2026-11"
    questions_json = db.Column(db.Text, default="[]")
    is_active = db.Column(db.Boolean, default=True)
    created_at = db.Column(db.DateTime, default=datetime.now)
    updated_at = db.Column(db.DateTime, default=datetime.now, onupdate=datetime.now)

    responses = db.relationship("MonthlyResponse", back_populates="survey")

    @property
    def questions(self):
        try:
            data = json.loads(self.questions_json or "[]")
        except ValueError:
            data = []
        return data if isinstance(data, list) else []

    def set_questions(self, questions):
        self.questions_json = json.dumps(questions, ensure_ascii=False)

    @property
    def period_label(self):
        return period_label(self.period)


class MonthlyResponse(db.Model):
    """月次アンケートの回答（〇月度 × 社員ID で1件。1人1回）"""

    __tablename__ = "monthly_responses"
    __table_args__ = (db.UniqueConstraint("period", "emp_id", name="uq_monthly_period_emp"),)
    id = db.Column(db.Integer, primary_key=True)
    period = db.Column(db.String(7), nullable=False)
    emp_id = db.Column(db.String(50), nullable=False)
    name = db.Column(db.String(100), default="")
    workplace_id = db.Column(db.Integer, db.ForeignKey("workplaces.id"), nullable=True)
    respondent_id = db.Column(db.Integer, db.ForeignKey("respondents.id"), nullable=True)
    survey_id = db.Column(db.Integer, db.ForeignKey("monthly_surveys.id"), nullable=True)
    answers_json = db.Column(db.Text, default="{}")  # {設問キー: 回答（選択肢の文言 or 記述）}
    comment = db.Column(db.Text, default="")
    created_at = db.Column(db.DateTime, default=datetime.now)
    updated_at = db.Column(db.DateTime, default=datetime.now, onupdate=datetime.now)
    is_demo = db.Column(db.Boolean, default=False)

    workplace = db.relationship("Workplace")
    survey = db.relationship("MonthlySurvey", back_populates="responses")

    @property
    def answers(self):
        try:
            data = json.loads(self.answers_json or "{}")
        except ValueError:
            data = {}
        return data if isinstance(data, dict) else {}

    def answer(self, key):
        return self.answers.get(key) or ""

    @property
    def hours_shift(self):
        """1シフトあたりの換算時間（推計月間時間 ＝ これ × SHIFTS_PER_MONTH）"""
        return C.MONTHLY_HOURS_VALUE.get(self.answer("hours"), 0.0)

    @property
    def period_label(self):
        return period_label(self.period)


def period_label(period):
    try:
        y, m = period.split("-")
        return f"{int(y)}年{int(m)}月度"
    except (ValueError, AttributeError):
        return period or ""


class FlowMap(db.Model):
    """業務マップ（流れ図）の手修正。絞り込み（職場・領域）ごとに1件（kind は "info" のみ）"""

    __tablename__ = "flow_maps"
    __table_args__ = (db.UniqueConstraint("kind", "scope", name="uq_flow_map_kind_scope"),)
    id = db.Column(db.Integer, primary_key=True)
    kind = db.Column(db.String(20), nullable=False)  # "info"（業務マップ）
    scope = db.Column(db.String(60), nullable=False)  # "wp:0|dom:0" など
    data = db.Column(db.Text, default="{}")  # グラフ JSON（位置・名前・追加・削除・向き）
    updated_by = db.Column(db.String(100), default="")
    updated_at = db.Column(db.DateTime, default=datetime.now, onupdate=datetime.now)


# 既存DBに後から追加した列（SQLite は create_all で列を追加しないため起動時に補う）
_ADDED_COLUMNS = {
    "respondents": {
        "emp_id": "VARCHAR(50)",
    },
    "monthly_responses": {
        "survey_id": "INTEGER",
        "answers_json": "TEXT DEFAULT '{}'",
        "comment": "TEXT DEFAULT ''",
    },
    "tasks": {
        "actors": "TEXT DEFAULT ''",
        "info_from": "TEXT DEFAULT ''",
        "stage": "VARCHAR(20)",
        "stage_at": "DATETIME",
        "progress_note": "TEXT DEFAULT ''",
    },
}

# 列を追加した直後に、既存の行を埋める SQL（(テーブル, 列) → SQL のリスト）
_BACKFILL = {
    ("tasks", "stage"): [
        "UPDATE tasks SET stage = CASE status WHEN '完了' THEN 'done' WHEN '対応中' THEN 'request' ELSE 'start' END WHERE stage IS NULL",
        "UPDATE tasks SET stage_at = COALESCE(completed_at, updated_at, created_at) WHERE stage_at IS NULL",
        "INSERT INTO stage_logs (task_id, stage, at, by, note) SELECT id, stage, stage_at, '', '' FROM tasks WHERE stage <> 'start'",
    ],
}


# なくした機能のテーブル（起動時に片づける）: 教材・勉強会の登録、ログイン許可リスト、メール配信の履歴
_DROPPED_TABLES = ("resources", "allowed_users", "mail_logs")


def ensure_schema():
    """テーブル作成後に、不足している列を ALTER TABLE で追加し、なくした機能のテーブルを落とす。"""
    from sqlalchemy import inspect, text

    inspector = inspect(db.engine)
    changed = False
    added = []
    for table in _DROPPED_TABLES:
        if table in inspector.get_table_names():
            db.session.execute(text(f"DROP TABLE {table}"))
            changed = True
    for table, columns in _ADDED_COLUMNS.items():
        if table not in inspector.get_table_names():
            continue
        existing = {c["name"] for c in inspector.get_columns(table)}
        for name, ddl in columns.items():
            if name not in existing:
                db.session.execute(text(f"ALTER TABLE {table} ADD COLUMN {name} {ddl}"))
                added.append((table, name))
                changed = True
    for key in added:
        for sql in _BACKFILL.get(key, []):
            db.session.execute(text(sql))
    if changed:
        db.session.commit()


# ==============================================================================
# 社員ID照会の控え（ldap_auth の結果を employees に残し、足りない氏名・職場を補う）
# ==============================================================================
ldap_status_text = ldap_auth.status_text  # 管理画面に出す接続状態


def workplace_for(department):
    """所属名から職場マスタを推定する（名前の包含で一致）"""
    department = (department or "").strip()
    if not department:
        return None
    for w in Workplace.query.all():
        if w.name == department or w.name in department or department in w.name:
            return w
    return None


def _upsert(info, source):
    emp = Employee.query.filter_by(emp_id=info["emp_id"]).first()
    if emp is None:
        emp = Employee(emp_id=info["emp_id"], name=info["name"])
        db.session.add(emp)
    emp.name = info["name"] or emp.name
    emp.department = info.get("department") or emp.department
    emp.email = info.get("email") or emp.email
    emp.source = source
    wp = workplace_for(emp.department)
    if wp is not None:
        emp.workplace_id = wp.id
    db.session.commit()
    return emp


def lookup(emp_id):
    """社員IDから社員情報を返す。見つからなければ None。"""
    emp_id = (emp_id or "").strip()
    if not emp_id:
        return None
    info = ldap_auth.lookup_directory(emp_id)
    if info is not None:
        emp = _upsert(info, "ldap")
    else:
        emp = Employee.query.filter_by(emp_id=emp_id).first()
    if emp is None:
        return None
    name, workplace_id = emp.name or "", emp.workplace_id
    if not name or not workplace_id:  # 照会に氏名・職場がないときは、本人が以前に答えた内容で補う
        r_name, r_wp = remembered_person(emp_id)
        name, workplace_id = name or r_name, workplace_id or r_wp
    workplace = db.session.get(Workplace, workplace_id) if workplace_id else None
    return {
        "emp_id": emp.emp_id,
        "name": name,
        "department": emp.department or "",
        "email": emp.email or "",
        "workplace_id": workplace_id,
        "workplace": workplace.name if workplace else "",
        "source": emp.source,
    }


def remembered_person(emp_id):
    """本人が以前に答えた氏名・職場（自動化依頼の登録 → 直近の月次回答 の順）。なければ ("", None)"""
    emp_id = (emp_id or "").strip()
    name, workplace_id = "", None
    if not emp_id:
        return name, workplace_id
    r = Respondent.query.filter_by(emp_id=emp_id).order_by(Respondent.id.desc()).first()
    if r is not None:
        name, workplace_id = r.name or "", r.workplace_id
    if not name or not workplace_id:
        for row in MonthlyResponse.query.filter_by(emp_id=emp_id).order_by(MonthlyResponse.created_at.desc(), MonthlyResponse.id.desc()).all():
            name = name or (row.name or "")
            workplace_id = workplace_id or row.workplace_id
            if name and workplace_id:
                break
    return name, workplace_id


def remember_person(emp_id, name=None, workplace_id=None):
    """本人がアンケートで入力・選択した氏名・職場を社員情報の控えに残す（照会に氏名・職場がない環境向け）。commit は呼び出し側"""
    emp = Employee.query.filter_by(emp_id=(emp_id or "").strip()).first()
    if emp is None:
        return False
    changed = False
    if name and not emp.name:
        emp.name = name
        changed = True
    if workplace_id and emp.workplace_id != workplace_id:
        emp.workplace_id = workplace_id
        changed = True
    return changed


# ==============================================================================
# 集計・KPI・分析不足領域の判定・次アクション提案
# ==============================================================================
def _to_int(value, default=0):
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


# ------------------------------------------------------------
# KPI
# ------------------------------------------------------------
def kpis():
    tasks = Task.query.all()
    respondents = Respondent.query.all()
    target = _to_int(Setting.get("target_respondents", "0"))

    total_minutes = sum(t.monthly_minutes or 0 for t in tasks)
    saved_minutes = sum(t.saved_minutes for t in tasks)
    completed = [t for t in tasks if t.status == "完了"]
    status_counts = Counter(t.status for t in tasks)
    sufficient = [t for t in tasks if t.is_sufficient]
    filled = [t for t in tasks if t.required_filled]
    verified = [t for t in tasks if t.verified]
    self_improvers = [r for r in respondents if r.wants_to_improve]
    automated_or_eliminated = [t for t in completed if t.decision in ("廃止", "自動化")]
    candidates = [t for t in tasks if t.is_candidate]
    domain_rows = domain_analysis()

    return {
        "task_count": len(tasks),
        "respondent_count": len(respondents),
        "target_respondents": target,
        "response_rate": (len(respondents) / target * 100.0) if target else None,
        "total_minutes": total_minutes,
        "total_hours": total_minutes / 60.0,
        "annual_hours": total_minutes * 12 / 60.0,
        "saved_minutes": saved_minutes,
        "saved_annual_hours": saved_minutes * 12 / 60.0,
        "reduction_rate": (saved_minutes / total_minutes * 100.0) if total_minutes else 0.0,
        "status_counts": status_counts,
        "completed_count": len(completed),
        "candidate_count": len(candidates),
        "automated_count": len(automated_or_eliminated),
        "self_improver_count": len(self_improvers),
        "filled_rate": (len(filled) / len(tasks) * 100.0) if tasks else 0.0,
        "verified_rate": (len(verified) / len(tasks) * 100.0) if tasks else 0.0,
        "sufficient_rate": (len(sufficient) / len(tasks) * 100.0) if tasks else 0.0,
        "domain_rows": domain_rows,
        "insufficient_domains": [d for d in domain_rows if d["level"] <= 1],
        "ai_count": len([t for t in tasks if t.ai_source]),
    }


# ------------------------------------------------------------
# 職場別・領域別
# ------------------------------------------------------------
def workplace_stats():
    rows = []
    for wp in Workplace.query.order_by(Workplace.sort, Workplace.id).all():
        tasks = Task.query.filter_by(workplace_id=wp.id).all()
        respondents = Respondent.query.filter_by(workplace_id=wp.id).all()
        rows.append(
            {
                "workplace": wp,
                "respondents": len(respondents),
                "tasks": len(tasks),
                "minutes": sum(t.monthly_minutes or 0 for t in tasks),
                "candidates": len([t for t in tasks if t.is_candidate]),
                "self_improvers": len([r for r in respondents if r.wants_to_improve]),
            }
        )
    return rows


def judge_level(count, fill_rate):
    for level, state, grasp, min_count, min_fill in C.ANALYSIS_LEVELS:
        if count >= min_count and fill_rate >= min_fill:
            return level, state, grasp
    return 0, "未分析", "不明"


def domain_analysis():
    """業務領域ごとの登録状況と分析不足の判定（企画書 柱② 3. 分析不足領域の見える化）"""
    rows = []
    domains = Domain.query.filter_by(parent_id=None).order_by(Domain.sort, Domain.id).all()
    for d in domains:
        child_ids = [c.id for c in d.children]
        ids = [d.id] + child_ids
        tasks = Task.query.filter(Task.domain_id.in_(ids)).all()
        count = len(tasks)
        filled = len([t for t in tasks if t.required_filled])
        verified = len([t for t in tasks if t.verified])
        fill_rate = (filled / count) if count else 0.0
        verified_rate = (verified / count) if count else 0.0
        level, state, grasp = judge_level(count, fill_rate)
        if count and verified_rate < 0.5 and level == 3:
            # 必須項目は埋まっているが担当者確認が進んでいない
            level, state = 2, "追加調査"
        rows.append(
            {
                "domain": d,
                "count": count,
                "candidates": len([t for t in tasks if t.is_candidate]),
                "minutes": sum(t.monthly_minutes or 0 for t in tasks),
                "fill_rate": fill_rate * 100.0,
                "verified_rate": verified_rate * 100.0,
                "level": level,
                "state": state,
                "grasp": grasp,
                "completed": len([t for t in tasks if t.status == "完了"]),
            }
        )
    return rows


# ------------------------------------------------------------
# 案件候補・重複候補・つながり
# ------------------------------------------------------------
def candidate_tasks(limit=None):
    tasks = [t for t in Task.query.all() if t.is_candidate and t.status != "完了"]
    tasks.sort(key=lambda t: t.monthly_minutes or 0, reverse=True)
    return tasks[:limit] if limit else tasks


def elimination_candidates():
    tasks = [t for t in Task.query.all() if t.is_elimination_candidate and t.status != "完了"]
    tasks.sort(key=lambda t: t.monthly_minutes or 0, reverse=True)
    return tasks


_NORMALIZE_RE = re.compile(r"[\s　（）()・,、。.／/\-−―「」『』]")


def normalize_text(value):
    value = (value or "").strip().lower()
    return _NORMALIZE_RE.sub("", value)


def duplicate_output_groups(min_size=2):
    """同じアウトプットを複数の人が作っている業務群（ECRSの結合候補）"""
    groups = defaultdict(list)
    for t in Task.query.all():
        key = normalize_text(t.output)
        if len(key) >= 2:
            groups[key].append(t)
    result = [
        {"output": tasks[0].output, "tasks": tasks, "minutes": sum(x.monthly_minutes or 0 for x in tasks)}
        for tasks in groups.values()
        if len(tasks) >= min_size and len({x.respondent_id for x in tasks}) >= 2
    ]
    result.sort(key=lambda g: g["minutes"], reverse=True)
    return result


def duplicate_input_tasks():
    """「同じ情報を何度も入力する」と回答された業務（入力元別）"""
    rows = defaultdict(list)
    for t in Task.query.all():
        if "同じ情報を何度も入力する" in t.pains_list:
            key = "・".join(t.inputs_list) or "入力元不明"
            rows[key].append(t)
    return sorted(rows.items(), key=lambda kv: -len(kv[1]))


def input_source_stats():
    counter = Counter()
    minutes = Counter()
    for t in Task.query.all():
        for src in t.inputs_list:
            counter[src] += 1
            minutes[src] += t.monthly_minutes or 0
    return [(k, counter[k], minutes[k]) for k in sorted(counter, key=lambda k: -counter[k])]


def flow_rows(domain_id=None):
    """前の業務 → 業務 → 後の業務 の一覧（業務フロー可視化の材料）"""
    q = Task.query
    if domain_id:
        q = q.filter_by(domain_id=domain_id)
    rows = []
    for t in q.order_by(Task.domain_id, Task.name).all():
        if (t.prev_task or "").strip() or (t.next_task or "").strip():
            rows.append(t)
    return rows


def related_tasks_for(task, limit=8):
    """語の重なり・入出力の共通性から関連業務を推定する（ルールベース）"""
    words = _content_words(task.name + " " + (task.output or "") + " " + (task.purpose or ""))
    scored = []
    for other in Task.query.filter(Task.id != task.id).all():
        score = 0
        other_words = _content_words(other.name + " " + (other.output or "") + " " + (other.purpose or ""))
        score += 2 * len(words & other_words)
        if normalize_text(other.output) and normalize_text(other.output) == normalize_text(task.output):
            score += 5
        if task.next_task and normalize_text(task.next_task) in normalize_text(other.name):
            score += 4
        if task.prev_task and normalize_text(task.prev_task) in normalize_text(other.name):
            score += 4
        if other.domain_id and other.domain_id == task.domain_id:
            score += 1
        if set(task.inputs_list) & set(other.inputs_list):
            score += 1
        if score >= 3:
            scored.append((score, other))
    scored.sort(key=lambda x: -x[0])
    return [o for _, o in scored[:limit]]


_WORD_RE = re.compile(r"[一-龥ぁ-んァ-ヶーA-Za-z0-9]{2,}")
_STOPWORDS = {"する", "した", "して", "ため", "こと", "もの", "から", "まで", "ます", "です", "いる", "ある", "など", "作成", "業務", "確認"}


def _content_words(text):
    words = set()
    for w in _WORD_RE.findall(text or ""):
        if w in _STOPWORDS:
            continue
        words.add(w)
        # 漢字の連なりは2文字ずつにも分解して部分一致を拾う
        if len(w) >= 4 and re.fullmatch(r"[一-龥]+", w):
            for i in range(0, len(w) - 1):
                words.add(w[i : i + 2])
    return words


# ------------------------------------------------------------
# DX人財情報
# ------------------------------------------------------------
def skills_summary():
    have = Counter()
    learn = Counter()
    for r in Respondent.query.all():
        for s in r.skills:
            have[s] += 1
        for s in r.learn:
            learn[s] += 1
    return have, learn


def self_improvers():
    rows = []
    for r in Respondent.query.order_by(Respondent.workplace_id, Respondent.name).all():
        if not r.wants_to_improve:
            continue
        target_tasks = [t for t in r.tasks if t.self_improve in C.SELF_IMPROVE_POSITIVE]
        rows.append({"respondent": r, "tasks": target_tasks})
    return rows


def dx_people(workplace_id=None, wish=None):
    """DX人財情報: 自動化依頼の「改善への関わり方」と月次アンケート（改善をどう進めたいか）を人ごとにまとめる

    wish: "positive"（自分で改善したい人）または C1 の選択肢の文言
    """
    people = {}

    def key_for(emp_id, name):
        return ("id", emp_id) if emp_id else ("name", name)

    def blank(emp_id, name, workplace, wp_id):
        return {
            "emp_id": emp_id or "", "name": name, "workplace": workplace, "workplace_id": wp_id, "respondent": None,
            "wish": "", "skills": [], "learn": [], "hours": "", "share": "", "targets": "", "target_tasks": [], "task_count": 0, "monthly": [],
        }

    for r in Respondent.query.all():
        p = blank(r.emp_id, r.name, r.workplace, r.workplace_id)
        p.update(
            respondent=r, wish=r.c1_wish or "", skills=r.skills, learn=r.learn, hours=r.c4_hours or "", share=r.c6_share or "",
            targets=r.c5_targets or "", target_tasks=[t for t in r.tasks if t.self_improve in C.SELF_IMPROVE_POSITIVE], task_count=len(r.tasks),
        )
        people[key_for(r.emp_id, r.name)] = p
    for m in MonthlyResponse.query.order_by(MonthlyResponse.period.desc(), MonthlyResponse.id.desc()).all():
        k = key_for(m.emp_id, m.name)
        p = people.get(k)
        if p is None:
            p = people[k] = blank(m.emp_id, m.name, m.workplace, m.workplace_id)
        p["monthly"].append(m)

    wish_rank = {opt: i for i, opt in enumerate(C.C1_OPTIONS)}
    stance_options = next((q["options"] for q in C.MONTHLY_CORE_QUESTIONS if q["key"] == "stance"), [])
    stance_rank = {opt: i for i, opt in enumerate(stance_options)}
    rows = []
    for p in people.values():
        latest = p["monthly"][0] if p["monthly"] else None
        p["latest_period"] = latest.period if latest else ""
        p["latest_label"] = period_label(latest.period) if latest else ""
        p["stance_latest"] = latest.answer("stance") if latest else ""
        p["self_positive_count"] = sum(1 for m in p["monthly"] if m.answer("stance") in C.MONTHLY_KPI["self_positive"])
        p["months"] = len(p["monthly"])
        p["wish_positive"] = p["wish"] in C.C1_POSITIVE
        r = p["respondent"]
        p["positive"] = (r.wants_to_improve if r is not None else False) or p["stance_latest"] in C.MONTHLY_KPI["self_positive"]
        if workplace_id and p["workplace_id"] != workplace_id:
            continue
        if wish == "positive" and not p["positive"]:
            continue
        if wish and wish != "positive" and p["wish"] != wish:
            continue
        rows.append(p)
    rows.sort(key=lambda p: (not p["positive"], wish_rank.get(p["wish"], 9), stance_rank.get(p["stance_latest"], 9), p["workplace_id"] or 0, p["name"]))
    return rows


# ------------------------------------------------------------
# 次のアクション提案（ルールベース）
# ------------------------------------------------------------
def suggest_actions(limit=10):
    suggestions = []
    tasks = Task.query.all()

    # 1. 分析不足領域 → ヒアリング
    for row in domain_analysis():
        if row["level"] <= 1:
            suggestions.append(
                {
                    "kind": "調査",
                    "title": f"「{row['domain'].name}」領域の業務情報が不足しています",
                    "detail": f"登録 {row['count']} 件・必須項目の充足率 {row['fill_rate']:.0f}%。担当者へのヒアリングで業務の実態を把握してください。",
                    "url": ("bizmap.domain", {"domain_id": row["domain"].id}),
                }
            )

    # 2. 月間時間の大きい案件候補 → 優先案件
    for t in candidate_tasks(limit=3):
        suggestions.append(
            {
                "kind": "案件",
                "title": f"月間 {t.monthly_minutes / 60:.1f} 時間の「{t.name}」を優先案件に",
                "detail": f"改善希望: {t.wish or '－'}／困りごと: {'・'.join(t.pains_list) or '－'}。標準手順（アウトプット確認 → 目的 → ECRS → 自動化）で検討してください。",
                "url": ("cases.detail", {"task_id": t.id}),
            }
        )

    # 3. 目的不明・利用者不明 → 廃止候補
    unknown = [t for t in tasks if (t.purpose_unknown or "分からない" in t.output_users_list) and t.status != "完了"]
    if unknown:
        names = "、".join(t.name for t in unknown[:3])
        suggestions.append(
            {
                "kind": "ECRS 排除",
                "title": f"目的または利用者が不明な業務が {len(unknown)} 件あります",
                "detail": f"{names} など。アウトプットが本当に使われているかを確認し、使われていなければ廃止候補にしてください。",
                "url": ("cases.index", {"filter": "elimination"}),
            }
        )

    # 4. 重複入力 → 共通化
    dup_inputs = duplicate_input_tasks()
    total_dup = sum(len(v) for _, v in dup_inputs)
    if total_dup:
        top_src = dup_inputs[0][0]
        suggestions.append(
            {
                "kind": "ECRS 結合",
                "title": f"「同じ情報を何度も入力する」業務が {total_dup} 件あります",
                "detail": f"入力元「{top_src}」からの入力が最多。発生源で一度だけ入力する仕組みや、入力の共通化を検討してください。",
                "url": ("bizmap.index", {}),
            }
        )

    # 5. 同じアウトプットを複数人が作成 → 統合
    for g in duplicate_output_groups()[:2]:
        suggestions.append(
            {
                "kind": "ECRS 結合",
                "title": f"同じアウトプット「{g['output']}」を {len(g['tasks'])} 人が作成しています",
                "detail": f"合計 月間 {g['minutes'] / 60:.1f} 時間。1つにまとめる、または上流で自動生成する統合候補です。",
                "url": ("bizmap.index", {}),
            }
        )

    # 6. 案件が集中する領域 → プロセス再設計
    for row in domain_analysis():
        if row["candidates"] >= 5:
            suggestions.append(
                {
                    "kind": "業務設計",
                    "title": f"「{row['domain'].name}」領域に改善案件候補が {row['candidates']} 件集中しています",
                    "detail": "個別の自動化ではなく、業務プロセス全体（To-Be）の再設計を検討してください。",
                    "url": ("bizmap.domain", {"domain_id": row["domain"].id}),
                }
            )

    # 7. 学びたい技術 → 教育テーマ
    _, learn = skills_summary()
    if learn:
        skill, n = learn.most_common(1)[0]
        suggestions.append(
            {
                "kind": "人材育成",
                "title": f"教育テーマ候補: {skill}（学びたい人 {n} 人）",
                "detail": "自主改善希望者の業務を題材にした実践型の勉強会を計画してください。",
                "url": ("people.index", {}),
            }
        )

    # 8. 未回答の職場
    quiet = [w["workplace"].name for w in workplace_stats() if w["respondents"] == 0]
    if quiet:
        suggestions.append(
            {
                "kind": "回答促進",
                "title": f"まだ回答のない職場: {'、'.join(quiet)}",
                "detail": "推進メンバーから記入会の案内とフォローをお願いします。",
                "url": ("admin.index", {}),
            }
        )

    return suggestions[:limit]


# ------------------------------------------------------------
# 月次アンケートの集計（〇月度ごと・時系列）
# ------------------------------------------------------------
def monthly_denominator():
    active = Employee.query.filter_by(active=True).count()
    if active:
        return active
    return _to_int(Setting.get("target_respondents", "0"))


def monthly_summary():
    """月度ごとの集計（基本設問の割合。時系列で見る）"""
    responses = MonthlyResponse.query.all()
    periods = sorted({r.period for r in responses})
    denominator = monthly_denominator()
    kpi = C.MONTHLY_KPI
    rows = []
    for p in periods:
        rs = [r for r in responses if r.period == p]
        n = len(rs)

        def rate(key, positive):
            return (len([r for r in rs if r.answer(key) in positive]) / n * 100.0) if n else 0.0

        hours_shift = sum(r.hours_shift for r in rs)
        rows.append(
            {
                "period": p,
                "label": period_label(p),
                "n": n,
                "response_rate": (n / denominator * 100.0) if denominator else None,
                "has_rate": rate("has", kpi["has_positive"]),
                "has_strong_rate": rate("has", kpi["has_strong"]),
                "hours_shift": hours_shift,
                "hours_month": hours_shift * C.SHIFTS_PER_MONTH,
                "decreased_rate": rate("trend", kpi["trend_decreased"]),
                "increased_rate": rate("trend", kpi["trend_increased"]),
                "wants_rate": rate("wants", kpi["wants_positive"]),
                "ask_rate": rate("stance", kpi["ask_positive"]),
                "self_rate": rate("stance", kpi["self_positive"]),
                "new_count": len([r for r in rs if r.answer("new_tasks").strip()]),
                "named_tasks": len([r for r in rs if r.answer("wants_tasks").strip()]),
                "comment_count": len([r for r in rs if (r.comment or "").strip()]),
            }
        )
    return rows


def _responses_by_period():
    responses = MonthlyResponse.query.all()
    periods = sorted({r.period for r in responses})
    return periods, {p: [r for r in responses if r.period == p] for p in periods}


def monthly_question_trends():
    """基本設問（選択式）ごとに、選択肢の割合が月度でどう動いたか"""
    periods, by_period = _responses_by_period()
    trends = []
    for q in C.MONTHLY_CORE_QUESTIONS:
        if q["type"] != "choice":
            continue
        counters = [Counter(r.answer(q["key"]) for r in by_period[p] if r.answer(q["key"])) for p in periods]
        n_list = [sum(c.values()) for c in counters]
        series = []
        for opt in q["options"]:
            counts = [c.get(opt, 0) for c in counters]
            series.append({"option": opt, "counts": counts, "percents": [round(c / n * 100) if n else 0 for c, n in zip(counts, n_list)]})
        trends.append({"q": q, "label": C.MONTHLY_KPI_LABELS.get(q["key"], ""), "periods": periods, "labels": [period_label(p) for p in periods], "n": n_list, "series": series})
    return trends


def monthly_workplace_trends():
    """職場ごとに、繰り返し業務あり率と推計月間時間が月度でどう動いたか"""
    periods, by_period = _responses_by_period()
    kpi = C.MONTHLY_KPI
    rows = []
    for wp in Workplace.query.order_by(Workplace.sort, Workplace.id).all():
        cells = []
        for p in periods:
            sub = [r for r in by_period[p] if r.workplace_id == wp.id]
            n = len(sub)
            cells.append(
                {
                    "n": n,
                    "has_rate": (len([r for r in sub if r.answer("has") in kpi["has_positive"]]) / n * 100.0) if n else None,
                    "hours_month": sum(r.hours_shift for r in sub) * C.SHIFTS_PER_MONTH if n else None,
                }
            )
        if any(c["n"] for c in cells):
            rows.append({"workplace": wp, "cells": cells})
    return {"periods": periods, "labels": [period_label(p) for p in periods], "rows": rows}


def monthly_question_stats(survey, responses):
    """設問ごとの分布（選択）と記入一覧（記述）"""
    stats = []
    for q in survey.questions:
        key = q["key"]
        if q["type"] == "choice":
            counts = Counter(r.answer(key) for r in responses if r.answer(key))
            n = sum(counts.values())
            stats.append(
                {
                    "q": q,
                    "n": n,
                    "counts": [counts.get(o, 0) for o in q.get("options", [])],
                    "percents": [round(counts.get(o, 0) / n * 100) if n else 0 for o in q.get("options", [])],
                }
            )
        else:
            items = [(r.name, r.workplace.name if r.workplace else "", r.answer(key)) for r in responses if r.answer(key).strip()]
            stats.append({"q": q, "n": len(items), "items": items})
    return stats


def dashboard_chart_data():
    """Chart.js 用のデータ"""
    wp = workplace_stats()
    wish_counter = Counter(t.wish or "未回答" for t in Task.query.all())
    pain_counter = Counter()
    for t in Task.query.all():
        for p in t.pains_list:
            if p != "特になし":
                pain_counter[p] += 1
    return {
        "workplace": {
            "labels": [w["workplace"].name for w in wp],
            "tasks": [w["tasks"] for w in wp],
            "hours": [round(w["minutes"] / 60.0, 1) for w in wp],
        },
        "wish": {
            "labels": [w for w in C.WISHES if wish_counter.get(w)] + (["未回答"] if wish_counter.get("未回答") else []),
            "values": [wish_counter[w] for w in C.WISHES if wish_counter.get(w)] + ([wish_counter["未回答"]] if wish_counter.get("未回答") else []),
        },
        "pains": {
            "labels": [p for p, _ in pain_counter.most_common()],
            "values": [n for _, n in pain_counter.most_common()],
        },
    }


# ------------------------------------------------------------
# 活動状況（開始 → 要求定義 → 要件定義 → 作成/伴走 → テスト → 確認 → 完了）
# ------------------------------------------------------------
def board_data(tasks, done_limit=8):
    """活動状況に出す案件を、進行中・保留・完了・受付に分けて並べる"""
    now = datetime.now()
    floor = datetime.min

    def since(t):
        return t.stage_since or floor

    active = sorted([t for t in tasks if t.status == "対応中"], key=lambda t: (t.stage_index, since(t)), reverse=True)
    hold = sorted([t for t in tasks if t.status == "保留"], key=since, reverse=True)
    done = sorted([t for t in tasks if t.status == "完了"], key=lambda t: t.completed_at or since(t), reverse=True)
    waiting = sorted([t for t in tasks if t.status == "受付"], key=lambda t: t.created_at or floor, reverse=True)

    done_this_month = len([t for t in done if t.completed_at and (t.completed_at.year, t.completed_at.month) == (now.year, now.month)])
    saved = sum(t.saved_minutes for t in done)
    return {
        "active": active,
        "hold": hold,
        "done": done,
        "done_shown": done[:done_limit] if done_limit else done,
        "waiting": waiting,
        "waiting_total": len(waiting),
        "done_this_month": done_this_month,
        "saved_annual_hours": saved * 12 / 60.0,
    }


# ==============================================================================
# 業務マップ（1枚の流れ図: グラフ生成、Mermaid 変換、手修正のマージ、つながりの表）
# ==============================================================================
_ESCAPES = [
    ('"', "”"), ("#", "＃"), ("<", "＜"), (">", "＞"), ("&", "＆"), ("`", "'"), ("|", "｜"),
    ("[", "［"), ("]", "］"), ("{", "｛"), ("}", "｝"), ("(", "（"), (")", "）"), ("\\", "＼"),
    ("\r", " "), ("\n", " "), (";", "；"),
]

SHAPES = {
    "rect": ('["', '"]'),
    "round": ('("', '")'),
    "stadium": ('(["', '"])'),
    "cyl": ('[("', '")]'),
    "para": ('[/"', '"/]'),
    "hex": ('{{"', '"}}'),
}

# ノードの種類（cls）ごとの Mermaid の形
CLS_SHAPE = {
    "task": "rect", "done": "rect", "manual": "rect", "source": "round", "note": "round", "system": "cyl",
    "output": "para", "dupout": "para", "wp": "stadium", "ext": "stadium", "person": "stadium", "place": "stadium", "warn": "stadium", "goods": "hex",
}
CLS_LABELS = {
    "task": "業務", "done": "業務（完了）", "source": "情報源（帳票・人・メール）", "system": "システム", "output": "アウトプット", "dupout": "同じアウトプット（統合候補）",
    "wp": "職場", "ext": "他部署・社外", "person": "ヒト", "warn": "要確認", "note": "メモ", "manual": "追加ノード",
}

CLASS_DEFS = [
    "classDef task fill:#e7f1ff,stroke:#0d6efd,color:#052c65",
    "classDef done fill:#d1e7dd,stroke:#198754,color:#0f5132",
    "classDef source fill:#fff3cd,stroke:#ffc107,color:#664d03",
    "classDef system fill:#f8f9fa,stroke:#6c757d,color:#343a40",
    "classDef output fill:#ffffff,stroke:#0d6efd,color:#052c65,stroke-dasharray:4 2",
    "classDef dupout fill:#fff3cd,stroke:#fd7e14,color:#664d03,stroke-width:2px,stroke-dasharray:4 2",
    "classDef person fill:#fde2e4,stroke:#dc3545,color:#842029",
    "classDef wp fill:#e7f1ff,stroke:#0d6efd,color:#052c65",
    "classDef ext fill:#fff3cd,stroke:#fd7e14,color:#664d03",
    "classDef goods fill:#e2f0d9,stroke:#198754,color:#0f5132",
    "classDef place fill:#f1f3f5,stroke:#adb5bd,color:#343a40",
    "classDef warn fill:#f8d7da,stroke:#dc3545,color:#842029",
    "classDef note fill:#ffffff,stroke:#adb5bd,color:#6c757d,stroke-dasharray:2 2",
    "classDef manual fill:#e9d5ff,stroke:#6f42c1,color:#3b0764",
]

ARROWS = {"solid": "-->", "dotted": "-.->", "thick": "==>", "line": "---"}
STYLE_KIND = {"solid": "info", "thick": "goods", "dotted": "seq", "line": "relation"}
KIND_STYLE = {"info": "solid", "goods": "thick", "seq": "dotted", "relation": "line", "manual": "solid"}


def esc(text):
    s = str(text or "").strip()
    for a, b in _ESCAPES:
        s = s.replace(a, b)
    return s


def wrap(text, width=14, max_lines=3):
    s = esc(text)
    if len(s) <= width:
        return s
    lines = [s[i : i + width] for i in range(0, len(s), width)]
    if len(lines) > max_lines:
        lines = lines[:max_lines]
        lines[-1] = lines[-1][: width - 1] + "…"
    return "<br/>".join(lines)


def nid(prefix, text):
    return prefix + "_" + hashlib.md5(normalize_text(text).encode("utf-8")).hexdigest()[:8]


def scope_key(workplace_id=None, domain_id=None):
    return f"wp:{workplace_id or 0}|dom:{domain_id or 0}"


class Diagram:
    """ノード・エッジを蓄積し、Mermaid 定義またはグラフ JSON を出力する"""

    def __init__(self, direction="LR"):
        self.direction = direction
        self.nodes = OrderedDict()  # id -> dict(text, cls, shape, width, group, group_title)
        self.subgraphs = OrderedDict()  # key -> title
        self.edges = OrderedDict()  # (a, b) -> dict(label, style)
        self.clicks = OrderedDict()  # id -> (url, tooltip)

    def node(self, node_id, text, cls="task", shape=None, subgraph=None, width=14):
        if node_id not in self.nodes:
            self.nodes[node_id] = {
                "text": str(text or "").strip(), "cls": cls, "shape": shape or CLS_SHAPE.get(cls, "rect"),
                "width": width, "group": None, "group_title": None,
            }
        if subgraph is not None:
            key, title = subgraph
            self.subgraphs.setdefault(key, title)
            self.nodes[node_id]["group"] = key
            self.nodes[node_id]["group_title"] = title
        return node_id

    def edge(self, a, b, label="", style="solid"):
        if not a or not b or a == b:
            return
        key = (a, b)
        if key in self.edges:
            old = self.edges[key]
            if old["style"] != "solid" and style == "solid":
                self.edges[key] = {"label": label or old["label"], "style": style}
            return
        self.edges[key] = {"label": label, "style": style}

    def click(self, node_id, url, tooltip=""):
        self.clicks[node_id] = (url, tooltip)

    @property
    def empty(self):
        return not self.nodes

    @property
    def node_count(self):
        return len(self.nodes)

    # ---- Mermaid ----
    def _node_line(self, node_id):
        n = self.nodes[node_id]
        open_, close = SHAPES.get(n["shape"], SHAPES["rect"])
        return f"{node_id}{open_}{wrap(n['text'], n.get('width') or 14)}{close}:::{n['cls']}"

    def render(self):
        lines = [f"flowchart {self.direction}"]
        lines += ["  " + c for c in CLASS_DEFS]
        for key, title in self.subgraphs.items():
            members = [i for i, n in self.nodes.items() if n.get("group") == key]
            if not members:
                continue
            lines.append(f'  subgraph {key}["{esc(title)}"]')
            lines.append(f"    direction {self.direction}")
            for i in members:
                lines.append("    " + self._node_line(i))
            lines.append("  end")
        for i, n in self.nodes.items():
            if not n.get("group"):
                lines.append("  " + self._node_line(i))
        for (a, b), e in self.edges.items():
            label, style = e["label"], e["style"]
            if label:
                lab = esc(label)
                if style == "dotted":
                    lines.append(f'  {a} -. "{lab}" .-> {b}')
                elif style == "thick":
                    lines.append(f'  {a} == "{lab}" ==> {b}')
                elif style == "line":
                    lines.append(f'  {a} -- "{lab}" --- {b}')
                else:
                    lines.append(f'  {a} -- "{lab}" --> {b}')
            else:
                lines.append(f"  {a} {ARROWS.get(style, '-->')} {b}")
        for i, (url, tip) in self.clicks.items():
            lines.append(f'  click {i} "{url}" "{esc(tip)}"')
        return "\n".join(lines)

    # ---- グラフ JSON（エディタ・保存用） ----
    def to_graph(self):
        nodes = []
        for i, n in self.nodes.items():
            nodes.append(
                {
                    "id": i, "text": n["text"], "cls": n["cls"], "shape": n["shape"],
                    "group": n.get("group"), "group_title": n.get("group_title"),
                    "url": self.clicks[i][0] if i in self.clicks else None,
                    "source": "auto", "x": None, "y": None, "color": None, "text_edited": False,
                }
            )
        edges = []
        for (a, b), e in self.edges.items():
            edges.append(
                {
                    "id": f"{a}->{b}", "from": a, "to": b, "label": e["label"], "style": e["style"],
                    "kind": STYLE_KIND.get(e["style"], "info"), "source": "auto", "reversed": False, "label_edited": False,
                }
            )
        return {"nodes": nodes, "edges": edges, "suppressed_nodes": [], "suppressed_edges": [], "view": None}


def graph_to_diagram(graph, direction="LR"):
    """グラフ JSON（手修正をマージ済み）から Mermaid 用の Diagram を組み立てる"""
    d = Diagram(direction)
    for n in graph.get("nodes", []):
        cls = n.get("cls") or "manual"
        sub = (n["group"], n.get("group_title") or n["group"]) if n.get("group") else None
        d.node(n["id"], n.get("text") or "", cls, n.get("shape") or CLS_SHAPE.get(cls, "rect"), subgraph=sub)
        if n.get("url"):
            d.click(n["id"], n["url"], "案件詳細を開く")
    for e in graph.get("edges", []):
        a, b = (e["to"], e["from"]) if e.get("reversed") else (e["from"], e["to"])
        style = e.get("style") or KIND_STYLE.get(e.get("kind") or "info", "solid")
        d.edge(a, b, e.get("label") or "", style)
    return d


def merge_graph(auto, saved):
    """自動生成したグラフに、保存された手修正（位置・名前・色・追加・削除・向き・ラベル）を重ねる"""
    result = {
        "nodes": [], "edges": [], "suppressed_nodes": [], "suppressed_edges": [], "view": None,
        "new_node_ids": [], "needs_layout": True,
    }
    if not saved:
        result["nodes"] = [dict(n) for n in auto["nodes"]]
        result["edges"] = [dict(e) for e in auto["edges"]]
        return result

    s_nodes = {n["id"]: n for n in saved.get("nodes", []) if n.get("id")}
    s_edges = {e["id"]: e for e in saved.get("edges", []) if e.get("id")}
    sup_n = set(saved.get("suppressed_nodes") or [])
    sup_e = set(saved.get("suppressed_edges") or [])
    nodes = []
    for n in auto["nodes"]:
        if n["id"] in sup_n:
            continue
        m = dict(n)
        sn = s_nodes.get(n["id"])
        if sn:
            m["x"], m["y"], m["color"] = sn.get("x"), sn.get("y"), sn.get("color")
            if sn.get("text_edited") and sn.get("text"):
                m["text"] = sn["text"]
                m["text_edited"] = True
        else:
            result["new_node_ids"].append(n["id"])
        nodes.append(m)
    auto_ids = {n["id"] for n in nodes}
    for sn in saved.get("nodes", []):
        if sn.get("source") == "manual" and sn.get("id") and sn["id"] not in auto_ids:
            m = dict(sn)
            m.setdefault("cls", "manual")
            m.setdefault("shape", CLS_SHAPE.get(m["cls"], "rect"))
            nodes.append(m)
    node_ids = {n["id"] for n in nodes}

    edges = []
    seen = set()
    for e in auto["edges"]:
        if e["id"] in sup_e or e["from"] not in node_ids or e["to"] not in node_ids:
            continue
        m = dict(e)
        se = s_edges.get(e["id"])
        if se:
            m["reversed"] = bool(se.get("reversed"))
            if se.get("label_edited"):
                m["label"] = se.get("label") or ""
                m["label_edited"] = True
        edges.append(m)
        seen.add(e["id"])
    for se in saved.get("edges", []):
        if se.get("source") == "manual" and se.get("id") not in seen and se.get("from") in node_ids and se.get("to") in node_ids:
            m = dict(se)
            m.setdefault("kind", "manual")
            m.setdefault("style", KIND_STYLE.get(m["kind"], "solid"))
            edges.append(m)
            seen.add(m["id"])

    result.update(
        nodes=nodes, edges=edges, suppressed_nodes=sorted(sup_n), suppressed_edges=sorted(sup_e),
        view=saved.get("view"), needs_layout=False,
    )
    return result


def sanitize_graph(data):
    """クライアントから受け取ったグラフ JSON を保存用に整える（型と上限を確認）"""
    if not isinstance(data, dict):
        raise ValueError("graph must be an object")
    nodes, edges = [], []
    for n in (data.get("nodes") or [])[:2000]:
        if not isinstance(n, dict) or not n.get("id"):
            continue
        nodes.append(
            {
                "id": str(n["id"])[:80], "text": str(n.get("text") or "")[:200], "cls": str(n.get("cls") or "manual")[:20],
                "shape": str(n.get("shape") or "")[:20] or None, "group": n.get("group"), "group_title": n.get("group_title"),
                "url": n.get("url"), "source": "manual" if n.get("source") == "manual" else "auto",
                "x": _num(n.get("x")), "y": _num(n.get("y")), "color": (str(n.get("color"))[:20] if n.get("color") else None),
                "text_edited": bool(n.get("text_edited")),
            }
        )
    for e in (data.get("edges") or [])[:4000]:
        if not isinstance(e, dict) or not e.get("from") or not e.get("to"):
            continue
        edges.append(
            {
                "id": str(e.get("id") or f"{e['from']}->{e['to']}")[:170], "from": str(e["from"])[:80], "to": str(e["to"])[:80],
                "label": str(e.get("label") or "")[:120], "style": str(e.get("style") or "")[:10] or None,
                "kind": str(e.get("kind") or "manual")[:10], "source": "manual" if e.get("source") == "manual" else "auto",
                "reversed": bool(e.get("reversed")), "label_edited": bool(e.get("label_edited")),
            }
        )
    view = data.get("view") if isinstance(data.get("view"), dict) else None
    return {
        "nodes": nodes, "edges": edges,
        "suppressed_nodes": [str(x)[:80] for x in (data.get("suppressed_nodes") or [])][:2000],
        "suppressed_edges": [str(x)[:170] for x in (data.get("suppressed_edges") or [])][:4000],
        "view": {"x": _num(view.get("x")) or 0, "y": _num(view.get("y")) or 0, "zoom": _num(view.get("zoom")) or 1} if view else None,
    }


def _num(v):
    try:
        return float(v) if v is not None else None
    except (TypeError, ValueError):
        return None


# ------------------------------------------------------------
# 手修正の業務データへの書き戻し
# ------------------------------------------------------------
def _task_id_of(node_id):
    """t12 / o12 のようなノードIDから業務IDを取り出す（それ以外は None）"""
    if not node_id or len(node_id) < 2 or node_id[0] not in ("t", "o") or not node_id[1:].isdigit():
        return None
    return int(node_id[1:])


def _contains(text, value):
    nt, nv = normalize_text(text), normalize_text(value)
    return bool(nv) and nv in nt


def append_text(existing, value):
    existing = (existing or "").strip()
    value = (value or "").strip()
    if not value:
        return existing
    if not existing:
        return value
    if _contains(existing, value):
        return existing
    return existing + "、" + value


def remove_text(existing, value):
    existing = (existing or "").strip()
    value = (value or "").strip()
    if not existing or not value:
        return existing
    import re

    parts = [p.strip() for p in re.split(r"[、,，/／\n;；]+", existing) if p.strip()]
    kept = [p for p in parts if not (_contains(p, value) or _contains(value, p))]
    return "、".join(kept)


def writeback_candidates(saved, tasks_by_id):
    """保存された手修正のうち、業務データに反映できるものを候補として列挙する"""
    cands = []
    if not saved:
        return cands
    nodes = {n["id"]: n for n in saved.get("nodes", []) if n.get("id")}

    def task_of(node_id):
        tid = _task_id_of(node_id)
        return tasks_by_id.get(tid) if tid else None

    def text_of(node_id):
        return (nodes.get(node_id) or {}).get("text") or ""

    def cls_of(node_id):
        return (nodes.get(node_id) or {}).get("cls") or ""

    # 1. 業務名の修正
    for n in saved.get("nodes", []):
        if n.get("text_edited") and str(n.get("id", "")).startswith("t"):
            t = task_of(n["id"])
            if t and (n.get("text") or "").strip() and n["text"].strip() != t.name:
                cands.append(
                    {
                        "id": f"rename:{t.id}", "task_id": t.id, "text": f"業務名「{t.name}」を「{n['text'].strip()}」に変更",
                        "ops": [{"task_id": t.id, "field": "name", "action": "set", "value": n["text"].strip()}],
                    }
                )

    # 2. 手で引いた矢印
    for e in saved.get("edges", []):
        if e.get("source") != "manual":
            continue
        a, b = (e["to"], e["from"]) if e.get("reversed") else (e["from"], e["to"])
        ta, tb = task_of(a), task_of(b)
        ca, cb = cls_of(a), cls_of(b)
        label = (e.get("label") or "").strip()
        if ta and tb and ta.id != tb.id and a.startswith("t") and b.startswith("t"):
            cands.append(
                {
                    "id": f"seq:{ta.id}:{tb.id}", "task_id": ta.id,
                    "text": f"「{ta.name}」の後の業務に「{tb.name}」、「{tb.name}」の前の業務に「{ta.name}」を追加",
                    "ops": [
                        {"task_id": ta.id, "field": "next_task", "action": "append", "value": tb.name},
                        {"task_id": tb.id, "field": "prev_task", "action": "append", "value": ta.name},
                    ],
                }
            )
        elif tb and not ta and b.startswith("t"):
            src = text_of(a)
            if not src:
                continue
            if ca == "system" and src in C.SYSTEM_SOURCES:
                cands.append({"id": f"input:{tb.id}:{a}", "task_id": tb.id, "text": f"「{tb.name}」の入力情報に「{src}」を追加",
                              "ops": [{"task_id": tb.id, "field": "inputs", "action": "add_list", "value": src}]})
            elif ca in ("source", "system", "note", "manual", "output"):
                cands.append({"id": f"info_from:{tb.id}:{a}", "task_id": tb.id, "text": f"「{tb.name}」の情報の受け取り元に「{src}」を追加",
                              "ops": [{"task_id": tb.id, "field": "info_from", "action": "append", "value": src}]})
            elif ca in ("wp", "ext", "person", "warn"):
                cands.append({"id": f"actor:{tb.id}:{a}", "task_id": tb.id, "text": f"「{tb.name}」の関わる人・部署に「{src}」を追加",
                              "ops": [{"task_id": tb.id, "field": "actors", "action": "append", "value": src}]})
        elif ta and not tb and (a.startswith("t") or a.startswith("o")):
            dst = text_of(b)
            if not dst:
                continue
            if cb in ("wp", "ext", "person", "warn"):
                ops = [{"task_id": ta.id, "field": "output_users_note", "action": "append", "value": dst}]
                if cb == "ext":
                    ops.append({"task_id": ta.id, "field": "output_users", "action": "add_list", "value": "他部署"})
                cands.append({"id": f"user:{ta.id}:{b}", "task_id": ta.id, "text": f"「{ta.name}」のアウトプットの利用者に「{dst}」を追加", "ops": ops})
            elif cb in ("note", "manual"):
                cands.append({"id": f"next:{ta.id}:{b}", "task_id": ta.id, "text": f"「{ta.name}」の後の業務に「{dst}」を追加",
                              "ops": [{"task_id": ta.id, "field": "next_task", "action": "append", "value": dst}]})

    # 3. 消した「前後の業務」の矢印
    for eid in saved.get("suppressed_edges", []):
        if "->" not in str(eid):
            continue
        a, b = str(eid).split("->", 1)
        ta, tb = task_of(a), task_of(b)
        if ta and tb and a.startswith("t") and b.startswith("t"):
            if _contains(ta.next_task, tb.name) or _contains(tb.prev_task, ta.name):
                cands.append(
                    {
                        "id": f"unseq:{ta.id}:{tb.id}", "task_id": ta.id,
                        "text": f"「{ta.name}」の後の業務から「{tb.name}」、「{tb.name}」の前の業務から「{ta.name}」を外す",
                        "ops": [
                            {"task_id": ta.id, "field": "next_task", "action": "remove", "value": tb.name},
                            {"task_id": tb.id, "field": "prev_task", "action": "remove", "value": ta.name},
                        ],
                    }
                )

    # すでに反映済み（反映しても何も変わらない）候補は除く
    def changes_something(c):
        for op in c["ops"]:
            t = tasks_by_id.get(op["task_id"])
            if t is None or op["field"] not in WRITEBACK_FIELDS:
                continue
            current = getattr(t, op["field"]) or ""
            if _new_value(current, op) != current:
                return True
        return False

    return [c for c in cands if changes_something(c)]


WRITEBACK_FIELDS = {"name", "next_task", "prev_task", "inputs", "info_from", "actors", "output_users_note", "output_users"}


def _new_value(current, op):
    action = op["action"]
    if action == "set":
        return op["value"]
    if action == "append":
        return append_text(current, op["value"])
    if action == "remove":
        return remove_text(current, op["value"])
    if action == "add_list":
        items = C.split_list(current)
        if op["value"] not in items:
            items.append(op["value"])
        return C.join_list(items)
    return current


def apply_writeback(candidates, selected_ids, tasks_by_id):
    """選ばれた候補を業務データに反映する。変更した件数を返す（commit は呼び出し側）"""
    applied = 0
    for c in candidates:
        if c["id"] not in selected_ids:
            continue
        for op in c["ops"]:
            t = tasks_by_id.get(op["task_id"])
            field = op["field"]
            if t is None or field not in WRITEBACK_FIELDS:
                continue
            current = getattr(t, field) or ""
            new = _new_value(current, op)
            if new != current:
                setattr(t, field, new)
                applied += 1
    return applied


# ------------------------------------------------------------
# 共通
# ------------------------------------------------------------
def _wp_name(t):
    return t.workplace.name if t.workplace else "職場未設定"


def _wp_key(t):
    return f"wp{t.workplace_id or 0}"


def _task_node(d, t, width=14):
    tid = f"t{t.id}"
    d.node(tid, t.name, "done" if t.status == "完了" else "task", "rect", subgraph=(_wp_key(t), _wp_name(t)), width=width)
    d.click(tid, url_for("cases.detail", task_id=t.id), "案件詳細を開く")
    return tid


def _match_tasks(text, tasks, exclude=None, limit=3):
    """前の業務・後の業務の記述に一致する登録業務を探す"""
    nt = normalize_text(text)
    if len(nt) < 2:
        return []
    hits = []
    for t in tasks:
        if t is exclude:
            continue
        nn = normalize_text(t.name)
        if len(nn) >= 3 and (nn in nt or (len(nt) >= 4 and nt in nn)):
            hits.append((len(nn), t))
    hits.sort(key=lambda x: -x[0])
    return [t for _, t in hits[:limit]]


def _user_node(d, u, t):
    wp = _wp_name(t)
    if u == "職場内":
        return d.node(nid("p", wp + "職場内"), f"{wp}（職場内）", "wp")
    if u == "上司":
        return d.node(nid("p", wp + "上司"), f"{wp} 上司", "wp")
    if u == "他部署":
        if (t.output_users_note or "").strip():
            return d.node(nid("p", t.output_users_note), t.output_users_note, "ext")
        return d.node("p_other", "他部署", "ext")
    if u == "顧客・社外":
        return d.node("p_customer", "顧客・社外", "ext")
    if u == "分からない":
        return d.node("p_unknown", "利用者不明", "warn")
    return None


def _add_info_edges(d, t, tid, dup_outputs=None):
    """情報源 → 業務 → アウトプット → 利用者 のエッジを追加し、アウトプットのノードIDを返す"""
    inputs = t.inputs_list
    specific = (t.info_from or "").strip()
    used_specific = False
    for src in inputs:
        if src in C.SYSTEM_SOURCES:
            d.edge(d.node(nid("s", src), src, "system"), tid)
        elif specific:
            d.edge(d.node(nid("s", specific), specific, "source"), tid, "" if src == "その他" else src)
            used_specific = True
        else:
            d.edge(d.node(nid("s", src), src, "source"), tid)
    if specific and not used_specific and not inputs:
        d.edge(d.node(nid("s", specific), specific, "source"), tid)
    target = tid
    if (t.output or "").strip():
        oid = f"o{t.id}"
        d.node(oid, t.output, "dupout" if dup_outputs and normalize_text(t.output) in dup_outputs else "output")
        d.edge(tid, oid)
        target = oid
    for u in t.output_users_list:
        if u == "自分":
            continue
        uid = _user_node(d, u, t)
        if uid:
            d.edge(target, uid)
    return target


# ------------------------------------------------------------
# 情報の流れ
# ------------------------------------------------------------
def build_info_flow(tasks):
    d = Diagram("LR")
    dups = {normalize_text(g["output"]) for g in duplicate_output_groups()}  # 同じアウトプットを複数人が作っている（統合候補）
    for t in tasks:
        tid = _task_node(d, t)
        _add_info_edges(d, t, tid, dups)
    for t in tasks:
        tid = f"t{t.id}"
        for other in _match_tasks(t.next_task, tasks, exclude=t):
            d.edge(tid, f"t{other.id}", "次の業務", "dotted")
        for other in _match_tasks(t.prev_task, tasks, exclude=t):
            d.edge(f"t{other.id}", tid, "前の業務", "dotted")
        no = normalize_text(t.output)
        if len(no) < 3:
            continue
        for other in tasks:
            if other is t:
                continue
            for field in (other.info_from, other.prev_task, other.inputs_note):
                nf = normalize_text(field)
                if len(nf) >= 3 and (no in nf or nf in no):
                    d.edge(f"o{t.id}", f"t{other.id}", "", "dotted")
                    break
    return d


# ------------------------------------------------------------
# 図に添える表: 職場・部署の間の受け渡し、システム・情報源ごとの業務
# ------------------------------------------------------------
def handoff_rows(tasks, workplace_names):
    """職場・部署の間で情報をどう受け渡しているか（件数）。関与は「関わる人・部署」に書かれた相手"""
    agg = defaultdict(lambda: {"info": 0, "involve": 0})

    def is_internal(name):
        return any(w and w in name for w in workplace_names)

    for t in tasks:
        wp = _wp_name(t)
        inputs = t.inputs_list
        if (t.info_from or "").strip():
            agg[(t.info_from.strip(), wp)]["info"] += 1
        elif "人から聞く" in inputs:
            agg[("口頭での聞き取り", wp)]["info"] += 1
        for u in t.output_users_list:
            if u in ("自分", "職場内"):
                continue
            if u == "上司":
                r = f"{wp} 上司"
            elif u == "他部署":
                r = (t.output_users_note or "").strip() or "他部署"
            elif u == "顧客・社外":
                r = "顧客・社外"
            elif u == "分からない":
                r = "利用者不明"
            else:
                continue
            agg[(wp, r)]["info"] += 1
        for a_ in t.actors_list:
            if (a_, wp) not in agg and (wp, a_) not in agg:
                agg[(wp, a_)]["involve"] += 1
    rows = [
        {"src": a_, "dst": b_, "info": c["info"], "involve": c["involve"], "internal": is_internal(a_) and is_internal(b_)}
        for (a_, b_), c in agg.items()
    ]
    rows.sort(key=lambda r: (-r["info"], -r["involve"], r["src"], r["dst"]))
    return rows


def source_rows(tasks):
    """システム・情報源ごとに、それを使っている業務（システムと業務のつながり）"""
    by = OrderedDict()
    for t in tasks:
        specific = (t.info_from or "").strip()
        names = []
        for src in t.inputs_list:
            if src in C.SYSTEM_SOURCES:
                names.append((src, "system"))
            elif specific:
                names.append((specific, "source"))
            else:
                names.append((src, "source"))
        if specific and not t.inputs_list:
            names.append((specific, "source"))
        for name, cls in dict.fromkeys(names):
            row = by.setdefault(name, {"name": name, "cls": cls, "tasks": [], "minutes": 0.0})
            row["tasks"].append(t)
            row["minutes"] += t.monthly_minutes or 0
    rows = list(by.values())
    rows.sort(key=lambda r: (r["cls"] != "system", -len(r["tasks"]), r["name"]))
    return rows


# ------------------------------------------------------------
# 1業務の流れ（案件詳細用）
# ------------------------------------------------------------
def build_task_flow(t):
    d = Diagram("LR")
    tid = f"t{t.id}"
    d.node(tid, t.name, "done" if t.status == "完了" else "task", "rect", width=16)
    _add_info_edges(d, t, tid)
    for a in t.actors_list:
        d.edge(d.node(nid("a", a), a, "person"), tid, "関与", "line")
    if (t.prev_task or "").strip():
        d.edge(d.node("prev_task", t.prev_task, "note"), tid, "前の業務", "dotted")
    if (t.next_task or "").strip():
        d.edge(tid, d.node("next_task", t.next_task, "note"), "次の業務", "dotted")
    return d


# ==============================================================================
# AI連携（Claude API。なければルールベース）
# ==============================================================================
log = logging.getLogger(__name__)

MODEL = "claude-opus-5-5"  # config.py の ANTHROPIC_MODEL が空のときの既定

SYSTEM_PROMPT = """あなたは製造部の「PC定型業務ゼロ化プロジェクト」を支援する業務分析アシスタントです。
現場から登録された1件のPC業務を分析し、業務全体像の把握と改善検討に使う情報をJSONで返します。

プロジェクトの原則:
- 受けた案件をそのまま自動化しない。まずアウトプットと「なぜやるのか」を確認し、
  ECRS（排除・結合・入替・簡素化）の観点でムリ・ムダ・ムラを取り除いて業務プロセスを最適化してから自動化する。
- 目的が不明、アウトプットの利用者が分からない業務は廃止候補として扱う。
- 同じ情報を何度も入力している業務、同じアウトプットを複数人が作っている業務は結合候補として扱う。
- AIの推定は担当者が確認して確定する。断定せず、確認すべき点も示す。

出力は次のキーを持つJSONオブジェクトのみ（前後に説明文を付けない）:
{
  "domain": "<与えられた業務領域リストの中から最も近い1つ>",
  "related": ["<候補リストにある関連業務名>", ...],   // 最大5件。候補リストにない名前は書かない
  "summary": "<この業務の要点と業務全体の中での位置づけを2文以内で>",
  "ecrs": {
    "eliminate": "<排除の観点で検討すべきこと。該当なしなら null>",
    "combine": "<結合の観点。該当なしなら null>",
    "rearrange": "<入替（順序・担当・場所・発生源での入力）の観点。該当なしなら null>",
    "simplify": "<簡素化の観点。該当なしなら null>"
  },
  "automation": "<最適化後に残る作業を自動化するなら、どのツール（Python/Excel VBA/RPA/BI/生成AI等）でどう自動化するか。該当なしなら null>",
  "flags": ["<次の中から該当するもの: 目的不明, 廃止候補, 統合候補, 重複入力候補, 自動化候補, 判断業務, 要ヒアリング>"],
  "priority": "<高 | 中 | 低>",
  "questions": ["<担当者に確認すべき質問（最大3件）>"]
}
日本語で、簡潔に書いてください。"""


def _api_key():
    return str(conf.get("ANTHROPIC_API_KEY", "") or "").strip()


def _model():
    return str(conf.get("ANTHROPIC_MODEL", "") or "").strip() or MODEL


def _sdk_available():
    try:
        import anthropic  # noqa: F401

        return True
    except ImportError:
        return False


def ai_enabled():
    return bool(_api_key()) and _sdk_available()


def ai_status_text():
    if not _sdk_available():
        return "anthropic SDK が未インストール（pip install anthropic）"
    if not _api_key():
        return "config.py の ANTHROPIC_API_KEY が未設定（ルールベースで動作中）"
    return f"Claude API 連携 有効（モデル: {_model()}）"


# ------------------------------------------------------------
# 共通: Claude 呼び出し
# ------------------------------------------------------------
def _call_claude(system, user_text, max_tokens=4000):
    """Claude API を呼び、テキスト本文を返す。拒否（refusal）や失敗時は None。"""
    import anthropic

    client = anthropic.Anthropic(timeout=90.0, max_retries=1)
    base_kwargs = dict(
        model=_model(),
        max_tokens=max_tokens,
        system=system,
        messages=[{"role": "user", "content": user_text}],
    )
    try:
        # 分類用途なので effort は low。安全分類による拒否時はサーバー側フォールバックに任せる。
        response = client.messages.create(
            **base_kwargs,
            extra_body={"output_config": {"effort": "low"}, "fallbacks": "default"},
            extra_headers={"anthropic-beta": "server-side-fallback-2026-07-01"},
        )
    except anthropic.BadRequestError as e:
        log.warning("Claude API: 拡張パラメータが拒否されたため基本形で再試行します: %s", e)
        response = client.messages.create(**base_kwargs)
    except anthropic.AuthenticationError:
        log.error("Claude API: APIキーが無効です")
        return None
    except anthropic.RateLimitError:
        log.warning("Claude API: レート制限に達しました")
        return None
    except anthropic.APIStatusError as e:
        log.warning("Claude API: ステータスエラー %s", e.status_code)
        return None
    except anthropic.APIConnectionError:
        log.warning("Claude API: 接続できません")
        return None

    if getattr(response, "stop_reason", None) == "refusal":
        log.warning("Claude API: 安全上の理由で応答が拒否されました")
        return None
    return "".join(getattr(b, "text", "") for b in response.content if getattr(b, "type", "") == "text")


def _extract_json(text):
    if not text:
        return None
    text = text.strip()
    # ```json ... ``` を剥がす
    fence = re.search(r"```(?:json)?\s*(.*?)```", text, re.S)
    if fence:
        text = fence.group(1).strip()
    start = text.find("{")
    end = text.rfind("}")
    if start == -1 or end == -1:
        return None
    try:
        return json.loads(text[start : end + 1])
    except ValueError:
        return None


# ------------------------------------------------------------
# 業務（案件）の分析
# ------------------------------------------------------------
def task_payload(task):
    return {
        "業務名": task.name,
        "業務領域（回答者の選択）": task.domain.name if task.domain else "未選択",
        "目的": task.purpose or "",
        "アウトプット": task.output or "",
        "アウトプットの利用者": task.output_users_list + ([task.output_users_note] if task.output_users_note else []),
        "入力情報": task.inputs_list + ([task.inputs_note] if task.inputs_note else []),
        "前の業務": task.prev_task or "",
        "後の業務": task.next_task or "",
        "関わる人・部署": task.actors_list,
        "情報の受け取り元": task.info_from or "",
        "頻度": task.frequency or "",
        "1回あたり所要時間": task.duration or "",
        "担当人数": task.headcount or 1,
        "月間時間（分）": task.monthly_minutes or 0,
        "使用ツール": task.tools_list,
        "作業の性質": task.nature_list,
        "困りごと": task.pains_list + ([task.pains_note] if task.pains_note else []),
        "改善の希望": task.wish or "",
        "自分で改善": task.self_improve or "",
        "補足": task.note or "",
    }


def analyze_task(task, domains, candidate_tasks):
    """1業務を分析して dict を返す（必ず何らかの結果を返す）"""
    result = None
    if ai_enabled():
        try:
            result = _claude_analyze(task, domains, candidate_tasks)
        except Exception as e:  # 想定外の失敗はルールベースで補う
            log.exception("Claude 分析に失敗しました: %s", e)
            result = None
    if result is None:
        result = rule_analyze(task, domains, candidate_tasks)
        result["source"] = "rules"
    else:
        result["source"] = "claude"
    return result


def _claude_analyze(task, domains, candidate_tasks):
    domain_names = [d.name for d in domains]
    candidates = [
        {"業務名": t.name, "領域": t.domain.name if t.domain else "", "アウトプット": t.output or "", "入力情報": t.inputs_list}
        for t in candidate_tasks[:40]
    ]
    user_text = (
        "## 分析対象の業務\n"
        + json.dumps(task_payload(task), ensure_ascii=False, indent=1)
        + "\n\n## 業務領域リスト（domain はこの中から選ぶ）\n"
        + json.dumps(domain_names, ensure_ascii=False)
        + "\n\n## 関連候補の業務（related はこの中から選ぶ）\n"
        + json.dumps(candidates, ensure_ascii=False, indent=1)
    )
    text = _call_claude(SYSTEM_PROMPT, user_text)
    data = _extract_json(text)
    if not data:
        return None
    related = [r for r in (data.get("related") or []) if isinstance(r, str)]
    candidate_names = {t.name for t in candidate_tasks}
    related = [r for r in related if r in candidate_names][:5]
    ecrs = data.get("ecrs") or {}
    return {
        "domain": data.get("domain") if data.get("domain") in domain_names else None,
        "related": related,
        "summary": data.get("summary") or "",
        "ecrs": {
            "eliminate": ecrs.get("eliminate"),
            "combine": ecrs.get("combine"),
            "rearrange": ecrs.get("rearrange"),
            "simplify": ecrs.get("simplify"),
            "automation": data.get("automation"),
            "questions": [q for q in (data.get("questions") or []) if isinstance(q, str)][:3],
        },
        "flags": [f for f in (data.get("flags") or []) if f in C.AI_FLAGS],
        "priority": data.get("priority") if data.get("priority") in ("高", "中", "低") else "中",
    }


# ------------------------------------------------------------
# ルールベース（APIキーなしでも同じ形式で動く）
# ------------------------------------------------------------
DOMAIN_KEYWORDS = {
    "生産計画・管理": ["生産計画", "計画", "日程", "納期", "手配", "受注", "出荷", "在庫", "所要量", "指示"],
    "製造実行・管理": ["日報", "実績", "製造", "進捗", "稼働", "出来高", "ライン", "工程", "段取り"],
    "品質管理": ["品質", "検査", "不良", "不適合", "クレーム", "是正", "測定", "規格", "合否"],
    "設備管理": ["設備", "保全", "点検", "故障", "修理", "部品", "メンテ", "予防"],
    "人員・労務管理": ["人員", "勤怠", "シフト", "配置", "残業", "有給", "教育", "資格", "出勤"],
    "原価・生産性管理": ["原価", "コスト", "生産性", "工数", "能率", "予算", "経費", "歩留"],
    "安全・環境管理": ["安全", "ヒヤリ", "災害", "環境", "廃棄", "KY", "5S", "パトロール"],
    "改善・技術開発": ["改善", "提案", "技術", "開発", "試作", "標準", "手順書"],
}


def rule_analyze(task, domains, candidate_tasks):
    text = " ".join([task.name or "", task.purpose or "", task.output or "", task.note or ""])
    # 領域推定
    best, best_score = None, 0
    domain_names = [d.name for d in domains]
    for name, words in DOMAIN_KEYWORDS.items():
        if name not in domain_names:
            continue
        score = sum(text.count(w) for w in words)
        if score > best_score:
            best, best_score = name, score
    if best is None and task.domain:
        best = task.domain.name

    flags = []
    ecrs = {"eliminate": None, "combine": None, "rearrange": None, "simplify": None, "automation": None, "questions": []}
    pains = task.pains_list
    nature = task.nature_list
    inputs = task.inputs_list
    users = task.output_users_list

    if task.purpose_unknown or "何のためか分からない" in pains:
        flags += ["目的不明", "廃止候補"]
        ecrs["eliminate"] = "目的が不明です。アウトプットが実際に使われているかを利用者に確認し、使われていなければ廃止を検討してください。"
        ecrs["questions"].append("このアウトプットは誰が、どの判断に使っていますか？")
    if "分からない" in users:
        if "廃止候補" not in flags:
            flags.append("廃止候補")
        ecrs["questions"].append("アウトプットの届け先と用途を確認してください。")
    if "同じ情報を何度も入力する" in pains:
        flags += ["重複入力候補", "統合候補"]
        ecrs["combine"] = "同じ情報を複数箇所に入力しています。入力を1回にまとめる、または上流のデータから自動で引き当てることを検討してください。"
    if "紙の帳票" in inputs or "人から聞く" in inputs:
        ecrs["rearrange"] = "紙や口頭で受け取った情報を転記しています。発生源（現場・システム）で直接入力する形に順序・担当を変えられないか検討してください。"
    if "手順が人によって違う" in pains:
        ecrs["simplify"] = "手順が人によって違います。手順を標準化し、入力項目や確認工程を最小限にしてください。"
    elif "確認・チェック" in nature and "ミスが起きやすい" in pains:
        ecrs["simplify"] = "確認工程でミスが起きやすい状態です。チェック項目を絞り、システムによる自動チェックに置き換えられないか検討してください。"
    if "判断を伴う" in nature:
        flags.append("判断業務")
    if ("転記・コピー" in nature or "集計・グラフ作成" in nature or "データ取得" in nature) and task.wish != "このままでよい":
        flags.append("自動化候補")
        tool = "Python または Power Query"
        if "Excel" in task.tools_list and "VBA・マクロ" not in task.tools_list:
            tool = "Excel VBA / Power Query"
        if "社内システム" in task.tools_list:
            tool = "RPA またはシステムのデータ出力機能"
        ecrs["automation"] = f"最適化後に残る転記・集計は {tool} での自動化が候補です。"
    if not task.required_filled:
        flags.append("要ヒアリング")
        ecrs["questions"].append("必須項目（目的・アウトプット・利用者・入力元・頻度・時間）を担当者と確認してください。")

    related = [t.name for t in candidate_tasks[:5]]
    minutes = task.monthly_minutes or 0
    if minutes >= 600 or "廃止候補" in flags:
        priority = "高"
    elif minutes >= 120 or task.is_candidate:
        priority = "中"
    else:
        priority = "低"

    summary = f"「{task.name}」は{best or '領域未定'}に属する{task.frequency or '頻度不明'}の業務で、月間約 {minutes / 60:.1f} 時間。"
    if flags:
        summary += "検討ポイント: " + "・".join(dict.fromkeys(flags)) + "。"
    return {
        "domain": best,
        "related": related,
        "summary": summary,
        "ecrs": ecrs,
        "flags": list(dict.fromkeys(flags)),
        "priority": priority,
    }


# ------------------------------------------------------------
# 次のアクション提案（AI）
# ------------------------------------------------------------
ACTION_SYSTEM_PROMPT = """あなたは製造部の「PC定型業務ゼロ化プロジェクト」のデジタル改善チームを支援するアナリストです。
蓄積された業務登録データの集計を読み、次に取るべきアクションを提案してください。

観点:
1. 業務情報が不足している領域（登録が少ない・充足率が低い）への追加調査・ヒアリング
2. 同じ情報の重複入力や同じアウトプットの重複作成 → 共通化・統合（ECRSの結合）
3. 目的不明・利用者不明の業務 → 廃止候補の確認（ECRSの排除）
4. 案件が集中している業務領域 → 個別改善ではなく業務プロセス全体の再設計
5. 自主改善希望者とスキルの状況 → 教育テーマの提案

出力は次の形式のJSON配列のみ（最大6件、重要な順）:
[{"kind": "調査|ECRS 排除|ECRS 結合|業務設計|案件|人材育成|回答促進", "title": "<30字以内>", "detail": "<80字以内。根拠の数字を含める>"}]
日本語で簡潔に。"""


def suggest_actions_ai(context):
    """集計データ（dict）を渡し、AIの提案リストを返す。失敗時は None。"""
    if not ai_enabled():
        return None
    try:
        text = _call_claude(ACTION_SYSTEM_PROMPT, json.dumps(context, ensure_ascii=False, indent=1), max_tokens=3000)
    except Exception as e:
        log.exception("AI提案の生成に失敗: %s", e)
        return None
    if not text:
        return None
    text = text.strip()
    fence = re.search(r"```(?:json)?\s*(.*?)```", text, re.S)
    if fence:
        text = fence.group(1).strip()
    start, end = text.find("["), text.rfind("]")
    if start == -1 or end == -1:
        return None
    try:
        items = json.loads(text[start : end + 1])
    except ValueError:
        return None
    result = []
    for item in items:
        if not isinstance(item, dict) or not item.get("title"):
            continue
        result.append(
            {
                "kind": str(item.get("kind") or "提案"),
                "title": str(item.get("title")),
                "detail": str(item.get("detail") or ""),
                "generated_at": datetime.now().strftime("%Y/%m/%d %H:%M"),
            }
        )
    return result[:6]


# ==============================================================================
# 初期データ・デモデータ・既存データの移行
# ==============================================================================
DEFAULT_WORKPLACES = ["製造1課", "製造2課", "組立課", "検査・品質保証課", "生産管理課", "設備保全課"]

DEFAULT_DOMAINS = [
    ("生産計画・管理", "生産計画、手配、納期、在庫、出荷に関する業務"),
    ("製造実行・管理", "日報、実績、進捗、稼働、工程管理に関する業務"),
    ("品質管理", "検査、不良、クレーム、是正、測定に関する業務"),
    ("設備管理", "設備の点検、保全、故障対応、部品に関する業務"),
    ("人員・労務管理", "勤怠、シフト、人員配置、教育、資格に関する業務"),
    ("原価・生産性管理", "原価、工数、生産性指標、予算に関する業務"),
    ("安全・環境管理", "安全パトロール、ヒヤリハット、環境、5Sに関する業務"),
    ("改善・技術開発", "改善提案、標準化、手順書、技術開発に関する業務"),
]

DEFAULT_SETTINGS = {
    "project_name": "PC定型業務ゼロ化プロジェクト",
    "catchphrase": "なくすのは作業。増やすのは人の可能性。",
    "sub_message": "業務をなくす。自ら改善する。みんなで進化する。",
    "target_respondents": "50",
    "deadline": "3年後（1年目の棚卸し結果を踏まえて確定）",
    "ai_suggestions": "[]",
    "board_show_names": "1",
}

# 旧バージョンの既定値（起動時に新しい既定値へ置き換える）
RENAMED_DEFAULTS = {
    "project_name": ("製造部 繰り返しPC業務ゼロ化ポータル", "製造部 PC定型業務ゼロ化ポータル"),
}

# なくした機能の設定キー（メール配信・案内文・旧掲示板メッセージ）。起動時に消す
REMOVED_SETTINGS = ("survey_deadline", "contact", "portal_base_url", "monthly_deadline", "mail_subject", "mail_body", "board_message")

def seed_masters():
    """マスタが空なら初期値を投入する（起動時に実行）"""
    changed = False
    if Workplace.query.count() == 0:
        for i, name in enumerate(DEFAULT_WORKPLACES):
            db.session.add(Workplace(name=name, sort=i))
        changed = True
    if Domain.query.count() == 0:
        for i, (name, desc) in enumerate(DEFAULT_DOMAINS):
            db.session.add(Domain(name=name, description=desc, sort=i))
        changed = True
    for key, value in DEFAULT_SETTINGS.items():
        if db.session.get(Setting, key) is None:
            db.session.add(Setting(key=key, value=value))
            changed = True
    # 名称を「PC定型業務ゼロ化」に改めた。旧既定値のままの設定だけを置き換える（手で変えた値はそのまま）
    for key, old_values in RENAMED_DEFAULTS.items():
        row = db.session.get(Setting, key)
        if row is not None and row.value in old_values:
            row.value = DEFAULT_SETTINGS[key]
            changed = True
    for key in REMOVED_SETTINGS:
        row = db.session.get(Setting, key)
        if row is not None:
            db.session.delete(row)
            changed = True
    changed = _migrate_monthly_questions() or changed
    for m in FlowMap.query.filter(FlowMap.kind != "info").all():  # 以前の「ヒトの流れ」「モノの流れ」の手修正
        db.session.delete(m)
        changed = True
    if changed:
        db.session.commit()


def _stance_from_old(ask, self_answer):
    """旧2問（頼みたいか／自分で取り組みたいか）の回答を、新1問「どう進めたいか」に読み替える"""
    if self_answer in ("取り組みたい", "教えてもらえれば取り組みたい"):
        return "自分で取り組みたい"
    if ask in ("頼みたい", "相談したい"):
        return "改善チームに依頼したい"
    return ""


# 「どのように進めたいか」の以前の選択肢（4択）を今の2択に読み替える
_STANCE_OLD_TO_NEW = {"教えてもらいながら自分で取り組みたい": "自分で取り組みたい", "改善チームに頼みたい": "改善チームに依頼したい", "今は特にない": ""}


# 週あたりで聞いていた時間を、1シフトあたりの選択肢に読み替える（週5シフト換算）
_HOURS_WEEK_TO_SHIFT = {"1時間未満": "30分未満", "1〜3時間": "30分未満", "3〜5時間": "30分〜1時間", "5〜10時間": "1〜2時間", "10時間以上": "2〜4時間"}
# 基本設問の以前の文面。これと同じなら新しい文面に置き換える（手で変えた文面はそのまま）
_PREVIOUS_TEXTS = {
    "has": {"今月、PC定型業務（繰り返しの手作業）はありますか"},
    "hours": {"1週間あたり、PC定型業務にかかっている時間は"},
    "wants": {"なくしたい・自動化したい業務はありますか"},
    "stance": {"なくしたい業務の改善は、どう進めたいですか"},
    "new_tasks": {"今月、新たに発生したPC定型業務"},
}


def _migrate_monthly_questions():
    """月次の設問セットと回答を、現在の基本設問に合わせる（起動時。移行済みなら何もしない）

    - 旧2問（頼みたいか／自分でやりたいか）→ 新1問「どう進めたいか」
    - 時間の設問を週あたりから1シフトあたりへ（選択肢を読み替え）
    - 「先月と比べて」の「初回なので分からない」は未回答扱い
    - 基本設問の選択肢・必須・条件（depends）を定義どおりに、以前の既定の文面は新しい文面に
    """
    core = {q["key"]: q for q in C.MONTHLY_CORE_QUESTIONS}
    stance_q = core["stance"]
    changed = False
    for survey in MonthlySurvey.query.all():
        qs = survey.questions
        new_qs, placed, touched = [], False, False
        for q in qs:
            if q.get("key") in ("ask", "self"):
                if not placed:
                    new_qs.append(dict(stance_q, options=list(stance_q["options"])))
                    placed = True
                touched = True
                continue
            base = core.get(q.get("key"))
            if base is not None:
                if q.get("text") in _PREVIOUS_TEXTS.get(q["key"], set()):
                    q["text"] = base["text"]
                    touched = True
                if base["type"] == "choice" and list(q.get("options") or []) != list(base["options"]):
                    q["options"] = list(base["options"])
                    touched = True
                if bool(q.get("required")) != bool(base.get("required")):
                    q["required"] = bool(base.get("required"))
                    touched = True
                if (base.get("depends") or None) != (q.get("depends") or None):
                    if base.get("depends"):
                        q["depends"] = dict(base["depends"])
                    else:
                        q.pop("depends", None)
                    touched = True
            new_qs.append(q)
        # 基本設問の並びを定義の順に揃える（追加設問の位置はそのまま）
        core_order = [q["key"] for q in C.MONTHLY_CORE_QUESTIONS]
        slots = [i for i, q in enumerate(new_qs) if q.get("key") in core]
        core_sorted = sorted((new_qs[i] for i in slots), key=lambda q: core_order.index(q["key"]))
        if [new_qs[i]["key"] for i in slots] != [q["key"] for q in core_sorted]:
            for i, q in zip(slots, core_sorted):
                new_qs[i] = q
            touched = True
        if touched:
            survey.set_questions(new_qs)
            changed = True
    for r in MonthlyResponse.query.all():
        answers = r.answers
        touched = False
        if "stance" not in answers and ("ask" in answers or "self" in answers):
            value = _stance_from_old(answers.get("ask", ""), answers.get("self", ""))
            if value:
                answers["stance"] = value
                touched = True
        if answers.get("hours") in _HOURS_WEEK_TO_SHIFT:
            answers["hours"] = _HOURS_WEEK_TO_SHIFT[answers["hours"]]
            touched = True
        if not answers.get("hours") and answers.get("has") == "ない":  # 時間の設問は必須になった。「ない」なら 0 時間
            answers["hours"] = "0時間"
            touched = True
        if answers.get("trend") == "初回なので分からない":
            answers["trend"] = ""
            touched = True
        if answers.get("stance") in _STANCE_OLD_TO_NEW:
            answers["stance"] = _STANCE_OLD_TO_NEW[answers["stance"]]
            touched = True
        if answers.get("wants") == "ない" and answers.get("stance"):  # 「ない」の人には聞かない
            answers["stance"] = ""
            touched = True
        if touched:
            r.answers_json = json.dumps(answers, ensure_ascii=False)
            changed = True
    return changed


# ------------------------------------------------------------
# デモデータ
# ------------------------------------------------------------
DEMO_RESPONDENTS = [
    # name, workplace, c1, skills, learn, c4, c6
    ("田中 一郎", "製造1課", "教えてもらえれば取り組みたい", ["Excel関数"], ["Excel VBA", "RPA"], "1〜2時間", "できる"),
    ("佐藤 花子", "製造1課", "ぜひ取り組みたい", ["Excel関数", "ピボットテーブル", "Excel VBA"], ["Python", "Power BI"], "2〜4時間", "できる"),
    ("鈴木 健", "製造2課", "興味はあるが時間がない", ["Excel関数"], ["Excel VBA"], "1時間未満", "内容による"),
    ("高橋 美咲", "組立課", "教えてもらえれば取り組みたい", ["Excel関数"], ["Power Query", "RPA"], "1〜2時間", "できる"),
    ("伊藤 大輔", "検査・品質保証課", "ぜひ取り組みたい", ["Excel関数", "ピボットテーブル", "Access"], ["Python", "生成AI"], "2〜4時間", "できる"),
    ("渡辺 由美", "生産管理課", "今は考えていない", ["Excel関数", "ピボットテーブル"], [], "1時間未満", "難しい"),
    ("山本 剛", "設備保全課", "教えてもらえれば取り組みたい", ["特になし"], ["Excel関数", "Excel VBA"], "1〜2時間", "内容による"),
    ("中村 翔太", "製造2課", "興味はあるが時間がない", ["Excel関数", "ピボットテーブル"], ["Power BI"], "1時間未満", "内容による"),
]

# name, respondent, domain, purpose, output, users, inputs, prev, next, freq, dur, head, tools, nature, pains, wish, self, note
DEMO_TASKS = [
    ("日報の実績をExcelの集計表に転記する", "田中 一郎", "製造実行・管理", "日次の生産実績を職場の朝会で確認するため", "日次実績集計表（Excel）", ["職場内", "上司"], ["紙の帳票"], "各ラインの日報記入", "週報の作成", "毎日", "15〜30分", 2, ["Excel", "紙"], ["転記・コピー", "集計・グラフ作成"], ["時間がかかる", "同じ情報を何度も入力する"], "自動化したい", "教えてもらえれば取り組みたい", "日報を直接システムに入力できれば転記がなくなる"),
    ("週報の作成", "田中 一郎", "製造実行・管理", "週次の生産状況を課長に報告するため", "週報（Excel）", ["上司"], ["他のExcel"], "日次実績集計表の転記", "月報の作成", "毎週", "30分〜1時間", 1, ["Excel"], ["集計・グラフ作成", "資料作成"], ["時間がかかる"], "減らしたい", "改善チームに任せたい", ""),
    ("残業予定表の提出", "田中 一郎", "人員・労務管理", "翌日の残業人員を総務に伝えるため", "残業予定表（Excel）をメール送付", ["他部署"], ["人から聞く"], "", "", "毎日", "15分未満", 1, ["Excel", "メール"], ["資料作成", "メール・配布"], ["急ぎで割り込まれる"], "減らしたい", "分からない", ""),
    ("生産実績の月報作成", "佐藤 花子", "製造実行・管理", "月次の生産実績・稼働率を部長会議に報告するため", "月報（PowerPoint・Excel）", ["上司", "他部署"], ["他のExcel", "生産管理システム"], "週報の作成", "部長会議", "毎月", "2時間以上", 1, ["Excel"], ["集計・グラフ作成", "資料作成"], ["時間がかかる", "ミスが起きやすい"], "自動化したい", "取り組みたい", ""),
    ("不良品発生報告のメール配信", "佐藤 花子", "品質管理", "不良発生を関係部署に共有するため", "不良報告メール", ["他部署", "上司"], ["紙の帳票", "人から聞く"], "品質システムへの不良登録", "", "週に数回", "15〜30分", 1, ["メール", "紙"], ["転記・コピー", "メール・配布"], ["同じ情報を何度も入力する"], "なくしたい", "取り組みたい", "品質システムに入力した内容と同じことをメールにも書いている"),
    ("出勤簿の集計", "佐藤 花子", "人員・労務管理", "月次の勤怠を総務に提出するため", "勤怠集計表", ["他部署"], ["紙の帳票"], "", "総務での給与計算", "毎月", "1〜2時間", 1, ["Excel", "紙"], ["転記・コピー", "集計・グラフ作成"], ["時間がかかる", "ミスが起きやすい"], "自動化したい", "教えてもらえれば取り組みたい", ""),
    ("安全パトロール結果の入力", "佐藤 花子", "安全・環境管理", "パトロール指摘事項を記録し是正を管理するため", "指摘事項台帳", ["職場内", "上司"], ["紙の帳票"], "安全パトロール", "是正確認", "毎週", "15〜30分", 1, ["Excel", "紙"], ["台帳更新"], ["特になし"], "このままでよい", "改善チームに任せたい", ""),
    ("段取り替え時間の記録と集計", "鈴木 健", "原価・生産性管理", "段取り時間の推移を把握し改善テーマを決めるため", "段取り時間推移グラフ", ["職場内", "上司"], ["紙の帳票"], "段取り替え作業", "改善会議", "毎日", "15〜30分", 1, ["Excel", "紙"], ["転記・コピー", "集計・グラフ作成"], ["時間がかかる"], "自動化したい", "教えてもらえれば取り組みたい", ""),
    ("設備日常点検表のファイリング確認", "鈴木 健", "設備管理", "点検漏れがないか確認するため", "点検実施チェック表", ["職場内"], ["紙の帳票"], "日常点検", "", "毎日", "15分未満", 1, ["紙", "Excel"], ["確認・チェック"], ["特になし"], "分からない", "分からない", ""),
    ("原材料の払出し依頼票の作成", "鈴木 健", "生産計画・管理", "翌日分の材料を倉庫に依頼するため", "払出し依頼票（Excel）を印刷", ["他部署"], ["生産管理システム", "他のExcel"], "週次生産計画の受領", "倉庫での払出し", "毎日", "15〜30分", 2, ["Excel", "紙"], ["転記・コピー", "資料作成"], ["同じ情報を何度も入力する"], "なくしたい", "教えてもらえれば取り組みたい", "生産計画の数字を手で入れ直している"),
    ("課内会議資料の作成", "中村 翔太", "製造実行・管理", "週次の課内会議で実績と課題を共有するため", "会議資料（PowerPoint）", ["職場内"], ["他のExcel"], "週報の作成", "課内会議", "毎週", "1〜2時間", 1, ["Excel"], ["資料作成", "集計・グラフ作成"], ["時間がかかる"], "減らしたい", "改善チームに任せたい", ""),
    ("生産性指標の報告（部へ）", "中村 翔太", "原価・生産性管理", "部の生産性会議に課の指標を提出するため", "生産性指標シート", ["他部署", "上司"], ["他のExcel"], "月報の作成", "部の生産性会議", "毎月", "30分〜1時間", 1, ["Excel", "メール"], ["転記・コピー"], ["同じ情報を何度も入力する"], "なくしたい", "改善チームに任せたい", "月報と同じ数字を別フォーマットに入れ直している"),
    ("組立実績の日次入力（生産管理システム）", "高橋 美咲", "製造実行・管理", "進捗管理と在庫引き落としのため", "システム上の実績データ", ["他部署", "職場内"], ["紙の帳票"], "組立作業の日報記入", "生産管理課の進捗確認", "毎日", "30分〜1時間", 1, ["社内システム", "紙"], ["転記・コピー"], ["時間がかかる", "ミスが起きやすい"], "自動化したい", "教えてもらえれば取り組みたい", ""),
    ("外注先への納期確認メール", "高橋 美咲", "生産計画・管理", "外注部品の納期遅れを早めに把握するため", "納期確認メールと回答一覧", ["職場内", "上司"], ["メール", "他のExcel"], "発注", "組立計画の調整", "週に数回", "30分〜1時間", 1, ["メール", "Excel"], ["メール・配布", "確認・チェック"], ["急ぎで割り込まれる", "手順が人によって違う"], "減らしたい", "教えてもらえれば取り組みたい", ""),
    ("作業手順書の改訂管理台帳の更新", "高橋 美咲", "改善・技術開発", "手順書の版数を管理するため", "改訂管理台帳", ["職場内"], ["人から聞く"], "手順書の改訂", "", "不定期", "15〜30分", 1, ["Excel"], ["台帳更新"], ["特になし"], "このままでよい", "分からない", ""),
    ("日次実績のグラフ更新（掲示用）", "高橋 美咲", "製造実行・管理", "不明", "掲示板の実績グラフ", ["分からない"], ["他のExcel"], "日次実績集計表", "", "毎日", "15分未満", 1, ["Excel"], ["集計・グラフ作成"], ["何のためか分からない"], "分からない", "分からない", "前任者から引き継いだが誰が見ているか分からない"),
    ("検査成績書の作成", "伊藤 大輔", "品質管理", "出荷時に顧客へ品質を証明するため", "検査成績書（PDF）", ["顧客・社外"], ["品質システム", "他のExcel"], "出荷検査", "出荷", "毎日", "30分〜1時間", 2, ["Excel", "社内システム"], ["転記・コピー", "資料作成"], ["時間がかかる", "ミスが起きやすい"], "自動化したい", "取り組みたい", ""),
    ("不良データの集計とパレート図作成", "伊藤 大輔", "品質管理", "月次品質会議で重点不良を決めるため", "不良パレート図", ["職場内", "上司", "他部署"], ["品質システム"], "不良登録", "品質会議", "毎月", "1〜2時間", 1, ["Excel"], ["集計・グラフ作成"], ["時間がかかる"], "自動化したい", "取り組みたい", ""),
    ("測定器の校正期限チェック", "伊藤 大輔", "設備管理", "校正切れの測定器を使わないため", "校正期限一覧", ["職場内"], ["他のExcel"], "", "校正依頼", "毎週", "15分未満", 1, ["Excel"], ["確認・チェック"], ["特になし"], "自動化したい", "教えてもらえれば取り組みたい", ""),
    ("クレーム対応の進捗報告", "伊藤 大輔", "品質管理", "顧客クレームの対応状況を上司・営業に共有するため", "進捗報告メール", ["上司", "他部署"], ["人から聞く", "メール"], "クレーム受付", "顧客への回答", "週に数回", "15〜30分", 1, ["メール", "Excel"], ["資料作成", "メール・配布"], ["急ぎで割り込まれる"], "減らしたい", "改善チームに任せたい", ""),
    ("週次生産計画の作成と配布", "渡辺 由美", "生産計画・管理", "各課に翌週の生産順序を指示するため", "週次生産計画表", ["他部署"], ["生産管理システム", "メール"], "受注一覧の更新", "各課の段取り・材料手配", "毎週", "2時間以上", 1, ["Excel", "社内システム", "メール"], ["集計・グラフ作成", "資料作成", "メール・配布"], ["時間がかかる", "急ぎで割り込まれる"], "減らしたい", "改善チームに任せたい", ""),
    ("受注情報のExcel転記", "渡辺 由美", "生産計画・管理", "営業からの受注を計画表に反映するため", "受注一覧（Excel）", ["自分", "職場内"], ["メール"], "営業の受注入力", "週次生産計画の作成", "毎日", "30分〜1時間", 1, ["Excel", "メール"], ["転記・コピー"], ["同じ情報を何度も入力する", "ミスが起きやすい"], "自動化したい", "改善チームに任せたい", "営業が販売システムに入れた情報をメールで受けて転記している"),
    ("在庫報告書の作成", "渡辺 由美", "生産計画・管理", "月末在庫を経理に報告するため", "在庫報告書", ["他部署"], ["生産管理システム"], "月末棚卸", "経理の決算処理", "毎月", "1〜2時間", 1, ["Excel", "社内システム"], ["集計・グラフ作成", "資料作成"], ["時間がかかる"], "自動化したい", "改善チームに任せたい", ""),
    ("設備故障記録の台帳入力", "山本 剛", "設備管理", "故障履歴を残し予防保全に活かすため", "故障履歴台帳", ["職場内", "上司"], ["紙の帳票", "人から聞く"], "故障対応", "月次保全報告書の作成", "週に数回", "15〜30分", 1, ["Excel", "紙"], ["台帳更新", "転記・コピー"], ["手順が人によって違う"], "減らしたい", "教えてもらえれば取り組みたい", ""),
    ("保全部品の発注依頼書作成", "山本 剛", "設備管理", "部品を購買に発注してもらうため", "発注依頼書", ["他部署"], ["人から聞く", "他のExcel"], "故障対応・点検", "購買での発注", "週に数回", "15〜30分", 1, ["Excel", "メール"], ["資料作成", "メール・配布"], ["同じ情報を何度も入力する"], "減らしたい", "分からない", ""),
    ("月次保全報告書の作成", "山本 剛", "設備管理", "設備稼働率と故障件数を部長に報告するため", "月次保全報告書", ["上司"], ["他のExcel"], "故障履歴台帳の入力", "部長への報告", "毎月", "1〜2時間", 1, ["Excel"], ["集計・グラフ作成", "資料作成"], ["時間がかかる"], "自動化したい", "教えてもらえれば取り組みたい", ""),
]

# 進捗のあるデモ案件: name -> dict
# 案件の進み具合（status と stage は set_stage と同じ対応: 受付=start / 対応中=request〜check / 完了=done）
DEMO_PROGRESS = {
    "日報の実績をExcelの集計表に転記する": dict(status="対応中", stage="spec", steps_done="1|2", decision="自動化", assignment="自主改善", assignee="田中 一郎", verified=True, verified_by="推進メンバー",
                                        progress_note="朝会で使う指標を2つに絞り込み中。次は日次表から週報を自動で作る仕様を決めます",
                                        ecrs_eliminate="朝会で使う指標は出来高と不良数のみ。他の項目は転記不要", ecrs_combine="週報と同じ数字なので週報を日次表から自動生成する"),
    "不良品発生報告のメール配信": dict(status="完了", stage="done", steps_done="1|2|3", decision="廃止", assignment="自主改善", assignee="佐藤 花子", verified=True, verified_by="推進メンバー",
                              after_minutes=0.0, result_note="品質システムの登録通知メールで代替できることを確認し、手作業のメール作成を廃止した。",
                              progress_note="品質システムの通知メールで代替し、手作業のメール作成をなくしました",
                              ecrs_eliminate="品質システムの通知機能で関係部署に自動通知できる"),
    "出勤簿の集計": dict(status="完了", stage="done", steps_done="1|2|3|4", decision="自動化", assignment="改善チーム主導", assignee="デジタル改善チーム", verified=True, verified_by="推進メンバー",
                  after_minutes=10.0, result_note="紙の出勤簿を勤怠システムのCSV取込に変更し、Excelマクロで集計表を自動作成。確認のみ残った。",
                  progress_note="勤怠システムのCSV取込とマクロで自動集計。確認作業だけ残りました",
                  ecrs_rearrange="発生源（勤怠システム）から直接データを取る", ecrs_simplify="集計項目を総務が必要とする5項目に絞った"),
    "受注情報のExcel転記": dict(status="対応中", stage="build", steps_done="1|2|3", decision="統合", assignment="改善チーム主導", assignee="デジタル改善チーム", verified=True, verified_by="推進メンバー",
                        progress_note="販売システムからのCSV取込ツールを作成中。11月中旬に試行予定",
                        ecrs_combine="営業の販売システムから受注データをCSVで取得し、転記をなくす", ecrs_rearrange="営業側の入力時点で計画表に必要な項目を揃える"),
    "原材料の払出し依頼票の作成": dict(status="対応中", stage="request", steps_done="1", verified=True, verified_by="推進メンバー",
                            progress_note="倉庫・生産管理課と「依頼票を何に使っているか」をヒアリング中"),
    "検査成績書の作成": dict(status="対応中", stage="test", steps_done="1|2|3|4", decision="自動化", assignment="改善チーム主導", assignee="デジタル改善チーム", verified=True, verified_by="推進メンバー",
                     progress_note="品質システムから成績書を自動出力する仕組みを11月分の検査データでテスト中",
                     ecrs_simplify="成績書の様式を顧客要求の項目だけに絞った"),
    "週次生産計画の作成と配布": dict(status="対応中", stage="check", steps_done="1|2|3|4", decision="簡素化", assignment="自主改善", assignee="渡辺 由美", verified=True, verified_by="推進メンバー",
                         progress_note="新しい計画表で2週間運用し、各課リーダーに使い勝手を確認中",
                         ecrs_simplify="配布先ごとに作っていた3種類の計画表を1枚に統一"),
    "組立実績の日次入力（生産管理システム）": dict(status="保留", stage="spec", steps_done="1|2", assignment="改善チーム主導", assignee="デジタル改善チーム", verified=True, verified_by="推進メンバー",
                                  progress_note="生産管理システムの更新（12月）を待って再開します"),
}


def _seed_stage_history(task, now):
    """現在の段階までの履歴を、登録日から今までの間に等間隔で作る"""
    idx = task.stage_index
    task.stage_at = task.created_at
    if idx == 0:
        return
    created = task.created_at or now
    total = max((now - created).total_seconds() - 86400, 3600.0 * idx)
    for i in range(1, idx + 1):
        at = created + timedelta(seconds=total * i / idx)
        task.stage_logs.append(StageLog(stage=C.STAGE_KEYS[i], at=at, by=task.assignee or ""))
        task.stage_at = at
    if task.stage_key == "done":
        task.completed_at = task.stage_at

# ヒト・情報の流れ（B17・B18）: name -> dict
DEMO_FLOWS = {
    "日報の実績をExcelの集計表に転記する": dict(actors="各ライン作業者、課長", info_from="各ラインの手書き日報"),
    "週報の作成": dict(actors="課長", info_from="日次実績集計表（Excel）"),
    "残業予定表の提出": dict(actors="各ラインリーダー、総務", info_from="各ラインリーダー（口頭）"),
    "生産実績の月報作成": dict(actors="部長、生産管理課", info_from="週報・生産管理システム"),
    "不良品発生報告のメール配信": dict(actors="検査員、検査・品質保証課、生産管理課", info_from="検査員の不良報告書"),
    "出勤簿の集計": dict(actors="各作業者、総務", info_from="紙の出勤簿"),
    "原材料の払出し依頼票の作成": dict(actors="倉庫、生産管理課", info_from="生産管理課の週次生産計画表"),
    "課内会議資料の作成": dict(actors="課内メンバー", info_from="週報"),
    "生産性指標の報告（部へ）": dict(actors="部長", info_from="生産実績の月報"),
    "組立実績の日次入力（生産管理システム）": dict(actors="組立作業者、生産管理課", info_from="組立作業者の日報"),
    "外注先への納期確認メール": dict(actors="外注先、生産管理課", info_from="発注一覧（Excel）"),
    "検査成績書の作成": dict(actors="営業、顧客", info_from="品質システムの検査データ"),
    "不良データの集計とパレート図作成": dict(actors="検査員、製造各課", info_from="品質システムの不良登録データ"),
    "測定器の校正期限チェック": dict(actors="各職場、校正業者", info_from="校正台帳（Excel）"),
    "週次生産計画の作成と配布": dict(actors="営業、各課リーダー", info_from="受注一覧（Excel）・生産管理システム"),
    "受注情報のExcel転記": dict(actors="営業", info_from="営業からの受注メール"),
    "設備故障記録の台帳入力": dict(actors="設備オペレーター、保全担当", info_from="オペレーターの故障連絡（口頭・メモ）"),
    "保全部品の発注依頼書作成": dict(actors="購買、保全担当", info_from="保全担当のメモ"),
}

DEMO_D = {
    "田中 一郎": ("毎月、生産管理課へ実績を送っている", "日報の数字を日次表・週報・月報に3回入れている", "", "紙の日報をなくしてほしい"),
    "佐藤 花子": ("総務への勤怠提出", "不良報告は品質システムとメールの二重入力", "掲示用グラフの一部", "自動化したい業務が多い"),
    "渡辺 由美": ("営業からの受注連絡がメールで来るので転記が必要", "受注情報を受注一覧と計画表の両方に入れている", "", ""),
}


def seed_demo():
    """デモデータを投入する（既存のデモデータは削除してから投入）"""
    clear_demo()
    wp_by_name = {w.name: w for w in Workplace.query.all()}
    dom_by_name = {d.name: d for d in Domain.query.all()}
    base = datetime.now() - timedelta(days=20)

    resp_by_name = {}
    for i, (name, wp, c1, skills, learn, c4, c6) in enumerate(DEMO_RESPONDENTS):
        workplace = wp_by_name.get(wp)
        if workplace is None:
            workplace = Workplace(name=wp, sort=100 + i)
            db.session.add(workplace)
            db.session.flush()
            wp_by_name[wp] = workplace
        d = DEMO_D.get(name, ("", "", "", ""))
        r = Respondent(
            emp_id=f"E{i + 1:03d}",
            name=name, workplace=workplace, is_demo=True,
            c1_wish=c1, c2_skills=C.join_list(skills), c3_learn=C.join_list(learn), c4_hours=c4, c6_share=c6,
            part_c_done=True, part_d_done=bool(any(d)),
            d1=d[0], d2=d[1], d3=d[2], d4=d[3],
            created_at=base + timedelta(days=i),
        )
        db.session.add(r)
        resp_by_name[name] = r
    db.session.flush()

    tasks = []
    for i, row in enumerate(DEMO_TASKS):
        (name, rname, dname, purpose, output, users, inputs, prev, nxt, freq, dur, head, tools, nature, pains, wish, selfimp, note) = row
        r = resp_by_name[rname]
        domain = dom_by_name.get(dname)
        if domain is None:
            domain = Domain(name=dname, sort=100 + i)
            db.session.add(domain)
            db.session.flush()
            dom_by_name[dname] = domain
        t = Task(
            respondent=r, workplace_id=r.workplace_id, is_demo=True,
            name=name, domain=domain, purpose=purpose, output=output,
            output_users=C.join_list(users), inputs=C.join_list(inputs), prev_task=prev, next_task=nxt,
            frequency=freq, duration=dur, headcount=head, tools=C.join_list(tools), nature=C.join_list(nature),
            pains=C.join_list(pains), wish=wish, self_improve=selfimp, note=note,
            created_at=base + timedelta(days=i % 15, hours=i),
        )
        t.compute_minutes()
        flow = DEMO_FLOWS.get(name, {})
        for k, v in flow.items():
            setattr(t, k, C.join_list(v) if isinstance(v, list) else v)
        prog = DEMO_PROGRESS.get(name, {})
        for k, v in prog.items():
            setattr(t, k, v)
        _seed_stage_history(t, datetime.now())
        db.session.add(t)
        tasks.append(t)
    db.session.flush()

    domains = Domain.query.filter_by(parent_id=None).order_by(Domain.sort).all()
    for t in tasks:
        cands = related_tasks_for(t, limit=8)
        result = rule_analyze(t, domains, cands)
        result["source"] = "rules"
        t.apply_ai_result(result)
    _seed_employees_and_monthly(resp_by_name, wp_by_name)
    db.session.commit()
    return len(tasks)


DEMO_EXTRA_EMPLOYEES = [
    ("E009", "小林 直樹", "製造1課"),
    ("E010", "加藤 恵", "組立課"),
    ("E011", "吉田 悠", "検査・品質保証課"),
    ("E012", "松本 玲奈", "生産管理課"),
]


def recent_periods(n=4):
    """当月を含む直近 n か月の月度（古い順）"""
    now = datetime.now()
    periods = []
    for k in range(n - 1, -1, -1):
        yy, mm = now.year, now.month - k
        while mm <= 0:
            mm += 12
            yy -= 1
        periods.append(f"{yy:04d}-{mm:02d}")
    return periods


def _seed_employees_and_monthly(resp_by_name, wp_by_name):
    """社員マスタ（12人）と、直近4か月分の月次アンケート回答（改善が進む傾向）"""
    employees = []
    for i, (name, r) in enumerate(resp_by_name.items()):
        employees.append(
            Employee(
                emp_id=r.emp_id, name=name, department=r.workplace.name if r.workplace else "",
                workplace_id=r.workplace_id, email=f"user{i + 1:02d}@example.local", active=True, source="csv", is_demo=True,
            )
        )
    for j, (emp_id, name, wp) in enumerate(DEMO_EXTRA_EMPLOYEES):
        workplace = wp_by_name.get(wp)
        employees.append(
            Employee(
                emp_id=emp_id, name=name, department=wp, workplace_id=workplace.id if workplace else None,
                email=f"user{len(resp_by_name) + j + 1:02d}@example.local", active=True, source="csv", is_demo=True,
            )
        )
    for e in employees:
        db.session.add(e)
    db.session.flush()

    resp_by_emp = {r.emp_id: r for r in resp_by_name.values()}
    periods = recent_periods(4)
    last = len(periods) - 1
    hours_options = list(C.MONTHLY_HOURS_VALUE)
    stance_options = ["自分で取り組みたい", "改善チームに依頼したい"]
    # 月度ごとの設問セット（基本設問）。既にあればそのまま使う
    surveys = {}
    for period in periods:
        survey = MonthlySurvey.query.filter_by(period=period).first()
        if survey is None:
            survey = MonthlySurvey(period=period, is_active=True)
            survey.set_questions([dict(q, options=list(q.get("options", []))) for q in C.MONTHLY_CORE_QUESTIONS])
            db.session.add(survey)
        surveys[period] = survey
    db.session.flush()
    for i, emp in enumerate(employees):
        for p_idx, period in enumerate(periods):
            if (i * 7 + p_idx * 3) % 9 == 0 and p_idx != last:
                continue  # 一部は未回答（回答率 < 100% を再現）
            level = max(0, (i % 3) + 2 - p_idx)  # 月が進むほど繰り返し業務が減る
            prev_level = max(0, (i % 3) + 2 - (p_idx - 1))
            has = "ない" if level == 0 else ("少しある" if level == 1 else "ある")
            hours = hours_options[0] if not level else hours_options[min(len(hours_options) - 1, level + (1 if i % 4 == 0 else 0))]
            if p_idx == 0 or level == prev_level:
                trend = "変わらない"
            elif level < prev_level:
                trend = "減った"
            elif (i + p_idx) % 7 == 0:
                trend = "増えた"
            else:
                trend = "変わらない"
            wants = "ある" if level >= 1 and (i + p_idx) % 3 != 0 else "ない"
            stance_idx = max(0, min(1, (i % 3) - (1 if p_idx >= 2 else 0)))  # 月が進むほど「自分で取り組みたい」が増える
            r = resp_by_emp.get(emp.emp_id)
            answers = {
                "has": has,
                "hours": hours,
                "trend": trend,
                "wants": wants,
                "wants_tasks": ((r.tasks[0].name if r and r.tasks else "日報の実績の転記") if wants == "ある" else ""),
                "stance": stance_options[stance_idx] if wants == "ある" else "",
                "new_tasks": "新しい報告フォーマットへの転記" if (trend == "増えた" or (trend == "変わらない" and p_idx == last and i % 5 == 0)) else "",
            }
            comment = ""
            if p_idx == last and i % 4 == 1:
                comment = ["日報の転記が減って朝の時間に余裕ができました", "Excelの集計を自分で自動化してみたい", "新しい帳票が増えたので見直してほしい"][i % 3]
            yy, mm = (int(x) for x in period.split("-"))
            db.session.add(
                MonthlyResponse(
                    period=period, emp_id=emp.emp_id, name=emp.name, workplace_id=emp.workplace_id,
                    respondent_id=r.id if r else None, survey_id=surveys[period].id,
                    answers_json=json.dumps(answers, ensure_ascii=False), comment=comment,
                    is_demo=True, created_at=datetime(yy, mm, 5, 9, 0),
                )
            )


def clear_demo():
    demo_ids = db.session.query(Task.id).filter_by(is_demo=True)
    StageLog.query.filter(StageLog.task_id.in_(demo_ids)).delete(synchronize_session=False)
    Task.query.filter_by(is_demo=True).delete(synchronize_session=False)
    Respondent.query.filter_by(is_demo=True).delete(synchronize_session=False)
    MonthlyResponse.query.filter_by(is_demo=True).delete(synchronize_session=False)
    Employee.query.filter_by(is_demo=True).delete(synchronize_session=False)
    db.session.commit()


# ==============================================================================
# 画面: ログイン
# ==============================================================================
bp_auth = Blueprint("auth", __name__)

DEFAULT_ADMIN_ID = "admin"
DEFAULT_ADMIN_PASSWORD = "adminpass"


def admin_credentials():
    return (
        str(conf.get("ADMIN_ID", "") or DEFAULT_ADMIN_ID).strip(),
        str(conf.get("ADMIN_PASSWORD", "") or DEFAULT_ADMIN_PASSWORD),
    )


def authenticate(user_id, password):
    """ID とパスワードが管理者アカウントと一致すればユーザー情報を返す。違えば None"""
    admin_id, admin_password = admin_credentials()
    ok_id = hmac.compare_digest((user_id or "").strip(), admin_id)
    ok_pw = hmac.compare_digest(password or "", admin_password)
    if ok_id and ok_pw:
        return {"emp_id": admin_id, "name": "管理者"}
    return None


def auth_status_text():
    admin_id, admin_password = admin_credentials()
    if admin_id == DEFAULT_ADMIN_ID and admin_password == DEFAULT_ADMIN_PASSWORD:
        return f"ログイン: ID「{admin_id}」（既定のパスワード）。本番では config.py の ADMIN_ID / ADMIN_PASSWORD を変更してください"
    return f"ログイン: ID「{admin_id}」（config.py で設定したパスワード）"


def _safe_next(value):
    value = (value or "").strip()
    if value.startswith("/") and not value.startswith("//"):
        return value
    return None


@bp_auth.route("/login", methods=["GET", "POST"])
def login():
    nxt = _safe_next(request.values.get("next"))
    if session.get("user"):
        return redirect(nxt or url_for("main.index"))
    error = None
    user_id = ""
    if request.method == "POST":
        user_id = (request.form.get("emp_id") or "").strip()
        password = request.form.get("password") or ""
        if not user_id or not password:
            error = "IDとパスワードを入力してください"
        else:
            user = authenticate(user_id, password)
            if user is None:
                error = "IDまたはパスワードが正しくありません"
            else:
                session["user"] = user
                session.permanent = True
                return redirect(nxt or url_for("main.index"))
    return render_template("auth/login.html", error=error, emp_id=user_id, next=nxt or "")


@bp_auth.post("/logout")
def logout():
    session.pop("user", None)
    flash("ログアウトしました。", "secondary")
    return redirect(url_for("board.index"))


# ==============================================================================
# 画面: ダッシュボード
# ==============================================================================
bp_main = Blueprint("main", __name__)


def _resolve_urls(suggestions):
    for s in suggestions:
        endpoint, params = s.get("url", (None, {}))
        try:
            s["href"] = url_for(endpoint, **params) if endpoint else None
        except Exception:
            s["href"] = None
    return suggestions


@bp_main.route("/")
def index():
    if not session.get("user"):
        return redirect(url_for("board.index"))
    k = kpis()
    charts = dashboard_chart_data()
    suggestions = _resolve_urls(suggest_actions())
    try:
        ai_suggestions = json.loads(Setting.get("ai_suggestions", "[]") or "[]")
    except ValueError:
        ai_suggestions = []
    recent = Task.query.order_by(Task.created_at.desc()).limit(8).all()
    all_tasks = Task.query.all()
    monthly_rows = monthly_summary()
    charts["monthly"] = {
        "labels": [r["label"] for r in monthly_rows],
        "has_rate": [round(r["has_rate"], 1) for r in monthly_rows],
        "decreased_rate": [round(r["decreased_rate"], 1) for r in monthly_rows],
        "self_rate": [round(r["self_rate"], 1) for r in monthly_rows],
    }
    return render_template(
        "index.html",
        k=k,
        charts_json=json.dumps(charts, ensure_ascii=False),
        suggestions=suggestions,
        ai_suggestions=ai_suggestions,
        workplaces=workplace_stats(),
        recent=recent,
        monthly_rows=monthly_rows,
        monthly_latest=monthly_rows[-1] if monthly_rows else None,
        active_count=len([t for t in all_tasks if t.status == "対応中"]),
    )


@bp_main.post("/ai/suggest")
def ai_suggest():
    if not ai_enabled():
        flash("AI連携が無効です。ANTHROPIC_API_KEY を設定して再起動してください。", "warning")
        return redirect(url_for("main.index"))
    k = kpis()
    have, learn = skills_summary()
    context = {
        "登録業務件数": k["task_count"],
        "回答者数": k["respondent_count"],
        "対象者数": k["target_respondents"],
        "月間時間合計（時間）": round(k["total_hours"], 1),
        "案件候補数": k["candidate_count"],
        "完了件数": k["completed_count"],
        "年間削減工数（時間）": round(k["saved_annual_hours"], 1),
        "自主改善希望者数": k["self_improver_count"],
        "業務情報の充足率（%）": round(k["sufficient_rate"], 1),
        "領域別": [
            {
                "領域": d["domain"].name,
                "件数": d["count"],
                "案件候補": d["candidates"],
                "月間時間（時間）": round(d["minutes"] / 60, 1),
                "充足率（%）": round(d["fill_rate"]),
                "確認済み率（%）": round(d["verified_rate"]),
                "分析状態": d["state"],
            }
            for d in k["domain_rows"]
        ],
        "職場別": [
            {"職場": w["workplace"].name, "回答者": w["respondents"], "件数": w["tasks"], "月間時間（時間）": round(w["minutes"] / 60, 1)}
            for w in workplace_stats()
        ],
        "重複入力の指摘（入力元別件数）": [{"入力元": src, "件数": len(ts)} for src, ts in duplicate_input_tasks()],
        "同じアウトプットの重複作成": [{"アウトプット": g["output"], "件数": len(g["tasks"])} for g in duplicate_output_groups()[:5]],
        "目的不明の業務": [t.name for t in Task.query.all() if t.purpose_unknown][:10],
        "月間時間上位の案件候補": [
            {"業務": t.name, "月間時間（時間）": round((t.monthly_minutes or 0) / 60, 1), "改善希望": t.wish}
            for t in sorted([t for t in Task.query.all() if t.is_candidate], key=lambda x: -(x.monthly_minutes or 0))[:8]
        ],
        "保有スキル": dict(have.most_common()),
        "学びたい技術": dict(learn.most_common()),
    }
    result = suggest_actions_ai(context)
    if result is None:
        flash("AIからの提案を取得できませんでした。APIキーや接続を確認してください。", "danger")
    else:
        Setting.set("ai_suggestions", json.dumps(result, ensure_ascii=False))
        db.session.commit()
        flash(f"AIが {len(result)} 件の提案を生成しました。", "success")
    return redirect(url_for("main.index"))


# ==============================================================================
# 画面: 自動化依頼（1画面・JSON API）
# ==============================================================================
bp_survey = Blueprint("survey", __name__, url_prefix="/survey")


def _current():
    rid = session.get("respondent_id")
    return db.session.get(Respondent, rid) if rid else None


def _domains():
    return Domain.query.filter_by(parent_id=None).order_by(Domain.sort, Domain.id).all()


def _workplaces():
    return Workplace.query.order_by(Workplace.sort, Workplace.id).all()


class _JsonForm(dict):
    """JSON の辞書を request.form と同じ get / getlist で扱う"""

    def get(self, key, default=None, type=None):
        value = dict.get(self, key, default)
        if isinstance(value, list):
            value = value[0] if value else default
        if value is None:
            return default
        if type is not None:
            try:
                return type(value)
            except (TypeError, ValueError):
                return default
        return value if isinstance(value, str) else str(value)

    def getlist(self, key):
        value = dict.get(self, key)
        if value in (None, ""):
            return []
        return [str(v) for v in value] if isinstance(value, list) else [str(value)]


def _respondent_json(r):
    return {
        "id": r.id,
        "emp_id": r.emp_id or "",
        "name": r.name,
        "workplace_id": r.workplace_id,
        "workplace": r.workplace.name if r.workplace else "",
        "c1_wish": r.c1_wish or "",
        "c2_skills": r.skills,
        "c3_learn": r.learn,
        "c4_hours": r.c4_hours or "",
        "c5_targets": r.c5_targets or "",
        "c6_share": r.c6_share or "",
        "d1": r.d1 or "",
        "d2": r.d2 or "",
        "d3": r.d3 or "",
        "d4": r.d4 or "",
        "part_c_done": bool(r.part_c_done),
        "part_d_done": bool(r.part_d_done),
    }


def _task_json(t):
    return {
        "id": t.id,
        "name": t.name,
        "domain_id": t.domain_id,
        "purpose": t.purpose or "",
        "output": t.output or "",
        "output_users": t.output_users_list,
        "output_users_note": t.output_users_note or "",
        "inputs": t.inputs_list,
        "inputs_note": t.inputs_note or "",
        "prev_task": t.prev_task or "",
        "next_task": t.next_task or "",
        "frequency": t.frequency or "",
        "duration": t.duration or "",
        "headcount": t.headcount or 1,
        "tools": t.tools_list,
        "nature": t.nature_list,
        "pains": t.pains_list,
        "pains_note": t.pains_note or "",
        "wish": t.wish or "",
        "self_improve": t.self_improve or "",
        "note": t.note or "",
        "actors": t.actors or "",
        "info_from": t.info_from or "",
        "monthly_minutes": t.monthly_minutes or 0,
        "hours": round((t.monthly_minutes or 0) / 60.0, 1),
        "ai_flags": t.ai_flags_list,
    }


def _state(r):
    if r is None:
        return {"respondent": None, "tasks": []}
    tasks = Task.query.filter_by(respondent_id=r.id).order_by(Task.created_at).all()
    return {"respondent": _respondent_json(r), "tasks": [_task_json(t) for t in tasks]}


def _options():
    return {
        "workplaces": [{"id": w.id, "name": w.name} for w in _workplaces()],
        "domains": [{"id": d.id, "name": d.name} for d in _domains()],
        "person": survey_spec.PERSON_FIELDS,
        "task_sections": survey_spec.TASK_SECTIONS,
        "c_fields": survey_spec.C_FIELDS,
        "d_fields": survey_spec.D_FIELDS,
    }


@bp_survey.route("/")
def index():
    emp = (request.args.get("emp") or "").strip()
    return render_template("survey/index.html", state=_state(_current()), options=_options(), emp=emp)


@bp_survey.route("/start")
def start():
    """旧リンク用（メールの誘導など）。1画面の入口へ"""
    emp = (request.args.get("emp") or "").strip()
    return redirect(url_for("survey.index", emp=emp or None))


# ------------------------------------------------------------
# API: あなたのこと
# ------------------------------------------------------------
@bp_survey.post("/api/respondent")
def api_respondent():
    f = _JsonForm(request.get_json(silent=True) or {})
    emp_id = (f.get("emp_id") or "").strip()
    name = (f.get("name") or "").strip()
    workplace_id = f.get("workplace_id", type=int)
    if emp_id and (not name or not workplace_id):
        info = lookup(emp_id)
        if info:
            name = name or info["name"]
            workplace_id = workplace_id or info.get("workplace_id")
        if not name or not workplace_id:  # 以前に登録した人・月次に答えた人は、その内容で
            r_name, r_wp = remembered_person(emp_id)
            name = name or r_name
            workplace_id = workplace_id or r_wp
    errors = {}
    if not emp_id:
        errors["emp_id"] = "入力してください"
    if not name:
        errors["name"] = "入力してください"
    if not workplace_id:
        errors["workplace_id"] = "選択してください"
    if errors:
        return jsonify({"success": False, "errors": errors}), 400
    r = Respondent.query.filter_by(emp_id=emp_id).first()  # 同定は社員IDだけで行う
    resumed = r is not None
    if r is None:
        r = Respondent(emp_id=emp_id)
        db.session.add(r)
    r.emp_id = emp_id
    r.name = name
    r.workplace_id = workplace_id
    remember_person(emp_id, name, workplace_id)  # 照会に氏名・職場がない環境では、ここで控えに残す
    db.session.commit()
    session["respondent_id"] = r.id
    return jsonify({"success": True, "resumed": resumed, **_state(r)})


@bp_survey.post("/api/reset")
def api_reset():
    session.pop("respondent_id", None)
    return jsonify({"success": True})


# ------------------------------------------------------------
# API: 依頼する業務
# ------------------------------------------------------------
def _fill_task(task, form, errors):
    task.name = (form.get("name") or "").strip()
    task.domain_id = form.get("domain_id", type=int)
    task.purpose = (form.get("purpose") or "").strip()
    task.output = (form.get("output") or "").strip()
    task.output_users = C.join_list(form.getlist("output_users"))
    task.output_users_note = (form.get("output_users_note") or "").strip()
    task.inputs = C.join_list(form.getlist("inputs"))
    task.inputs_note = (form.get("inputs_note") or "").strip()
    task.prev_task = (form.get("prev_task") or "").strip()
    task.next_task = (form.get("next_task") or "").strip()
    task.frequency = form.get("frequency")
    task.duration = form.get("duration")
    task.headcount = form.get("headcount", type=int) or 1
    task.tools = C.join_list(form.getlist("tools"))
    task.nature = C.join_list(form.getlist("nature"))
    task.pains = C.join_list(form.getlist("pains"))
    task.pains_note = (form.get("pains_note") or "").strip()
    task.wish = form.get("wish")
    task.self_improve = form.get("self_improve")
    task.note = (form.get("note") or "").strip()
    # つながり（任意）
    task.actors = (form.get("actors") or "").strip()
    task.info_from = (form.get("info_from") or "").strip()

    if not task.name:
        errors["name"] = "入力してください"
    if not task.domain_id:
        errors["domain_id"] = "選択してください"
    if not task.purpose:
        errors["purpose"] = "入力してください（分からなければ「不明」）"
    if not task.output:
        errors["output"] = "入力してください"
    if not task.output_users:
        errors["output_users"] = "1つ以上選んでください"
    if not task.inputs:
        errors["inputs"] = "1つ以上選んでください"
    if task.frequency not in C.FREQUENCIES:
        errors["frequency"] = "選択してください"
    if task.duration not in C.DURATIONS:
        errors["duration"] = "選択してください"
    if task.headcount < 1:
        errors["headcount"] = "1以上にしてください"
    if not task.tools:
        errors["tools"] = "1つ以上選んでください"
    if not task.nature:
        errors["nature"] = "1つ以上選んでください"
    if not task.pains:
        errors["pains"] = "選んでください（なければ「特になし」）"
    if task.wish not in C.WISHES:
        errors["wish"] = "選択してください"
    if task.self_improve not in C.SELF_IMPROVE:
        errors["self_improve"] = "選択してください"
    task.compute_minutes()


def _run_ai(task):
    domains = _domains()
    cands = related_tasks_for(task, limit=40)
    if len(cands) < 10:
        extra = Task.query.filter(Task.id != task.id).order_by(Task.created_at.desc()).limit(20).all()
        seen = {t.id for t in cands}
        cands += [t for t in extra if t.id not in seen]
    result = analyze_task(task, domains, cands)
    task.apply_ai_result(result)
    return result


def _need_respondent():
    r = _current()
    if r is None:
        return None, (jsonify({"success": False, "message": "先に「あなたのこと」を保存してください", "errors": {"_": "respondent"}}), 400)
    return r, None


@bp_survey.post("/api/tasks")
def api_task_create():
    r, err = _need_respondent()
    if err:
        return err
    errors = {}
    task = Task()
    _fill_task(task, _JsonForm(request.get_json(silent=True) or {}), errors)
    if errors:
        return jsonify({"success": False, "errors": errors}), 400
    task.respondent_id = r.id
    task.workplace_id = r.workplace_id
    db.session.add(task)
    db.session.commit()
    result = _run_ai(task)
    db.session.commit()
    return jsonify({"success": True, "task": _task_json(task), "flags": result.get("flags") or [], "tasks": _state(r)["tasks"]})


@bp_survey.route("/api/tasks/<int:task_id>", methods=["PUT", "DELETE"])
def api_task_update(task_id):
    r, err = _need_respondent()
    if err:
        return err
    task = db.session.get(Task, task_id)
    if task is None or task.respondent_id != r.id:
        abort(404)
    if request.method == "DELETE":
        db.session.delete(task)
        db.session.commit()
        return jsonify({"success": True, "tasks": _state(r)["tasks"]})
    errors = {}
    _fill_task(task, _JsonForm(request.get_json(silent=True) or {}), errors)
    if errors:
        db.session.rollback()
        return jsonify({"success": False, "errors": errors}), 400
    db.session.commit()
    result = _run_ai(task)
    db.session.commit()
    return jsonify({"success": True, "task": _task_json(task), "flags": result.get("flags") or [], "tasks": _state(r)["tasks"]})


# ------------------------------------------------------------
# API: 改善への関わり方・部全体について
# ------------------------------------------------------------
@bp_survey.post("/api/part-c")
def api_part_c():
    r, err = _need_respondent()
    if err:
        return err
    f = _JsonForm(request.get_json(silent=True) or {})
    errors = {}
    c1 = f.get("c1_wish")
    skills = f.getlist("c2_skills")
    if c1 not in C.C1_OPTIONS:
        errors["c1_wish"] = "選択してください"
    if not skills:
        errors["c2_skills"] = "1つ以上選んでください（なければ「特になし」）"
    if errors:
        return jsonify({"success": False, "errors": errors}), 400
    r.c1_wish = c1
    r.c2_skills = C.join_list(skills)
    r.c3_learn = C.join_list(f.getlist("c3_learn"))
    r.c4_hours = f.get("c4_hours") or None
    r.c5_targets = (f.get("c5_targets") or "").strip()
    r.c6_share = f.get("c6_share") or None
    r.part_c_done = True
    db.session.commit()
    return jsonify({"success": True, "respondent": _respondent_json(r)})


@bp_survey.post("/api/part-d")
def api_part_d():
    r, err = _need_respondent()
    if err:
        return err
    f = _JsonForm(request.get_json(silent=True) or {})
    r.d1 = (f.get("d1") or "").strip()
    r.d2 = (f.get("d2") or "").strip()
    r.d3 = (f.get("d3") or "").strip()
    r.d4 = (f.get("d4") or "").strip()
    r.part_d_done = True
    db.session.commit()
    return jsonify({"success": True, **_state(r)})


# ==============================================================================
# 画面: 改善案件
# ==============================================================================
bp_cases = Blueprint("cases", __name__, url_prefix="/cases")


@bp_cases.route("/")
def index():
    q = Task.query
    status = request.args.get("status") or ""
    workplace_id = request.args.get("workplace_id", type=int)
    domain_id = request.args.get("domain_id", type=int)
    wish = request.args.get("wish") or ""
    flt = request.args.get("filter") or ""
    keyword = (request.args.get("q") or "").strip()
    sort = request.args.get("sort") or "minutes"

    if status:
        q = q.filter(Task.status == status)
    stage = request.args.get("stage") or ""
    if stage in C.STAGE_INDEX:
        q = q.filter(Task.stage == stage)
    if workplace_id:
        q = q.filter(Task.workplace_id == workplace_id)
    if domain_id:
        q = q.filter(Task.domain_id == domain_id)
    if wish:
        q = q.filter(Task.wish == wish)
    if keyword:
        like = f"%{keyword}%"
        q = q.filter(db.or_(Task.name.like(like), Task.output.like(like), Task.purpose.like(like)))
    tasks = q.all()
    if flt == "candidates":
        tasks = [t for t in tasks if t.is_candidate]
    elif flt == "elimination":
        tasks = [t for t in tasks if t.is_elimination_candidate]
    elif flt == "self":
        tasks = [t for t in tasks if t.self_improve in C.SELF_IMPROVE_POSITIVE]
    elif flt == "unverified":
        tasks = [t for t in tasks if not t.verified]

    if sort == "new":
        tasks.sort(key=lambda t: t.created_at or datetime.min, reverse=True)
    elif sort == "name":
        tasks.sort(key=lambda t: t.name)
    else:
        tasks.sort(key=lambda t: t.monthly_minutes or 0, reverse=True)

    return render_template(
        "cases/index.html",
        tasks=tasks,
        workplaces=Workplace.query.order_by(Workplace.sort).all(),
        domains=Domain.query.filter_by(parent_id=None).order_by(Domain.sort).all(),
        args=request.args,
        total_minutes=sum(t.monthly_minutes or 0 for t in tasks),
    )


@bp_cases.route("/<int:task_id>")
def detail(task_id):
    task = db.session.get(Task, task_id)
    if task is None:
        abort(404)
    related = related_tasks_for(task, limit=6)
    related_by_ai = []
    names = set(task.ai_related_list)
    if names:
        related_by_ai = Task.query.filter(Task.name.in_(list(names)), Task.id != task.id).all()
    return render_template(
        "cases/detail.html",
        task=task,
        related=related,
        related_by_ai=related_by_ai,
        task_flow=build_task_flow(task).render(),
    )


@bp_cases.post("/<int:task_id>/update")
def update(task_id):
    task = db.session.get(Task, task_id)
    if task is None:
        abort(404)
    f = request.form
    task.progress_note = (f.get("progress_note") or "").strip()
    decision = f.get("decision")
    if decision in C.DECISIONS:
        task.decision = decision
    assignment = f.get("assignment")
    if assignment in C.ASSIGNMENTS:
        task.assignment = assignment
    task.assignee = (f.get("assignee") or "").strip()
    task.steps_done = C.join_list(sorted(set(f.getlist("steps")), key=lambda s: int(s) if s.isdigit() else 0))
    task.ecrs_eliminate = (f.get("ecrs_eliminate") or "").strip()
    task.ecrs_combine = (f.get("ecrs_combine") or "").strip()
    task.ecrs_rearrange = (f.get("ecrs_rearrange") or "").strip()
    task.ecrs_simplify = (f.get("ecrs_simplify") or "").strip()
    after = (f.get("after_minutes") or "").strip()
    if after == "":
        task.after_minutes = None
    else:
        try:
            task.after_minutes = max(float(after), 0.0)
        except ValueError:
            flash("改善後の月間時間は数値で入力してください。", "danger")
    task.result_note = (f.get("result_note") or "").strip()
    db.session.commit()
    flash("案件を更新しました。", "success")
    return redirect(url_for("cases.detail", task_id=task.id))


@bp_cases.post("/<int:task_id>/stage")
def set_stage(task_id):
    """活動状況の段階を1つ進める／戻す、または直接指定する"""
    task = db.session.get(Task, task_id)
    if task is None:
        abort(404)
    f = request.form
    key = f.get("stage")
    if key not in C.STAGE_INDEX:
        step = 1 if f.get("direction") == "next" else -1
        idx = min(max(task.stage_index + step, 0), len(C.STAGE_KEYS) - 1)
        key = C.STAGE_KEYS[idx]
    if task.is_on_hold:
        task.status = "対応中"  # 段階を動かしたら保留は解除する
    note = (f.get("progress_note") or "").strip()
    if note:
        task.progress_note = note
    changed = task.set_stage(key, (f.get("by") or task.assignee or "").strip())
    db.session.commit()
    if changed:
        flash(f"「{task.name}」を「{task.stage_name}」にしました。活動状況に反映されます。", "success")
    else:
        flash("段階は変わっていません。", "secondary")
    return redirect(f.get("next") or url_for("cases.detail", task_id=task.id))


@bp_cases.post("/<int:task_id>/hold")
def toggle_hold(task_id):
    """保留にする／保留を解除する（段階はそのまま）"""
    task = db.session.get(Task, task_id)
    if task is None:
        abort(404)
    if task.is_on_hold:
        task.status = "完了" if task.stage_key == "done" else ("受付" if task.stage_key == "start" else "対応中")
        flash("保留を解除しました。", "success")
    elif task.stage_key == "done":
        flash("完了した案件は保留にできません。", "warning")
    else:
        task.status = "保留"
        flash("保留にしました。活動状況では「保留中」に表示されます。", "success")
    db.session.commit()
    return redirect(url_for("cases.detail", task_id=task.id))


@bp_cases.post("/<int:task_id>/verify")
def verify(task_id):
    task = db.session.get(Task, task_id)
    if task is None:
        abort(404)
    task.verified = not task.verified
    task.verified_by = (request.form.get("verified_by") or "推進メンバー").strip() if task.verified else ""
    if task.verified and task.ai_domain and task.domain and task.ai_domain != task.domain.name and request.form.get("adopt_ai_domain"):
        d = Domain.query.filter_by(name=task.ai_domain).first()
        if d:
            task.domain = d
    db.session.commit()
    flash("担当者確認を更新しました。" if task.verified else "担当者確認を解除しました。", "success")
    return redirect(url_for("cases.detail", task_id=task.id))


@bp_cases.post("/<int:task_id>/reanalyze")
def reanalyze(task_id):
    task = db.session.get(Task, task_id)
    if task is None:
        abort(404)
    domains = Domain.query.filter_by(parent_id=None).order_by(Domain.sort).all()
    cands = related_tasks_for(task, limit=40)
    result = analyze_task(task, domains, cands)
    task.apply_ai_result(result)
    db.session.commit()
    flash("AI分析を更新しました（" + ("Claude API" if result.get("source") == "claude" else "ルールベース") + "）。", "success")
    return redirect(url_for("cases.detail", task_id=task.id))


@bp_cases.route("/export.csv")
def export_csv():
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(
        [
            "ID", "業務名", "回答者", "職場", "業務領域", "目的", "アウトプット", "利用者", "利用者補足", "入力情報", "入力情報補足",
            "前の業務", "後の業務", "関わる人・部署", "情報の受取元",
            "頻度", "1回あたり時間", "担当人数", "月間時間(分)", "使用ツール", "作業の性質", "困りごと", "困りごと補足",
            "改善の希望", "自分で改善", "補足", "ステータス", "進捗段階", "段階更新日", "活動状況のひとこと", "改善方針", "担当区分", "担当者", "手順完了", "改善後月間時間(分)", "削減(分/月)",
            "担当者確認", "AI領域", "AI優先度", "AIフラグ", "AI要約", "登録日",
        ]
    )
    for t in Task.query.order_by(Task.id).all():
        w.writerow(
            [
                t.id, t.name, t.respondent.name if t.respondent else "", t.workplace.name if t.workplace else "",
                t.domain.name if t.domain else "", t.purpose, t.output, "・".join(t.output_users_list), t.output_users_note,
                "・".join(t.inputs_list), t.inputs_note, t.prev_task, t.next_task,
                t.actors, t.info_from,
                t.frequency, t.duration, t.headcount,
                t.monthly_minutes, "・".join(t.tools_list), "・".join(t.nature_list), "・".join(t.pains_list), t.pains_note,
                t.wish, t.self_improve, t.note, t.status, t.stage_name, t.stage_since.strftime("%Y-%m-%d") if t.stage_since else "", t.progress_note,
                t.decision, t.assignment, t.assignee, "・".join(str(s) for s in t.steps_done_list),
                "" if t.after_minutes is None else t.after_minutes, t.saved_minutes, "済" if t.verified else "",
                t.ai_domain or "", t.ai_priority or "", "・".join(t.ai_flags_list), t.ai_summary or "",
                t.created_at.strftime("%Y-%m-%d %H:%M") if t.created_at else "",
            ]
        )
    data = buf.getvalue().encode("utf-8-sig")
    return Response(
        data,
        mimetype="text/csv",
        headers={"Content-Disposition": "attachment; filename=tasks_{}.csv".format(datetime.now().strftime("%Y%m%d"))},
    )


# ==============================================================================
# 画面: 業務マップ
# ==============================================================================
bp_bizmap = Blueprint("bizmap", __name__, url_prefix="/bizmap")

MAX_TASKS = 80  # 1枚の図に載せる業務数の上限（超える場合は絞り込みを促す）
KINDS = ("info",)  # 図は1枚（以前のヒト・モノの図はなくした）
KIND_LABELS = {"info": "業務マップ"}


def _filtered_tasks(domain_id, workplace_id):
    q = Task.query
    if domain_id:
        d = db.session.get(Domain, domain_id)
        ids = [domain_id] + ([c.id for c in d.children] if d else [])
        q = q.filter(Task.domain_id.in_(ids))
    if workplace_id:
        q = q.filter(Task.workplace_id == workplace_id)
    return q.order_by(Task.workplace_id, Task.created_at).all()


def _build(kind, tasks):
    return build_info_flow(tasks)


def _scope_label(workplace_id, domain_id):
    parts = []
    if workplace_id:
        w = db.session.get(Workplace, workplace_id)
        parts.append(w.name if w else f"職場 {workplace_id}")
    if domain_id:
        d = db.session.get(Domain, domain_id)
        parts.append(d.name if d else f"領域 {domain_id}")
    return "／".join(parts) if parts else "製造部全体"


def _merged_graph(kind, workplace_id, domain_id):
    tasks = _filtered_tasks(domain_id, workplace_id)
    truncated = len(tasks) > MAX_TASKS
    tasks = tasks[:MAX_TASKS]
    auto = _build(kind, tasks).to_graph()
    scope = scope_key(workplace_id, domain_id)
    row = FlowMap.query.filter_by(kind=kind, scope=scope).first()
    saved = None
    if row:
        try:
            saved = json.loads(row.data or "{}")
        except ValueError:
            saved = None
    merged = merge_graph(auto, saved)
    merged["meta"] = {
        "kind": kind,
        "kind_label": KIND_LABELS.get(kind, kind),
        "scope": scope,
        "scope_label": _scope_label(workplace_id, domain_id),
        "workplace_id": workplace_id,
        "domain_id": domain_id,
        "saved_at": row.updated_at.strftime("%Y/%m/%d %H:%M") if row else None,
        "task_count": len(tasks),
        "truncated": truncated,
        "auto_node_count": len(auto["nodes"]),
        "auto_edge_count": len(auto["edges"]),
    }
    return merged


@bp_bizmap.route("/")
def index():
    domain_id = request.args.get("domain_id", type=int)
    workplace_id = request.args.get("workplace_id", type=int)
    tasks = _filtered_tasks(domain_id, workplace_id)
    truncated = len(tasks) > MAX_TASKS
    tasks = tasks[:MAX_TASKS]

    flows, saved_at = {}, {}
    for kind in KINDS:
        merged = _merged_graph(kind, workplace_id, domain_id)
        flows[kind] = graph_to_diagram(merged).render() if merged["nodes"] else ""
        saved_at[kind] = merged["meta"]["saved_at"]
    rows = domain_analysis()
    return render_template(
        "bizmap/index.html",
        flows=flows,
        saved_at=saved_at,
        task_count=len(tasks),
        flow_info_count=len([t for t in tasks if t.has_flow_info]),
        source_rows=source_rows(tasks),
        handoff_rows=handoff_rows(tasks, [w.name for w in Workplace.query.all()]),
        truncated=truncated,
        max_tasks=MAX_TASKS,
        workplaces=Workplace.query.order_by(Workplace.sort, Workplace.id).all(),
        domains=Domain.query.filter_by(parent_id=None).order_by(Domain.sort, Domain.id).all(),
        args=request.args,
        rows=rows,
        dup_outputs=duplicate_output_groups(),
        dup_inputs=duplicate_input_tasks(),
        flow_rows=flow_rows(),
        total=sum(r["count"] for r in rows),
    )


@bp_bizmap.route("/edit")
def edit():
    kind = request.args.get("kind") or "info"
    if kind not in KINDS:
        abort(404)
    workplace_id = request.args.get("workplace_id", type=int)
    domain_id = request.args.get("domain_id", type=int)
    return render_template(
        "bizmap/edit.html",
        kind=kind,
        kind_label=KIND_LABELS[kind],
        workplace_id=workplace_id,
        domain_id=domain_id,
        scope_label=_scope_label(workplace_id, domain_id),
        cls_labels=CLS_LABELS,
        workplaces=Workplace.query.order_by(Workplace.sort, Workplace.id).all(),
        domains=Domain.query.filter_by(parent_id=None).order_by(Domain.sort, Domain.id).all(),
    )


# ------------------------------------------------------------
# グラフ API（エディタ用）
# ------------------------------------------------------------
@bp_bizmap.route("/api/graph", methods=["GET"])
def api_graph_get():
    kind = request.args.get("kind") or "info"
    if kind not in KINDS:
        abort(404)
    workplace_id = request.args.get("workplace_id", type=int)
    domain_id = request.args.get("domain_id", type=int)
    return jsonify(_merged_graph(kind, workplace_id, domain_id))


@bp_bizmap.route("/api/graph", methods=["POST"])
def api_graph_save():
    payload = request.get_json(silent=True) or {}
    kind = payload.get("kind") or "info"
    if kind not in KINDS:
        abort(404)
    workplace_id = payload.get("workplace_id") or None
    domain_id = payload.get("domain_id") or None
    try:
        data = sanitize_graph(payload.get("data"))
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    scope = scope_key(workplace_id, domain_id)
    row = FlowMap.query.filter_by(kind=kind, scope=scope).first()
    if row is None:
        row = FlowMap(kind=kind, scope=scope)
        db.session.add(row)
    row.data = json.dumps(data, ensure_ascii=False)
    row.updated_by = (payload.get("updated_by") or "")[:100]
    row.updated_at = datetime.now()
    db.session.commit()
    return jsonify({"ok": True, "saved_at": row.updated_at.strftime("%Y/%m/%d %H:%M")})


@bp_bizmap.route("/api/graph", methods=["DELETE"])
def api_graph_reset():
    kind = request.args.get("kind") or "info"
    if kind not in KINDS:
        abort(404)
    workplace_id = request.args.get("workplace_id", type=int)
    domain_id = request.args.get("domain_id", type=int)
    row = FlowMap.query.filter_by(kind=kind, scope=scope_key(workplace_id, domain_id)).first()
    if row:
        db.session.delete(row)
        db.session.commit()
    return jsonify({"ok": True})


# ------------------------------------------------------------
# 手修正の業務データへの書き戻し
# ------------------------------------------------------------
def _saved_graph(kind, workplace_id, domain_id):
    row = FlowMap.query.filter_by(kind=kind, scope=scope_key(workplace_id, domain_id)).first()
    if not row:
        return None
    try:
        return json.loads(row.data or "{}")
    except ValueError:
        return None


def _tasks_in_graph(saved):
    ids = set()
    for n in (saved or {}).get("nodes", []):
        nid = str(n.get("id", ""))
        if nid[:1] in ("t", "o") and nid[1:].isdigit():
            ids.add(int(nid[1:]))
    for eid in (saved or {}).get("suppressed_edges", []):
        for part in str(eid).split("->"):
            if part[:1] in ("t", "o") and part[1:].isdigit():
                ids.add(int(part[1:]))
    if not ids:
        return {}
    return {t.id: t for t in Task.query.filter(Task.id.in_(list(ids))).all()}


@bp_bizmap.route("/api/writeback", methods=["GET"])
def api_writeback_candidates():
    kind = request.args.get("kind") or "info"
    if kind not in KINDS:
        abort(404)
    workplace_id = request.args.get("workplace_id", type=int)
    domain_id = request.args.get("domain_id", type=int)
    saved = _saved_graph(kind, workplace_id, domain_id)
    cands = writeback_candidates(saved, _tasks_in_graph(saved))
    return jsonify({"candidates": [{"id": c["id"], "task_id": c["task_id"], "text": c["text"]} for c in cands]})


@bp_bizmap.route("/api/writeback", methods=["POST"])
def api_writeback_apply():
    payload = request.get_json(silent=True) or {}
    kind = payload.get("kind") or "info"
    if kind not in KINDS:
        abort(404)
    workplace_id = payload.get("workplace_id") or None
    domain_id = payload.get("domain_id") or None
    selected = set(str(x) for x in (payload.get("ids") or []))
    saved = _saved_graph(kind, workplace_id, domain_id)
    tasks_by_id = _tasks_in_graph(saved)
    cands = writeback_candidates(saved, tasks_by_id)
    applied = apply_writeback(cands, selected, tasks_by_id)
    db.session.commit()
    return jsonify({"ok": True, "applied": applied, "selected": len(selected)})


@bp_bizmap.route("/domain/<int:domain_id>")
def domain(domain_id):
    d = db.session.get(Domain, domain_id)
    if d is None:
        abort(404)
    row = next((r for r in domain_analysis() if r["domain"].id == d.id), None)
    ids = [d.id] + [c.id for c in d.children]
    tasks = Task.query.filter(Task.domain_id.in_(ids)).all()
    tasks.sort(key=lambda t: t.monthly_minutes or 0, reverse=True)
    merged = _merged_graph("info", None, d.id)
    return render_template(
        "bizmap/domain.html",
        d=d,
        row=row,
        tasks=tasks,
        missing=[t for t in tasks if not t.required_filled],
        unverified=[t for t in tasks if not t.verified],
        flows=[t for t in tasks if (t.prev_task or "").strip() or (t.next_task or "").strip()],
        info_flow=graph_to_diagram(merged).render() if merged["nodes"] else "",
    )


# ==============================================================================
# 画面: DX人財情報
# ==============================================================================
bp_people = Blueprint("people", __name__, url_prefix="/people")


@bp_people.route("/")
def index():
    workplace_id = request.args.get("workplace_id", type=int)
    wish = (request.args.get("wish") or "").strip()
    everyone = dx_people()
    people = dx_people(workplace_id=workplace_id, wish=wish or None) if (workplace_id or wish) else everyone
    have, learn = skills_summary()
    latest_period = max((p["latest_period"] for p in everyone if p["latest_period"]), default="")
    summary = {
        "total": len(everyone),
        "positive": len([p for p in everyone if p["positive"]]),
        "self_now": len([p for p in everyone if latest_period and p["latest_period"] == latest_period and p["stance_latest"] in C.MONTHLY_KPI["self_positive"]]),
        "period_label": period_label(latest_period) if latest_period else "月次の回答はまだありません",
        "share": len([p for p in everyone if p["share"] == "できる"]),
    }
    c1_counts = [(opt, len([p for p in everyone if p["wish"] == opt])) for opt in C.C1_OPTIONS]
    return render_template(
        "people/index.html",
        people=people,
        summary=summary,
        c1_counts=c1_counts,
        have=have.most_common(),
        learn=learn.most_common(),
        workplaces=Workplace.query.order_by(Workplace.sort, Workplace.id).all(),
        args={"workplace_id": workplace_id, "wish": wish},
    )


# ==============================================================================
# 画面: 管理
# ==============================================================================
bp_admin = Blueprint("admin", __name__, url_prefix="/admin")


@bp_admin.route("/")
def index():
    return render_template(
        "admin/index.html",
        workplaces=Workplace.query.order_by(Workplace.sort, Workplace.id).all(),
        domains=Domain.query.filter_by(parent_id=None).order_by(Domain.sort, Domain.id).all(),
        ai_status=ai_status_text(),
        ldap_status=ldap_status_text(),
        demo_count=Task.query.filter_by(is_demo=True).count(),
        respondent_count=Respondent.query.count(),
        task_count=Task.query.count(),
        employee_count=Employee.query.count(),
        auth_status=auth_status_text(),
    )


@bp_admin.post("/settings")
def save_settings():
    for key in ("project_name", "catchphrase", "sub_message", "target_respondents", "deadline"):
        Setting.set(key, (request.form.get(key) or "").strip())
    Setting.set("board_show_names", "1" if request.form.get("board_show_names") else "0")
    db.session.commit()
    flash("設定を保存しました。", "success")
    return redirect(url_for("admin.index"))


# ------------------------------------------------------------
# 職場・業務領域
# ------------------------------------------------------------
@bp_admin.post("/workplace/add")
def add_workplace():
    name = (request.form.get("name") or "").strip()
    if not name:
        flash("職場名を入力してください。", "danger")
    elif Workplace.query.filter_by(name=name).first():
        flash("同じ名前の職場があります。", "warning")
    else:
        max_sort = db.session.query(db.func.max(Workplace.sort)).scalar() or 0
        db.session.add(Workplace(name=name, sort=max_sort + 1))
        db.session.commit()
        flash("職場を追加しました。", "success")
    return redirect(url_for("admin.index"))


@bp_admin.post("/workplace/<int:workplace_id>/delete")
def delete_workplace(workplace_id):
    w = db.session.get(Workplace, workplace_id)
    if w is None:
        abort(404)
    if w.respondents:
        flash("回答者が紐づいている職場は削除できません。", "warning")
    else:
        db.session.delete(w)
        db.session.commit()
        flash("職場を削除しました。", "info")
    return redirect(url_for("admin.index"))


@bp_admin.post("/domain/add")
def add_domain():
    name = (request.form.get("name") or "").strip()
    if not name:
        flash("領域名を入力してください。", "danger")
    elif Domain.query.filter_by(name=name, parent_id=None).first():
        flash("同じ名前の領域があります。", "warning")
    else:
        max_sort = db.session.query(db.func.max(Domain.sort)).scalar() or 0
        db.session.add(Domain(name=name, description=(request.form.get("description") or "").strip(), sort=max_sort + 1))
        db.session.commit()
        flash("業務領域を追加しました。", "success")
    return redirect(url_for("admin.index"))


@bp_admin.post("/domain/<int:domain_id>/update")
def update_domain(domain_id):
    d = db.session.get(Domain, domain_id)
    if d is None:
        abort(404)
    name = (request.form.get("name") or "").strip()
    if name:
        d.name = name
    d.description = (request.form.get("description") or "").strip()
    d.sort = request.form.get("sort", type=int) or d.sort
    db.session.commit()
    flash("業務領域を更新しました。", "success")
    return redirect(url_for("admin.index"))


@bp_admin.post("/domain/<int:domain_id>/delete")
def delete_domain(domain_id):
    d = db.session.get(Domain, domain_id)
    if d is None:
        abort(404)
    if d.tasks or d.children:
        flash("業務が紐づいている領域は削除できません。先に業務の領域を変更してください。", "warning")
    else:
        db.session.delete(d)
        db.session.commit()
        flash("業務領域を削除しました。", "info")
    return redirect(url_for("admin.index"))


# ------------------------------------------------------------
# 社員マスタ
# ------------------------------------------------------------
@bp_admin.route("/employees")
def employees():
    rows = Employee.query.order_by(Employee.workplace_id, Employee.emp_id).all()
    return render_template(
        "admin/employees.html",
        employees=rows,
        active_count=Employee.query.filter_by(active=True).count(),
        ldap_status=ldap_status_text(),
        workplaces=Workplace.query.order_by(Workplace.sort, Workplace.id).all(),
    )


def _upsert_employee(emp_id, name, department, email, workplace_id=None, source="manual"):
    emp = Employee.query.filter_by(emp_id=emp_id).first()
    created = emp is None
    if emp is None:
        emp = Employee(emp_id=emp_id, name=name)
        db.session.add(emp)
    emp.name = name or emp.name
    emp.department = department if department is not None else emp.department
    emp.email = email if email is not None else emp.email
    emp.source = source
    if workplace_id:
        emp.workplace_id = workplace_id
    else:
        wp = workplace_for(emp.department)
        if wp is not None:
            emp.workplace_id = wp.id
    return created


@bp_admin.post("/employees/add")
def add_employee():
    emp_id = (request.form.get("emp_id") or "").strip()
    name = (request.form.get("name") or "").strip()
    if not emp_id or not name:
        flash("社員IDと氏名を入力してください。", "danger")
        return redirect(url_for("admin.employees"))
    created = _upsert_employee(
        emp_id, name, (request.form.get("department") or "").strip(), (request.form.get("email") or "").strip(),
        request.form.get("workplace_id", type=int), "manual",
    )
    db.session.commit()
    flash("社員を追加しました。" if created else "社員情報を更新しました。", "success")
    return redirect(url_for("admin.employees"))


@bp_admin.post("/employees/<int:employee_id>/toggle")
def toggle_employee(employee_id):
    e = db.session.get(Employee, employee_id)
    if e is None:
        abort(404)
    e.active = not e.active
    db.session.commit()
    return redirect(url_for("admin.employees"))


@bp_admin.post("/employees/<int:employee_id>/delete")
def delete_employee(employee_id):
    e = db.session.get(Employee, employee_id)
    if e is None:
        abort(404)
    db.session.delete(e)
    db.session.commit()
    flash("社員を削除しました。", "info")
    return redirect(url_for("admin.employees"))


_HEADER_ALIASES = {
    "emp_id": ["社員id", "社員番号", "社員コード", "emp_id", "employee_id", "employeeid", "id"],
    "name": ["氏名", "名前", "name", "displayname"],
    "department": ["所属", "部署", "職場", "department", "dept"],
    "email": ["メール", "メールアドレス", "email", "mail", "e-mail"],
}


def _detect_columns(fieldnames):
    mapping = {}
    for key, aliases in _HEADER_ALIASES.items():
        for f in fieldnames or []:
            if (f or "").strip().lower().replace(" ", "") in aliases:
                mapping[key] = f
                break
    return mapping


@bp_admin.post("/employees/import")
def import_employees():
    file = request.files.get("file")
    if file is None or not file.filename:
        flash("CSVファイルを選択してください。", "danger")
        return redirect(url_for("admin.employees"))
    raw = file.read()
    text = None
    for enc in ("utf-8-sig", "cp932", "utf-8"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        flash("文字コードを判別できませんでした（UTF-8 または Shift_JIS で保存してください）。", "danger")
        return redirect(url_for("admin.employees"))
    reader = csv.DictReader(io.StringIO(text))
    cols = _detect_columns(reader.fieldnames)
    if "emp_id" not in cols or "name" not in cols:
        flash("見出し行に「社員ID」と「氏名」が必要です。", "danger")
        return redirect(url_for("admin.employees"))
    created = updated = 0
    for row in reader:
        emp_id = (row.get(cols["emp_id"]) or "").strip()
        name = (row.get(cols["name"]) or "").strip()
        if not emp_id or not name:
            continue
        dept = (row.get(cols["department"]) or "").strip() if "department" in cols else None
        email = (row.get(cols["email"]) or "").strip() if "email" in cols else None
        if _upsert_employee(emp_id, name, dept, email, None, "csv"):
            created += 1
        else:
            updated += 1
    db.session.commit()
    flash(f"社員マスタを取り込みました（追加 {created} 人、更新 {updated} 人）。", "success")
    return redirect(url_for("admin.employees"))


@bp_admin.route("/employees/export.csv")
def export_employees():
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["社員ID", "氏名", "所属", "職場", "メール", "対象", "元"])
    for e in Employee.query.order_by(Employee.workplace_id, Employee.emp_id).all():
        w.writerow([e.emp_id, e.name, e.department, e.workplace.name if e.workplace else "", e.email, "1" if e.active else "0", e.source])
    return Response(buf.getvalue().encode("utf-8-sig"), mimetype="text/csv", headers={"Content-Disposition": "attachment; filename=employees.csv"})


# ------------------------------------------------------------
# デモデータ・エクスポート
# ------------------------------------------------------------
@bp_admin.post("/demo/load")
def load_demo():
    n = seed_demo()
    flash(f"デモデータを投入しました（業務 {n} 件、社員マスタと月次アンケート4か月分を含む）。", "success")
    return redirect(url_for("main.index"))


@bp_admin.post("/demo/clear")
def remove_demo():
    clear_demo()
    flash("デモデータを削除しました。", "info")
    return redirect(url_for("admin.index"))


@bp_admin.route("/export/respondents.csv")
def export_respondents():
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["ID", "社員ID", "氏名", "職場", "登録業務数", "月間時間(分)", "C1 改善意欲", "C2 使える技術", "C3 学びたい技術", "C4 割ける時間", "C5 取り組みたい業務", "C6 発表協力", "D1 他部署とのやりとり", "D2 重複入力", "D3 目的不明の帳票", "D4 自由記述", "登録日"])
    for r in Respondent.query.order_by(Respondent.id).all():
        w.writerow(
            [
                r.id, r.emp_id or "", r.name, r.workplace.name if r.workplace else "", len(r.tasks), r.total_minutes,
                r.c1_wish or "", "・".join(r.skills), "・".join(r.learn), r.c4_hours or "", r.c5_targets, r.c6_share or "",
                r.d1, r.d2, r.d3, r.d4, r.created_at.strftime("%Y-%m-%d %H:%M") if r.created_at else "",
            ]
        )
    data = buf.getvalue().encode("utf-8-sig")
    return Response(
        data,
        mimetype="text/csv",
        headers={"Content-Disposition": "attachment; filename=respondents_{}.csv".format(datetime.now().strftime("%Y%m%d"))},
    )


# ==============================================================================
# 画面: 月次アンケート
# ==============================================================================
bp_monthly = Blueprint("monthly", __name__, url_prefix="/monthly")

_PERIOD_RE = re.compile(r"^\d{4}-(0[1-9]|1[0-2])$")
QUESTION_TYPES = ("choice", "text")


def current_period():
    return datetime.now().strftime("%Y-%m")


def parse_period(value):
    value = (value or "").strip()
    return value if _PERIOD_RE.match(value) else current_period()


def answerable_periods():
    """回答を受け付ける月度: 当月と、設問セットがすでにある前月（月初に前月分を答える人のため）"""
    cur = current_period()
    prev = shift_period(cur, -1)
    periods = [cur]
    if MonthlySurvey.query.filter_by(period=prev).first() is not None:
        periods.append(prev)
    return periods


def public_period(value):
    """回答画面の月度。URLで別の月度を指定しても、受け付けていない月度なら当月にする"""
    period = parse_period(value)
    return period if period in answerable_periods() else current_period()


def shift_period(period, months):
    y, m = (int(x) for x in period.split("-"))
    m += months
    while m <= 0:
        m += 12
        y -= 1
    while m > 12:
        m -= 12
        y += 1
    return f"{y:04d}-{m:02d}"


def _workplaces():
    return Workplace.query.order_by(Workplace.sort, Workplace.id).all()


def default_questions():
    return [dict(q, options=list(q.get("options", []))) for q in C.MONTHLY_CORE_QUESTIONS]


def get_or_create_survey(period):
    """その月度の設問セット。なければ直近の月度（なければ基本設問）をコピーして作る"""
    survey = MonthlySurvey.query.filter_by(period=period).first()
    if survey is None:
        prev = MonthlySurvey.query.filter(MonthlySurvey.period < period).order_by(MonthlySurvey.period.desc()).first()
        survey = MonthlySurvey(period=period, is_active=True)
        survey.set_questions(prev.questions if prev else default_questions())
        db.session.add(survey)
        db.session.commit()
    return survey


def _wants_json():
    return request.is_json or "application/json" in (request.headers.get("Accept") or "")


def _person(emp_id):
    """社員IDから氏名・職場を引く。照会（LDAP → 控え）で見つからなくても、以前に答えた人なら分かる"""
    emp_id = (emp_id or "").strip()
    if not emp_id:
        return None
    info = lookup(emp_id)
    if info:
        return info
    name, workplace_id = remembered_person(emp_id)
    if not name and not workplace_id:
        return None
    workplace = db.session.get(Workplace, workplace_id) if workplace_id else None
    return {"emp_id": emp_id, "name": name, "department": "", "email": "", "workplace_id": workplace_id, "workplace": workplace.name if workplace else "", "source": "answers"}


# ------------------------------------------------------------
# 回答（ログイン不要）
# ------------------------------------------------------------
@bp_monthly.route("/")
def form():
    period = public_period(request.args.get("period"))
    survey = get_or_create_survey(period)
    emp_id = (request.args.get("emp") or "").strip()
    info = _person(emp_id)
    existing = MonthlyResponse.query.filter_by(period=period, emp_id=emp_id).first() if emp_id else None
    return render_template(
        "monthly/form.html",
        period=period,
        period_label=period_label(period),
        survey=survey,
        questions=survey.questions,
        emp_id=emp_id,
        info=info,
        existing=existing,
        workplaces=_workplaces(),
    )


@bp_monthly.post("/submit")
def submit():
    data = request.get_json(silent=True) if request.is_json else request.form
    data = data or {}
    period = public_period(data.get("period"))
    survey = get_or_create_survey(period)
    emp_id = (data.get("emp_id") or "").strip()
    name = (data.get("name") or "").strip()
    try:
        workplace_id = int(data.get("workplace_id") or 0) or None
    except (TypeError, ValueError):
        workplace_id = None
    info = _person(emp_id)
    if info:
        name = name or info["name"]
        workplace_id = workplace_id or info.get("workplace_id")

    errors = {}
    if not emp_id:
        errors["emp_id"] = "社員IDを入力してください"
    if not name:
        errors["name"] = "氏名を入力してください（照会で分からないときは初回だけ手入力）"
    if not workplace_id:
        errors["workplace_id"] = "職場を選んでください"
    raw = data.get("answers") if request.is_json else None
    answers = {}
    for q in survey.questions:
        key = q["key"]
        value = raw.get(key) if isinstance(raw, dict) else data.get("ans-" + key)
        value = value.strip() if isinstance(value, str) else ""
        if q["type"] == "choice":
            if value and value not in q.get("options", []):
                errors[key] = "選択肢から選んでください"
            elif not value and q.get("required"):
                errors[key] = "選択してください"
        elif not value and q.get("required"):
            errors[key] = "入力してください"
        answers[key] = value[:1000]
    for q in survey.questions:  # 条件つきの設問: 親の答えが条件に合うときだけ出す（合わなければ空に戻す）
        dep = q.get("depends") or {}
        if not dep.get("key"):
            continue
        parent = answers.get(dep["key"], "")
        show_if = dep.get("show_if")
        shown = (parent in show_if) if show_if is not None else (parent not in (dep.get("hidden_if") or []))
        if not shown:
            answers[q["key"]] = ""
            errors.pop(q["key"], None)
        elif (dep.get("required") or parent in (dep.get("required_if") or [])) and not answers.get(q["key"]):
            errors[q["key"]] = "選択してください" if q["type"] == "choice" else "入力してください"
    comment = (data.get("comment") or "").strip()[:500]
    if not survey.is_active:
        errors["survey"] = "この月度のアンケートは受け付けていません"
    if not errors and MonthlyResponse.query.filter_by(period=period, emp_id=emp_id).first():
        errors["answered"] = f"{period_label(period)}はすでに回答済みです（1人1回）"
    if errors:
        message = next(iter(errors.values()))
        if _wants_json():
            return jsonify({"success": False, "message": message, "errors": errors}), 400
        flash(message, "danger")
        return redirect(url_for("monthly.form", period=period, emp=emp_id or None))

    row = MonthlyResponse(
        period=period, emp_id=emp_id, name=name, workplace_id=workplace_id, survey_id=survey.id,
        answers_json=json.dumps(answers, ensure_ascii=False), comment=comment,
    )
    respondent = Respondent.query.filter_by(emp_id=emp_id).first()  # 同定は社員IDだけで行う
    row.respondent_id = respondent.id if respondent else None
    remember_person(emp_id, name, workplace_id)  # 照会に氏名・職場がない環境では、ここで控えに残す
    db.session.add(row)
    db.session.commit()
    session["monthly_last"] = row.id
    wants = answers.get("wants") == "ある"
    if _wants_json():
        return jsonify(
            {
                "success": True,
                "message": f"{period_label(period)}の回答を受け付けました。ご協力ありがとうございました。",
                "wants": wants,
                "register_url": url_for("survey.start", emp=emp_id),
            }
        )
    return redirect(url_for("monthly.done", period=period))


@bp_monthly.route("/done")
def done():
    period = public_period(request.args.get("period"))
    row = db.session.get(MonthlyResponse, session.get("monthly_last")) if session.get("monthly_last") else None
    return render_template("monthly/done.html", period=period, period_label=period_label(period), row=row)


@bp_monthly.route("/api/employee/<emp_id>")
def api_employee(emp_id):
    period = public_period(request.args.get("period"))
    emp_id = (emp_id or "").strip()
    info = _person(emp_id)
    answered = MonthlyResponse.query.filter_by(period=period, emp_id=emp_id).first() is not None
    if not info:
        return jsonify({"found": False, "emp_id": emp_id, "answered": answered})
    return jsonify({"found": True, "answered": answered, **info})


# ------------------------------------------------------------
# 集計（要ログイン）
# ------------------------------------------------------------
@bp_monthly.route("/results")
def results():
    rows = monthly_summary()
    surveys = MonthlySurvey.query.order_by(MonthlySurvey.period.desc()).all()
    default_period = rows[-1]["period"] if rows else (surveys[0].period if surveys else current_period())
    selected = parse_period(request.args.get("period") or default_period)
    survey = MonthlySurvey.query.filter_by(period=selected).first()
    responses = MonthlyResponse.query.filter_by(period=selected).order_by(MonthlyResponse.workplace_id, MonthlyResponse.name).all()
    # 基本設問の選択式は「設問ごとの推移」で見るので、月度の詳細には追加設問と記述設問だけ出す
    stats = [s for s in (monthly_question_stats(survey, responses) if survey else []) if s["q"]["key"] not in C.MONTHLY_CORE_KEYS or s["q"]["type"] == "text"]
    summary = next((r for r in rows if r["period"] == selected), None)
    trends = monthly_question_trends()
    trends_json = [{"key": t["q"]["key"], "labels": t["labels"], "options": [x["option"] for x in t["series"]], "percents": [x["percents"] for x in t["series"]]} for t in trends]
    chart = {
        "labels": [r["label"] for r in rows],
        "n": [r["n"] for r in rows],
        "hours_month": [round(r["hours_month"], 1) for r in rows],
        "has_rate": [round(r["has_rate"], 1) for r in rows],
        "decreased_rate": [round(r["decreased_rate"], 1) for r in rows],
        "wants_rate": [round(r["wants_rate"], 1) for r in rows],
        "ask_rate": [round(r["ask_rate"], 1) for r in rows],
        "self_rate": [round(r["self_rate"], 1) for r in rows],
    }
    return render_template(
        "monthly/results.html",
        rows=rows,
        surveys=surveys,
        survey=survey,
        selected=selected,
        selected_label=period_label(selected),
        summary=summary,
        responses=responses,
        stats=stats,
        comments=[r for r in responses if (r.comment or "").strip()],
        trends=trends,
        trends_json=json.dumps(trends_json, ensure_ascii=False),
        wp_trends=monthly_workplace_trends(),
        chart_json=json.dumps(chart, ensure_ascii=False),
        choice_keys=[q["key"] for q in (survey.questions if survey else []) if q["type"] == "choice"],
    )


@bp_monthly.route("/export.csv")
def export_csv():
    surveys = {s.period: s for s in MonthlySurvey.query.all()}
    keys, texts = [], {}
    for s in sorted(surveys.values(), key=lambda s: s.period):
        for q in s.questions:
            if q["key"] not in texts:
                keys.append(q["key"])
                texts[q["key"]] = q["text"]
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["月度", "社員ID", "氏名", "職場"] + [texts[k] for k in keys] + ["コメント", "回答日時"])
    for r in MonthlyResponse.query.order_by(MonthlyResponse.period, MonthlyResponse.workplace_id, MonthlyResponse.name).all():
        a = r.answers
        w.writerow(
            [r.period, r.emp_id, r.name, r.workplace.name if r.workplace else ""]
            + [a.get(k, "") for k in keys]
            + [r.comment or "", r.created_at.strftime("%Y-%m-%d %H:%M") if r.created_at else ""]
        )
    return Response(
        buf.getvalue().encode("utf-8-sig"),
        mimetype="text/csv",
        headers={"Content-Disposition": "attachment; filename=monthly_{}.csv".format(datetime.now().strftime("%Y%m%d"))},
    )


# ------------------------------------------------------------
# 設問の編集（要ログイン）
# ------------------------------------------------------------
@bp_monthly.route("/admin")
def survey_admin():
    period = parse_period(request.args.get("period") or current_period())
    survey = get_or_create_survey(period)
    surveys = MonthlySurvey.query.order_by(MonthlySurvey.period.desc()).all()
    counts = dict(db.session.query(MonthlyResponse.period, db.func.count(MonthlyResponse.id)).group_by(MonthlyResponse.period).all())
    return render_template(
        "monthly/admin.html",
        period=period,
        period_label=period_label(period),
        survey=survey,
        questions_json=json.dumps(survey.questions, ensure_ascii=False),
        core_keys=sorted(C.MONTHLY_CORE_KEYS),
        surveys=surveys,
        counts=counts,
        response_count=counts.get(period, 0),
        prev_period=shift_period(period, -1),
        current_period=current_period(),
    )


@bp_monthly.route("/admin/questions")
def survey_admin_questions():
    """コピー用: 指定月度の設問（period=default なら基本設問）"""
    which = (request.args.get("period") or "default").strip()
    if which == "default":
        return jsonify({"questions": default_questions(), "label": "基本設問"})
    survey = MonthlySurvey.query.filter_by(period=parse_period(which)).first()
    if survey is None:
        return jsonify({"questions": None, "message": f"{period_label(parse_period(which))}の設問はありません"})
    return jsonify({"questions": survey.questions, "label": survey.period_label})


def _clean_questions(items):
    """管理画面から送られた設問を検証して整える。問題があれば (None, メッセージ)

    基本設問は文面だけ変えられる。形式・選択肢・必須は定義どおりに戻し、外すこともできない（推移を途切れさせないため）。
    """
    core = {q["key"]: q for q in C.MONTHLY_CORE_QUESTIONS}
    cleaned, seen = [], set()
    for i, q in enumerate(items or [], start=1):
        if not isinstance(q, dict):
            return None, f"設問{i}の形式が不正です"
        text = str(q.get("text") or "").strip()
        if not text:
            return None, f"設問{i}の文章を入力してください"
        key = str(q.get("key") or "").strip()
        if key in core:
            if key in seen:
                return None, f"設問{i}のキーが重複しています"
            seen.add(key)
            base = core[key]
            item = {"key": key, "type": base["type"], "text": text[:200], "required": bool(base.get("required"))}
            if base["type"] == "choice":
                item["options"] = list(base["options"])
            else:
                item["placeholder"] = str(q.get("placeholder") or base.get("placeholder") or "").strip()[:100]
            if base.get("depends"):
                item["depends"] = dict(base["depends"])
            cleaned.append(item)
            continue
        qtype = q.get("type") if q.get("type") in QUESTION_TYPES else "choice"
        if not re.match(r"^[A-Za-z0-9_]{1,40}$", key):
            key = "q_" + secrets.token_hex(3)
        if key in seen:
            return None, f"設問{i}のキーが重複しています"
        seen.add(key)
        item = {"key": key, "type": qtype, "text": text[:200], "required": bool(q.get("required"))}
        if qtype == "choice":
            options = [str(o).strip()[:60] for o in (q.get("options") or []) if str(o).strip()]
            if len(options) < 2 or len(options) > 7:
                return None, f"設問{i}「{text[:20]}」の選択肢は2〜7個にしてください"
            if len(set(options)) != len(options):
                return None, f"設問{i}「{text[:20]}」の選択肢が重複しています"
            item["options"] = options
        else:
            item["placeholder"] = str(q.get("placeholder") or "").strip()[:100]
        cleaned.append(item)
    missing = [q["text"] for q in C.MONTHLY_CORE_QUESTIONS if q["key"] not in seen]
    if missing:
        return None, "基本設問は外せません: " + "、".join(missing)
    return cleaned, None


@bp_monthly.post("/admin/save")
def survey_admin_save():
    data = request.get_json(silent=True) or {}
    period = parse_period(data.get("period"))
    questions, error = _clean_questions(data.get("questions"))
    if error:
        return jsonify({"success": False, "message": error}), 400
    survey = get_or_create_survey(period)
    survey.set_questions(questions)
    survey.is_active = bool(data.get("is_active", True))
    db.session.commit()
    return jsonify({"success": True, "message": f"{period_label(period)}の設問を保存しました（{len(questions)}問）。", "count": len(questions)})


@bp_monthly.post("/admin/<int:survey_id>/delete")
def survey_admin_delete(survey_id):
    survey = db.session.get(MonthlySurvey, survey_id)
    if survey is None:
        abort(404)
    if MonthlyResponse.query.filter_by(period=survey.period).count():
        flash("回答がある月度の設問は削除できません。", "warning")
    else:
        db.session.delete(survey)
        db.session.commit()
        flash(f"{survey.period_label}の設問セットを削除しました。", "success")
    return redirect(url_for("monthly.survey_admin", period=current_period()))


# ==============================================================================
# 画面: 活動状況
# ==============================================================================
bp_board = Blueprint("board", __name__, url_prefix="/board")


@bp_board.route("/")
def index():
    workplace_id = request.args.get("workplace_id", type=int)
    q = Task.query
    if workplace_id:
        q = q.filter(Task.workplace_id == workplace_id)
    tasks = q.all()
    workplaces = Workplace.query.order_by(Workplace.sort, Workplace.id).all()
    selected = next((w for w in workplaces if w.id == workplace_id), None)
    return render_template(
        "board/index.html",
        b=board_data(tasks),
        hl=highlights(tasks),
        workplaces=workplaces,
        selected_wp=selected,
        show_names=Setting.get("board_show_names", "1") != "0",
    )


def highlights(tasks):
    """依頼したくなる材料（ダッシュボードと同じ計算。職場を選んでいればその職場の業務だけ）: 累計改善案件、年間削減工数、直近の完了2件"""
    total_minutes = sum(t.monthly_minutes or 0 for t in tasks)
    saved_minutes = sum(t.saved_minutes for t in tasks)
    completed = [t for t in tasks if t.status == "完了"]
    recent = sorted(completed, key=lambda t: (t.completed_at or datetime.min, t.id), reverse=True)[:2]
    return {
        "tasks": len(tasks),
        "status_counts": Counter(t.status for t in tasks),
        "completed": len(completed),
        "saved_hours": saved_minutes / 60.0,
        "saved_annual_hours": saved_minutes * 12 / 60.0,
        "reduction_rate": (saved_minutes / total_minutes * 100.0) if total_minutes else 0.0,
        "total_hours": total_minutes / 60.0,
        "recent": recent,
    }


# ==============================================================================
# テンプレート
# ==============================================================================
# 画面のテンプレート（Jinja2）。キーは以前のファイルパス。{% extends %} / {% include %} もこのキーで引く
TEMPLATES = {
    'admin/employees.html': r'''{% extends "base.html" %}
{% block title %}社員マスタ{% endblock %}
{% block content %}
<nav aria-label="breadcrumb"><ol class="breadcrumb small mb-2">
  <li class="breadcrumb-item"><a href="{{ url_for('admin.index') }}">管理</a></li>
  <li class="breadcrumb-item active">社員マスタ・LDAP照会</li>
</ol></nav>
<h2 class="mb-1">社員マスタ・LDAP照会</h2>
<p class="text-muted">アンケートの社員ID照会に使います。LDAP API が設定されていれば照会結果が自動でここに保存されます。</p>

<div class="row g-3 mb-3">
  <div class="col-lg-6">
    <div class="card h-100"><div class="card-body">
      <h5 class="card-title">LDAP 連携の状態</h5>
      <p class="mb-2">{{ ldap_status }}</p>
      <p class="small text-muted mb-0">社内 LDAP API サービスに合わせて、<code>config.py</code> に <code>LDAP_API_URL</code>（例: <code>http://ldap-api.local/employees/{id}</code>）と、レスポンスの項目名 <code>LDAP_API_FIELD_NAME</code> / <code>LDAP_API_FIELD_DEPT</code> / <code>LDAP_API_FIELD_MAIL</code> を設定します（README 参照）。</p>
      <form class="d-flex gap-2 mt-3" onsubmit="return false">
        <input type="text" id="test-emp" class="form-control form-control-sm" placeholder="社員IDで照会テスト">
        <button type="button" class="btn btn-sm btn-outline-primary text-nowrap" id="test-btn">照会</button>
      </form>
      <div id="test-msg" class="form-text"></div>
    </div></div>
  </div>
  <div class="col-lg-6">
    <div class="card h-100"><div class="card-body">
      <h5 class="card-title">CSV取込（社員ID, 氏名, 所属, メール）</h5>
      <form method="post" action="{{ url_for('admin.import_employees') }}" enctype="multipart/form-data" class="d-flex gap-2 align-items-center">
        <input type="file" name="file" accept=".csv,text/csv" class="form-control form-control-sm" required>
        <button class="btn btn-sm btn-primary text-nowrap">取込</button>
      </form>
      <p class="small text-muted mt-2 mb-0">1行目は見出し（社員ID／氏名／所属／メール、または emp_id／name／department／email）。同じ社員IDは上書きします。所属名に職場名が含まれていれば職場に紐づけます。文字コードは UTF-8 または Shift_JIS。</p>
      <div class="mt-2 d-flex gap-2">
        <a class="btn btn-sm btn-outline-secondary" href="{{ url_for('admin.export_employees') }}">現在の社員マスタをCSV出力</a>
      </div>
    </div></div>
  </div>
</div>

<div class="card mb-3"><div class="card-body">
  <h5 class="card-title">追加</h5>
  <form method="post" action="{{ url_for('admin.add_employee') }}" class="row g-2 align-items-end">
    <div class="col-6 col-md-2"><label class="form-label small mb-0">社員ID</label><input type="text" name="emp_id" class="form-control form-control-sm" required></div>
    <div class="col-6 col-md-2"><label class="form-label small mb-0">氏名</label><input type="text" name="name" class="form-control form-control-sm" required></div>
    <div class="col-6 col-md-2"><label class="form-label small mb-0">所属</label><input type="text" name="department" class="form-control form-control-sm"></div>
    <div class="col-6 col-md-2"><label class="form-label small mb-0">職場</label>
      <select name="workplace_id" class="form-select form-select-sm"><option value="">（所属から推定）</option>{% for w in workplaces %}<option value="{{ w.id }}">{{ w.name }}</option>{% endfor %}</select></div>
    <div class="col-8 col-md-3"><label class="form-label small mb-0">メール</label><input type="email" name="email" class="form-control form-control-sm"></div>
    <div class="col-4 col-md-1"><button class="btn btn-sm btn-primary w-100">追加</button></div>
  </form>
</div></div>

<div class="card"><div class="card-body">
  <h5 class="card-title">社員一覧（{{ employees|length }} 人、回答率の分母 {{ active_count }} 人）</h5>
  <div class="table-responsive">
  <table class="table table-sm align-middle mb-0">
    <thead><tr><th>社員ID</th><th>氏名</th><th>所属</th><th>職場</th><th>メール</th><th>元</th><th>対象</th><th></th></tr></thead>
    <tbody>
    {% for e in employees %}
    <tr class="{% if not e.active %}text-muted{% endif %}">
      <td>{{ e.emp_id }}</td><td>{{ e.name }}</td><td class="small">{{ e.department }}</td><td class="small">{{ e.workplace.name if e.workplace else '' }}</td>
      <td class="small">{{ e.email }}</td><td class="small">{{ e.source }}</td>
      <td>{% if e.active %}<span class="badge text-bg-success">対象</span>{% else %}<span class="badge text-bg-secondary">除外</span>{% endif %}</td>
      <td class="text-nowrap">
        <form method="post" action="{{ url_for('admin.toggle_employee', employee_id=e.id) }}" class="d-inline"><button class="btn btn-sm btn-outline-secondary">{{ '除外' if e.active else '対象に戻す' }}</button></form>
        <form method="post" action="{{ url_for('admin.delete_employee', employee_id=e.id) }}" class="d-inline" data-confirm="{{ e.name }} を削除しますか？"><button class="btn btn-sm btn-outline-danger">削除</button></form>
      </td>
    </tr>
    {% else %}
    <tr><td colspan="8" class="text-muted">社員が登録されていません。CSV取込か追加フォームから登録してください。</td></tr>
    {% endfor %}
    </tbody>
  </table>
  </div>
</div></div>
{% endblock %}
{% block scripts %}
<script>
document.getElementById("test-btn").addEventListener("click", async function () {
  var id = document.getElementById("test-emp").value.trim();
  var msg = document.getElementById("test-msg");
  if (!id) return;
  msg.textContent = "照会中…";
  try {
    var r = await fetch("{{ url_for('monthly.api_employee', emp_id='') }}" + encodeURIComponent(id));
    var j = await r.json();
    msg.textContent = j.found ? ("見つかりました: " + j.name + "（" + (j.department || "") + "／" + (j.email || "メールなし") + "／" + j.source + "）") : "見つかりませんでした";
  } catch (e) { msg.textContent = "照会に失敗しました"; }
});
</script>
{% endblock %}
''',
    'admin/index.html': r'''{% extends "base.html" %}
{% block title %}管理{% endblock %}
{% block content %}
<h2 class="mb-3">管理</h2>

<div class="row g-3">
  <div class="col-lg-6">
    <form method="post" action="{{ url_for('admin.save_settings') }}" class="card mb-3"><div class="card-body">
      <h5 class="card-title">プロジェクト設定</h5>
      <div class="mb-2"><label class="form-label small mb-0">ポータル名</label><input type="text" name="project_name" class="form-control form-control-sm" value="{{ settings.project_name }}"></div>
      <div class="mb-2"><label class="form-label small mb-0">キャッチコピー</label><input type="text" name="catchphrase" class="form-control form-control-sm" value="{{ settings.catchphrase }}"></div>
      <div class="mb-2"><label class="form-label small mb-0">サブメッセージ</label><input type="text" name="sub_message" class="form-control form-control-sm" value="{{ settings.sub_message }}"></div>
      <div class="row g-2">
        <div class="col-6 mb-2"><label class="form-label small mb-0">アンケート対象者数（回答率の分母）</label><input type="number" name="target_respondents" class="form-control form-control-sm" min="0" value="{{ settings.target_respondents }}"></div>
      </div>
      <div class="mb-2"><label class="form-label small mb-0">ゼロ化の目標期限</label><input type="text" name="deadline" class="form-control form-control-sm" value="{{ settings.deadline }}"></div>
      <div class="form-check mb-2">
        <input class="form-check-input" type="checkbox" name="board_show_names" value="1" id="board-show-names" {% if settings.board_show_names != '0' %}checked{% endif %}>
        <label class="form-check-label small" for="board-show-names">活動状況に依頼者の氏名を表示する（外すと職場名のみ）</label>
      </div>
      <button class="btn btn-primary btn-sm">保存</button>
    </div></form>

    <div class="card mb-3"><div class="card-body">
      <h5 class="card-title">ログイン</h5>
      <p class="small mb-0">{{ auth_status }}</p>
    </div></div>

    <div class="card mb-3"><div class="card-body">
      <h5 class="card-title">社員マスタ・月次アンケート</h5>
      <p class="small mb-1">社員 {{ employee_count }} 人登録。{{ ldap_status }}</p>
      <div class="d-flex flex-wrap gap-2">
        <a class="btn btn-sm btn-outline-primary" href="{{ url_for('admin.employees') }}">社員マスタ・LDAP照会</a>
        <a class="btn btn-sm btn-outline-primary" href="{{ url_for('monthly.survey_admin') }}">月次の設問を編集</a>
        <a class="btn btn-sm btn-outline-secondary" href="{{ url_for('monthly.results') }}">月次集計</a>
      </div>
    </div></div>

    <div class="card mb-3"><div class="card-body">
      <h5 class="card-title">AI連携</h5>
      <p class="mb-1">{{ ai_status }}</p>
      <p class="small text-muted mb-0">有効にするには、プロジェクト直下の <code>config.py</code> の <code>ANTHROPIC_API_KEY</code> にキーを書いて再起動してください。未設定の場合はルールベースで同じ形式の分析を行います。</p>
    </div></div>

    <div class="card mb-3"><div class="card-body">
      <h5 class="card-title">データ</h5>
      <p class="small mb-2">回答者 {{ respondent_count }} 人／業務 {{ task_count }} 件（うちデモ {{ demo_count }} 件）</p>
      <div class="d-flex flex-wrap gap-2">
        <form method="post" action="{{ url_for('admin.load_demo') }}" data-confirm="デモデータを投入します（既存のデモデータは入れ替えます）。よろしいですか？">
          <button class="btn btn-sm btn-outline-primary">デモデータを投入</button>
        </form>
        <form method="post" action="{{ url_for('admin.remove_demo') }}" data-confirm="デモデータを削除します。よろしいですか？">
          <button class="btn btn-sm btn-outline-danger" {% if demo_count == 0 %}disabled{% endif %}>デモデータを削除</button>
        </form>
        <a class="btn btn-sm btn-outline-secondary" href="{{ url_for('cases.export_csv') }}">業務一覧CSV</a>
        <a class="btn btn-sm btn-outline-secondary" href="{{ url_for('admin.export_respondents') }}">回答者CSV</a>
      </div>
      <p class="small text-muted mt-2 mb-0">データは <code>instance/portal.sqlite3</code> に保存されます。バックアップはこのファイルをコピーしてください。</p>
    </div></div>
  </div>

  <div class="col-lg-6">
    <div class="card mb-3"><div class="card-body">
      <h5 class="card-title">職場マスタ</h5>
      <table class="table table-sm mb-2">
        <tbody>
        {% for w in workplaces %}
        <tr>
          <td>{{ w.name }}</td>
          <td class="small text-muted">回答者 {{ w.respondents|length }} 人</td>
          <td class="text-end">
            <form method="post" action="{{ url_for('admin.delete_workplace', workplace_id=w.id) }}" data-confirm="「{{ w.name }}」を削除しますか？">
              <button class="btn btn-sm btn-outline-danger" {% if w.respondents %}disabled{% endif %}>削除</button>
            </form>
          </td>
        </tr>
        {% endfor %}
        </tbody>
      </table>
      <form method="post" action="{{ url_for('admin.add_workplace') }}" class="d-flex gap-2">
        <input type="text" name="name" class="form-control form-control-sm" placeholder="職場名" required>
        <button class="btn btn-sm btn-primary text-nowrap">追加</button>
      </form>
    </div></div>

    <div class="card"><div class="card-body">
      <h5 class="card-title">業務領域マスタ（業務体系の大分類）</h5>
      <p class="small text-muted">業務体系（大分類・中分類）の確定版に合わせて名称・説明を更新してください。業務が紐づいている領域は削除できません。</p>
      {% for d in domains %}
      <form method="post" action="{{ url_for('admin.update_domain', domain_id=d.id) }}" class="row g-1 align-items-center mb-1">
        <div class="col-1"><input type="number" name="sort" class="form-control form-control-sm" value="{{ d.sort }}" title="並び順"></div>
        <div class="col-4"><input type="text" name="name" class="form-control form-control-sm" value="{{ d.name }}"></div>
        <div class="col-5"><input type="text" name="description" class="form-control form-control-sm" value="{{ d.description }}" placeholder="説明"></div>
        <div class="col-2 d-flex gap-1">
          <button class="btn btn-sm btn-outline-primary">保存</button>
          <button class="btn btn-sm btn-outline-danger" formaction="{{ url_for('admin.delete_domain', domain_id=d.id) }}" {% if d.tasks or d.children %}disabled{% endif %} onclick="return confirm('「{{ d.name }}」を削除しますか？')">削除</button>
        </div>
      </form>
      {% endfor %}
      <form method="post" action="{{ url_for('admin.add_domain') }}" class="row g-1 align-items-center mt-2">
        <div class="col-5"><input type="text" name="name" class="form-control form-control-sm" placeholder="新しい領域名" required></div>
        <div class="col-5"><input type="text" name="description" class="form-control form-control-sm" placeholder="説明"></div>
        <div class="col-2"><button class="btn btn-sm btn-primary w-100">追加</button></div>
      </form>
    </div></div>
  </div>
</div>
{% endblock %}
''',
    'auth/login.html': r'''{% extends "base.html" %}
{% block title %}ログイン{% endblock %}
{% block content %}
<div class="login-wrap">
  <div class="card login-card"><div class="card-body">
    <div class="login-mark"><i class="bi bi-person-lock"></i></div>
    <h1 class="login-title">ログイン</h1>
    <p class="login-sub">推進メンバー用です。アンケートの回答と活動状況はログインなしで使えます。</p>
    {% if error %}<div class="alert alert-danger py-2 small">{{ error }}</div>{% endif %}
    <form method="post" action="{{ url_for('auth.login') }}">
      <input type="hidden" name="next" value="{{ next }}">
      <div class="mb-3">
        <label class="form-label" for="emp_id">ID</label>
        <input type="text" name="emp_id" id="emp_id" class="form-control form-control-lg" value="{{ emp_id }}" autocomplete="username" autofocus required>
      </div>
      <div class="mb-3">
        <label class="form-label" for="password">パスワード</label>
        <input type="password" name="password" id="password" class="form-control form-control-lg" autocomplete="current-password" required>
      </div>
      <button class="btn btn-primary btn-lg w-100">ログイン</button>
    </form>
  </div></div>
  <p class="text-center small mt-3"><a href="{{ url_for('board.index') }}">活動状況へ戻る</a></p>
</div>
{% endblock %}
''',
    'base.html': r'''<!doctype html>
<html lang="ja">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <meta name="theme-color" content="#f5f5f7">
  <title>{% block title %}{% endblock %} | {{ settings.project_name }}</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400..700&family=Noto+Sans+JP:wght@400..700&display=swap" rel="stylesheet">
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
  <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" rel="stylesheet">
  <link rel="stylesheet" href="{{ url_for('static', filename='style.css') }}">
</head>
<body>
<header class="app-header sticky-top">
<nav class="navbar navbar-expand-xl app-nav">
  <div class="container">
    <a class="navbar-brand" href="{{ url_for('main.index' if current_user else 'board.index') }}"><span class="brand-mark">0</span>PC定型業務ゼロ化プロジェクト</a>
    <button class="navbar-toggler" type="button" data-bs-toggle="collapse" data-bs-target="#nav" aria-label="メニュー">
      <span class="navbar-toggler-icon"></span>
    </button>
    <div class="collapse navbar-collapse" id="nav">
      {% set bp = request.blueprint %}
      <ul class="navbar-nav me-auto">
        <li class="nav-item"><a class="nav-link {% if bp == 'board' %}active{% endif %}" href="{{ url_for('board.index') }}">活動状況</a></li>
        <li class="nav-item"><a class="nav-link {% if bp == 'survey' %}active{% endif %}" href="{{ url_for('survey.index') }}">自動化依頼</a></li>
        <li class="nav-item"><a class="nav-link {% if request.endpoint in ('monthly.form', 'monthly.done') %}active{% endif %}" href="{{ url_for('monthly.form') }}">月次アンケート</a></li>
        {% if current_user %}
        <li class="nav-sep" aria-hidden="true"></li>
        <li class="nav-lock" title="管理者メニュー"><i class="bi bi-lock-fill"></i></li>
        <li class="nav-item"><a class="nav-link nav-link-admin {% if request.endpoint == 'main.index' %}active{% endif %}" href="{{ url_for('main.index') }}">ダッシュボード</a></li>
        <li class="nav-item"><a class="nav-link nav-link-admin {% if bp == 'cases' %}active{% endif %}" href="{{ url_for('cases.index') }}">改善案件</a></li>
        <li class="nav-item"><a class="nav-link nav-link-admin {% if bp == 'bizmap' %}active{% endif %}" href="{{ url_for('bizmap.index') }}">業務マップ</a></li>
        <li class="nav-item"><a class="nav-link nav-link-admin {% if bp == 'people' %}active{% endif %}" href="{{ url_for('people.index') }}">DX人財情報</a></li>
        <li class="nav-item"><a class="nav-link nav-link-admin {% if request.endpoint in ('monthly.results', 'monthly.survey_admin') %}active{% endif %}" href="{{ url_for('monthly.results') }}">月次集計</a></li>
        <li class="nav-item"><a class="nav-link nav-link-admin {% if bp == 'admin' %}active{% endif %}" href="{{ url_for('admin.index') }}">管理</a></li>
        {% endif %}
      </ul>
      {% if current_user %}
      <form method="post" action="{{ url_for('auth.logout') }}" class="d-inline"><button class="nav-admin" title="{{ current_user.name }}としてログイン中">ログアウト</button></form>
      {% else %}
      <a class="nav-admin" href="{{ url_for('auth.login') }}">管理者メニュー</a>
      {% endif %}
    </div>
  </div>
</nav>
</header>

<main class="container py-4">
  {% with messages = get_flashed_messages(with_categories=true) %}
    {% for category, message in messages %}
      <div class="alert alert-{{ 'secondary' if category == 'message' else category }} alert-dismissible fade show" role="alert">
        {{ message }}
        <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="閉じる"></button>
      </div>
    {% endfor %}
  {% endwith %}
  {% block content %}{% endblock %}
</main>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script>
<script src="{{ url_for('static', filename='app.js') }}"></script>
{% block scripts %}{% endblock %}
</body>
</html>
''',
    'bizmap/domain.html': r'''{% extends "base.html" %}
{% block title %}{{ d.name }}{% endblock %}
{% block content %}
<nav aria-label="breadcrumb"><ol class="breadcrumb small mb-2">
  <li class="breadcrumb-item"><a href="{{ url_for('bizmap.index') }}">業務マップ</a></li>
  <li class="breadcrumb-item active">{{ d.name }}</li>
</ol></nav>
<h2 class="mb-1">{{ d.name }}</h2>
<p class="text-muted">{{ d.description }}</p>

{% if row %}
<div class="row g-3 mb-3">
  <div class="col-6 col-md-2"><div class="card kpi-card"><div class="card-body py-2"><div class="kpi-label">登録</div><div class="kpi-value fs-3">{{ row.count }}</div></div></div></div>
  <div class="col-6 col-md-2"><div class="card kpi-card"><div class="card-body py-2"><div class="kpi-label">案件候補</div><div class="kpi-value fs-3">{{ row.candidates }}</div></div></div></div>
  <div class="col-6 col-md-2"><div class="card kpi-card"><div class="card-body py-2"><div class="kpi-label">月間時間</div><div class="kpi-value fs-3">{{ row.minutes|hours }}<span class="fs-6"> h</span></div></div></div></div>
  <div class="col-6 col-md-2"><div class="card kpi-card"><div class="card-body py-2"><div class="kpi-label">充足率</div><div class="kpi-value fs-3">{{ '%.0f'|format(row.fill_rate) }}<span class="fs-6">%</span></div></div></div></div>
  <div class="col-6 col-md-2"><div class="card kpi-card"><div class="card-body py-2"><div class="kpi-label">確認済み</div><div class="kpi-value fs-3">{{ '%.0f'|format(row.verified_rate) }}<span class="fs-6">%</span></div></div></div></div>
  <div class="col-6 col-md-2"><div class="card kpi-card heat-{{ row.level }}"><div class="card-body py-2"><div class="kpi-label">分析状態</div><div class="kpi-value fs-4">{{ row.state }}</div><div class="kpi-sub">業務把握: {{ row.grasp }}</div></div></div></div>
</div>
{% if row.level <= 1 %}
<div class="alert alert-warning">
  この領域は業務情報が不足しています。登録を促すか、担当者へのヒアリングで業務の実態（目的・アウトプット・頻度・時間）を把握してください。
</div>
{% endif %}
{% endif %}

{% if info_flow %}
<div class="card mb-3"><div class="card-body">
  <div class="d-flex justify-content-between align-items-start flex-wrap gap-2">
    <div>
      <h5 class="mb-1">この領域の情報の流れ</h5>
      <p class="small text-muted mb-0">情報源 → 業務 → アウトプット → 利用者。業務の箱をクリックすると案件詳細を開きます。</p>
    </div>
    <div class="btn-group btn-group-sm flow-toolbar" data-target="info" role="group">
      <button type="button" class="btn btn-outline-secondary" data-zoom="out">－</button>
      <button type="button" class="btn btn-outline-secondary" data-zoom="reset">100%</button>
      <button type="button" class="btn btn-outline-secondary" data-zoom="in">＋</button>
    </div>
  </div>
  <div class="flow-canvas"><div class="flow-inner" id="flow-info" data-def="def-info"><span class="text-muted small">描画中…</span></div></div>
  <script type="text/plain" id="def-info">{{ info_flow|safe }}</script>
</div></div>
{% endif %}

<div class="card mb-3"><div class="card-body">
  <h5 class="card-title">この領域の業務（月間時間順）</h5>
  <div class="table-responsive">
  <table class="table table-sm align-middle mb-0">
    <thead><tr><th>業務名</th><th>職場／回答者</th><th>目的</th><th>アウトプット</th><th class="text-end">月間</th><th>状態</th><th>確認</th></tr></thead>
    <tbody>
    {% for t in tasks %}
    <tr>
      <td><a href="{{ url_for('cases.detail', task_id=t.id) }}">{{ t.name }}</a></td>
      <td class="small">{{ t.workplace.name if t.workplace else '' }}／{{ t.respondent.name if t.respondent else '' }}</td>
      <td class="small">{{ t.purpose }}</td>
      <td class="small">{{ t.output }}</td>
      <td class="text-end">{{ t.monthly_minutes|hours }} h</td>
      <td><span class="badge text-bg-{{ {'受付':'secondary','対応中':'primary','完了':'success','保留':'warning'}.get(t.status, 'secondary') }}">{{ t.status }}</span></td>
      <td>{% if t.verified %}<span class="badge text-bg-success">済</span>{% else %}<span class="badge text-bg-light text-muted">未</span>{% endif %}</td>
    </tr>
    {% else %}
    <tr><td colspan="7" class="text-muted">まだ登録がありません</td></tr>
    {% endfor %}
    </tbody>
  </table>
  </div>
</div></div>

<div class="row g-3">
  <div class="col-lg-6">
    <div class="card h-100"><div class="card-body">
      <h5 class="card-title">ヒアリングが必要な業務</h5>
      <p class="small text-muted">必須項目が未入力、または担当者確認が済んでいない業務です。</p>
      <ul class="mb-0 small">
        {% for t in missing %}<li><a href="{{ url_for('cases.detail', task_id=t.id) }}">{{ t.name }}</a> — 必須項目に未入力</li>{% endfor %}
        {% for t in unverified %}{% if t not in missing %}<li><a href="{{ url_for('cases.detail', task_id=t.id) }}">{{ t.name }}</a> — 担当者未確認</li>{% endif %}{% endfor %}
        {% if not missing and not unverified %}<li class="text-muted">すべて確認済みです</li>{% endif %}
      </ul>
    </div></div>
  </div>
  <div class="col-lg-6">
    <div class="card h-100"><div class="card-body">
      <h5 class="card-title">業務のつながり</h5>
      {% if flows %}
      <ul class="mb-0 small">
        {% for t in flows %}<li>{{ t.prev_task or '－' }} → <a href="{{ url_for('cases.detail', task_id=t.id) }}">{{ t.name }}</a> → {{ t.next_task or '－' }}</li>{% endfor %}
      </ul>
      {% else %}
      <p class="text-muted small mb-0">前後の業務が登録されるとここに表示されます。</p>
      {% endif %}
    </div></div>
  </div>
</div>
{% endblock %}
{% block scripts %}
<script src="https://cdn.jsdelivr.net/npm/mermaid@11.4.1/dist/mermaid.min.js"></script>
<script src="{{ url_for('static', filename='flows.js') }}"></script>
{% endblock %}
''',
    'bizmap/edit.html': r'''{% extends "base.html" %}
{% block title %}{{ kind_label }}の編集{% endblock %}
{% block content %}
<div id="fe-app" class="fe-wrap"
     data-kind="{{ kind }}" data-workplace="{{ workplace_id or '' }}" data-domain="{{ domain_id or '' }}"
     data-api="{{ url_for('bizmap.api_graph_get') }}" data-back="{{ url_for('bizmap.index', workplace_id=workplace_id, domain_id=domain_id) }}">
  <div class="fe-toolbar">
    <div class="fe-title">
      <strong>{{ kind_label }}</strong> <span class="text-muted">— {{ scope_label }}</span>
      <span id="fe-status" class="fe-status text-muted"></span>
    </div>
    <div class="fe-buttons">
      <button type="button" class="btn btn-sm btn-success" id="fe-save" title="Ctrl+S">💾 保存</button>
      <button type="button" class="btn btn-sm btn-outline-secondary" id="fe-undo" title="Ctrl+Z" disabled>↩ 元に戻す</button>
      <button type="button" class="btn btn-sm btn-outline-secondary" id="fe-redo" title="Ctrl+Y" disabled>↪ やり直す</button>
      <span class="fe-sep"></span>
      <button type="button" class="btn btn-sm btn-outline-primary" id="fe-add">➕ ノード追加</button>
      <button type="button" class="btn btn-sm btn-outline-primary" id="fe-layout">自動レイアウト</button>
      <button type="button" class="btn btn-sm btn-outline-primary" id="fe-fit">全体表示</button>
      <div class="btn-group btn-group-sm" role="group">
        <button type="button" class="btn btn-outline-secondary" id="fe-zoom-out">－</button>
        <button type="button" class="btn btn-outline-secondary" id="fe-zoom-label">100%</button>
        <button type="button" class="btn btn-outline-secondary" id="fe-zoom-in">＋</button>
      </div>
      <span class="fe-sep"></span>
      <button type="button" class="btn btn-sm btn-outline-secondary" id="fe-reload" title="登録された業務を取り込み直します（保存した手修正は残ります）">再生成</button>
      <button type="button" class="btn btn-sm btn-outline-secondary" id="fe-restore" title="削除（非表示）にした自動生成のノード・矢印を戻します">非表示を戻す</button>
      <button type="button" class="btn btn-sm btn-outline-danger" id="fe-reset" title="手修正をすべて捨てて自動生成に戻します">自動生成に戻す</button>
      <button type="button" class="btn btn-sm btn-outline-success" id="fe-writeback" title="手で引いた矢印や直した業務名を、業務データ（前後の業務・情報の受け取り元・利用者・関わる人）に反映します">業務データに反映</button>
      <a class="btn btn-sm btn-outline-dark" href="{{ url_for('bizmap.index', workplace_id=workplace_id, domain_id=domain_id) }}">閲覧へ戻る</a>
    </div>
  </div>
  <div class="fe-hint">
    操作: ドラッグで移動 ／ 文字をダブルクリックで名前を編集 ／ 🔗 で矢印を引く（相手をクリック） ／ 矢印をクリックで向き反転・ラベル・削除 ／ Shift+ドラッグで範囲選択 ／ 空白をドラッグで画面移動 ／ ホイールで拡大縮小 ／ Delete で選択を削除
    <span class="fe-legend">
      {% for cls, label in cls_labels.items() if cls != 'done' %}<span class="legend-chip fe-lg-{{ cls }}">{{ label }}</span>{% endfor %}
      <span class="fe-edge-legend"><i class="fe-line fe-line-info"></i>情報 <i class="fe-line fe-line-seq"></i>前後の業務 <i class="fe-line fe-line-manual"></i>手で追加</span>
    </span>
  </div>
  <div id="fe-mode" class="fe-mode"></div>
  <div class="fe-container" id="fe-container">
    <div class="fe-canvas" id="fe-canvas">
      <svg class="fe-svg" id="fe-svg">
        <defs>
          <marker id="fe-arrow-info" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse"><path d="M0 0L10 5L0 10z" fill="#0d6efd"/></marker>
          <marker id="fe-arrow-seq" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse"><path d="M0 0L10 5L0 10z" fill="#6c757d"/></marker>
          <marker id="fe-arrow-manual" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse"><path d="M0 0L10 5L0 10z" fill="#6f42c1"/></marker>
        </defs>
        <g id="fe-edges"></g>
        <g id="fe-temp"></g>
      </svg>
      <div id="fe-nodes"></div>
    </div>
    <div class="fe-selection" id="fe-selection"></div>
    <div class="fe-popover" id="fe-node-popover">
      <div class="fe-pop-title">ノードの設定</div>
      <div class="fe-colors">
        {% for c in ['white','red','orange','yellow','green','blue','purple','pink'] %}<span class="fe-color" data-color="{{ c }}" title="{{ c }}"></span>{% endfor %}
        <span class="fe-color fe-color-reset" data-color="" title="元の色に戻す">×</span>
      </div>
      <div class="mt-2" id="fe-cls-row">
        <label class="small mb-0">種類</label>
        <select class="form-select form-select-sm" id="fe-cls-select">
          {% for cls, label in cls_labels.items() if cls not in ('done',) %}<option value="{{ cls }}">{{ label }}</option>{% endfor %}
        </select>
      </div>
      <div class="mt-2 d-flex gap-1">
        <button type="button" class="btn btn-sm btn-outline-danger" id="fe-node-delete">削除（非表示）</button>
        <button type="button" class="btn btn-sm btn-outline-secondary" id="fe-node-close">閉じる</button>
      </div>
    </div>
    <div class="fe-popover fe-popover-wide" id="fe-wb-popover">
      <div class="fe-pop-title">業務データに反映する内容</div>
      <p class="small text-muted mb-2">手で引いた矢印・直した業務名・消した前後の矢印から、反映できる変更を取り出しました。反映すると次回の再生成から自動で描かれます。</p>
      <div id="fe-wb-list" class="fe-wb-list"></div>
      <div class="mt-2 d-flex gap-1">
        <button type="button" class="btn btn-sm btn-success" id="fe-wb-apply">選んだ変更を反映</button>
        <button type="button" class="btn btn-sm btn-outline-secondary" id="fe-wb-close">閉じる</button>
      </div>
    </div>
    <div class="fe-popover" id="fe-edge-popover">
      <div class="fe-pop-title">矢印の設定</div>
      <input type="text" class="form-control form-control-sm" id="fe-edge-label" placeholder="ラベル（例：日報、材料）">
      <div class="mt-2 d-flex flex-wrap gap-1">
        <button type="button" class="btn btn-sm btn-outline-primary" id="fe-edge-reverse">向きを反転</button>
        <button type="button" class="btn btn-sm btn-outline-danger" id="fe-edge-delete">削除</button>
        <button type="button" class="btn btn-sm btn-outline-secondary" id="fe-edge-close">閉じる</button>
      </div>
    </div>
  </div>
</div>
{% endblock %}
{% block scripts %}
<script src="https://cdn.jsdelivr.net/npm/dagre@0.8.5/dist/dagre.min.js"></script>
<script src="{{ url_for('static', filename='flowedit.js') }}"></script>
{% endblock %}
''',
    'bizmap/index.html': r'''{% extends "base.html" %}
{% block title %}業務マップ{% endblock %}
{% macro flow_panel(key, title, desc, legend, definition, empty_msg) -%}
<div class="card"><div class="card-body">
  <div class="d-flex justify-content-between align-items-start flex-wrap gap-2">
    <div>
      <h5 class="mb-1">{{ title }}
        {% if saved_at[key] %}<span class="badge text-bg-warning align-middle" title="編集画面で直した配置・名前・追加・削除を反映しています">手修正あり（{{ saved_at[key] }}）</span>{% endif %}
      </h5>
      <p class="small text-muted mb-1">{{ desc }}</p>
      <div>{{ legend|safe }}</div>
    </div>
    <div class="d-flex flex-column align-items-end gap-1">
      <a class="btn btn-sm btn-primary" href="{{ url_for('bizmap.edit', kind=key, workplace_id=args.get('workplace_id') or None, domain_id=args.get('domain_id') or None) }}">✎ この図を編集する</a>
      {% if definition %}
      <div class="btn-group btn-group-sm flow-toolbar" data-target="{{ key }}" role="group" aria-label="表示操作">
        <button type="button" class="btn btn-outline-secondary" data-zoom="out" title="縮小">－</button>
        <button type="button" class="btn btn-outline-secondary" data-zoom="reset" title="元の大きさ">100%</button>
        <button type="button" class="btn btn-outline-secondary" data-zoom="in" title="拡大">＋</button>
      </div>
      {% endif %}
    </div>
  </div>
  {% if definition %}
  <div class="flow-canvas"><div class="flow-inner" id="flow-{{ key }}" data-def="def-{{ key }}"><span class="text-muted small">描画中…</span></div></div>
  <script type="text/plain" id="def-{{ key }}">{{ definition|safe }}</script>
  <p class="small text-muted mt-2 mb-0">業務の箱をクリックすると案件詳細を開きます。</p>
  {% else %}
  <div class="alert alert-light border mt-3 mb-0">{{ empty_msg }}</div>
  {% endif %}
</div></div>
{%- endmacro %}
{% block content %}
<h2 class="mb-1">業務体系・業務マップ</h2>
<p class="text-muted">登録された業務から、システム・帳票・人などの情報源 → 業務 → アウトプット → 受け渡し先を1枚の図に自動生成します。案件が増えるほど製造部の業務全体像が描かれていきます。間違いは「この図を編集する」で手で直せ、直した内容は保存され、新しい業務を取り込んで再生成しても残ります。</p>

<form method="get" class="card mb-3"><div class="card-body py-2">
  <div class="row g-2 align-items-end">
    <div class="col-6 col-md-3">
      <label class="form-label small mb-0">職場</label>
      <select name="workplace_id" class="form-select form-select-sm">
        <option value="">製造部全体</option>
        {% for w in workplaces %}<option value="{{ w.id }}" {% if args.get('workplace_id')|string == w.id|string %}selected{% endif %}>{{ w.name }}</option>{% endfor %}
      </select>
    </div>
    <div class="col-6 col-md-3">
      <label class="form-label small mb-0">業務領域</label>
      <select name="domain_id" class="form-select form-select-sm">
        <option value="">すべて</option>
        {% for d in domains %}<option value="{{ d.id }}" {% if args.get('domain_id')|string == d.id|string %}selected{% endif %}>{{ d.name }}</option>{% endfor %}
      </select>
    </div>
    <div class="col-6 col-md-2"><button class="btn btn-sm btn-primary">表示</button></div>
    <div class="col-12 col-md-4 small text-muted text-md-end">
      図に載せた業務 {{ task_count }} 件（関わる人・情報の受け取り元の入力あり {{ flow_info_count }} 件）
    </div>
  </div>
</div></form>
{% if truncated %}
<div class="alert alert-warning py-2 small">業務が {{ max_tasks }} 件を超えるため、先頭 {{ max_tasks }} 件だけを描いています。職場や領域で絞り込んでください。</div>
{% endif %}

<div class="bizmap-sections">
  <div id="tab-info">
    {{ flow_panel('info', '業務マップ',
       'システム・帳票・人などの情報源 → 業務 → アウトプット → 受け渡し先（職場内・上司・他部署・社外）。職場ごとの枠に業務を置き、点線は前後の業務やアウトプットの受け渡しです。',
       '<span class="legend-chip lg-system">システム</span><span class="legend-chip lg-source">帳票・人・メールなどの情報源</span><span class="legend-chip lg-task">業務</span><span class="legend-chip lg-done">完了した業務</span><span class="legend-chip lg-output">アウトプット</span><span class="legend-chip lg-dupout">同じアウトプット（統合候補）</span><span class="legend-chip lg-wp">職場内の利用者</span><span class="legend-chip lg-ext">他部署・社外</span><span class="legend-chip lg-warn">利用者不明</span>',
       flows.info, '業務が登録されると図が描かれます。') }}

    <div class="row g-3 mt-1">
      <div class="col-lg-7">
        <div class="card h-100"><div class="card-body">
          <h5 class="card-title">システム・情報源と業務のつながり</h5>
          <p class="small text-muted">同じシステムやファイルを何件の業務が読んでいるか。件数と時間が多いところほど、そこを起点にした自動化の効果が大きくなります。</p>
          <div class="table-responsive">
          <table class="table table-sm mb-0">
            <thead><tr><th>システム・情報源</th><th class="text-end">業務</th><th class="text-end">月間</th><th>業務名</th></tr></thead>
            <tbody>
            {% for r in source_rows %}
            <tr>
              <td>{% if r.cls == 'system' %}<span class="legend-chip lg-system">{{ r.name }}</span>{% else %}{{ r.name }}{% endif %}</td>
              <td class="text-end">{{ r.tasks|length }}</td>
              <td class="text-end">{{ r.minutes|hours }} h</td>
              <td class="small">{% for t in r.tasks %}<a href="{{ url_for('cases.detail', task_id=t.id) }}">{{ t.name }}</a>{% if not loop.last %}、{% endif %}{% endfor %}</td>
            </tr>
            {% else %}
            <tr><td colspan="4" class="text-muted">データなし</td></tr>
            {% endfor %}
            </tbody>
          </table>
          </div>
        </div></div>
      </div>
      <div class="col-lg-5">
        <div class="card h-100"><div class="card-body">
          <h5 class="card-title">職場・部署の間の受け渡し</h5>
          <p class="small text-muted">誰から受け取り、誰に渡しているか。「情報」は情報の受け取り元とアウトプットの利用者から、「関与」は関わる人・部署から数えています。</p>
          <div class="table-responsive">
          <table class="table table-sm mb-0">
            <thead><tr><th>から</th><th>へ</th><th class="text-end">情報</th><th class="text-end">関与</th></tr></thead>
            <tbody>
            {% for r in handoff_rows %}
            <tr><td>{{ r.src }}</td><td>{{ r.dst }}</td><td class="text-end">{{ r.info or '' }}</td><td class="text-end">{{ r.involve or '' }}</td></tr>
            {% else %}
            <tr><td colspan="4" class="text-muted">データなし</td></tr>
            {% endfor %}
            </tbody>
          </table>
          </div>
        </div></div>
      </div>
    </div>
  </div>

  <div id="tab-domains" class="mt-4">
    <h2 class="section-title">領域別の把握状況 <span class="section-sub">案件がない領域と、実態を把握できていない領域を区別する</span></h2>
    <div class="card"><div class="card-body">
      <h5 class="card-title">業務領域ごとの把握状況（ヒートマップ）</h5>
      <p class="small text-muted">案件がない領域と、業務実態を把握できていない領域を区別します。絞り込みに関係なく製造部全体で判定しています。</p>
      <div class="table-responsive">
      <table class="table table-sm align-middle mb-2">
        <thead><tr><th>業務領域</th><th>説明</th><th class="text-end">登録</th><th class="text-end">案件候補</th><th class="text-end">完了</th><th class="text-end">月間時間</th><th class="text-end">充足率</th><th class="text-end">確認済み</th><th>業務把握</th><th>分析状態</th></tr></thead>
        <tbody>
        {% for r in rows %}
        <tr>
          <td><a href="{{ url_for('bizmap.domain', domain_id=r.domain.id) }}">{{ r.domain.name }}</a></td>
          <td class="small text-muted">{{ r.domain.description }}</td>
          <td class="text-end">{{ r.count }}</td>
          <td class="text-end">{{ r.candidates }}</td>
          <td class="text-end">{{ r.completed }}</td>
          <td class="text-end">{{ r.minutes|hours }} h</td>
          <td class="text-end">{{ '%.0f'|format(r.fill_rate) }}%</td>
          <td class="text-end">{{ '%.0f'|format(r.verified_rate) }}%</td>
          <td class="heat-{{ r.level }}">{{ r.grasp }}</td>
          <td class="heat-{{ r.level }}">{{ r.state }}</td>
        </tr>
        {% endfor %}
        </tbody>
      </table>
      </div>
      <div class="small text-muted">
        判定: 登録件数と必須項目（目的・アウトプット・利用者・入力元・頻度・時間）の充足率、担当者確認率から判定。
        <span class="heat-3 px-1">分析済み</span> <span class="heat-2 px-1">追加調査</span> <span class="heat-1 px-1">情報不足</span> <span class="heat-0 px-1">未分析</span>
        ／ 業務領域は <a href="{{ url_for('admin.index') }}">管理</a> で業務体系の確定版に合わせて編集できます。
      </div>
    </div></div>
  </div>

  <div id="tab-dup" class="mt-4">
    <h2 class="section-title">重複・つながり <span class="section-sub">同じアウトプット・重複入力・前後の業務</span></h2>
    <div class="row g-3 mb-3">
      <div class="col-lg-6">
        <div class="card h-100"><div class="card-body">
          <h5 class="card-title">同じアウトプットを複数人が作成（統合候補）</h5>
          {% if dup_outputs %}
          <table class="table table-sm mb-0">
            <thead><tr><th>アウトプット</th><th>業務（回答者）</th><th class="text-end">月間</th></tr></thead>
            <tbody>
            {% for g in dup_outputs %}
            <tr>
              <td>{{ g.output }}</td>
              <td class="small">{% for t in g.tasks %}<a href="{{ url_for('cases.detail', task_id=t.id) }}">{{ t.name }}</a>（{{ t.respondent.name if t.respondent else '' }}）{% if not loop.last %}<br>{% endif %}{% endfor %}</td>
              <td class="text-end">{{ g.minutes|hours }} h</td>
            </tr>
            {% endfor %}
            </tbody>
          </table>
          {% else %}
          <p class="text-muted small mb-0">同じアウトプットの重複は見つかっていません。</p>
          {% endif %}
        </div></div>
      </div>
      <div class="col-lg-6">
        <div class="card h-100"><div class="card-body">
          <h5 class="card-title">「同じ情報を何度も入力する」業務（入力元別）</h5>
          {% if dup_inputs %}
          <table class="table table-sm mb-0">
            <thead><tr><th>入力元</th><th>業務</th></tr></thead>
            <tbody>
            {% for src, tasks in dup_inputs %}
            <tr>
              <td>{{ src }}<div class="small text-muted">{{ tasks|length }} 件</div></td>
              <td class="small">{% for t in tasks %}<a href="{{ url_for('cases.detail', task_id=t.id) }}">{{ t.name }}</a>{% if not loop.last %}<br>{% endif %}{% endfor %}</td>
            </tr>
            {% endfor %}
            </tbody>
          </table>
          {% else %}
          <p class="text-muted small mb-0">重複入力の指摘はありません。</p>
          {% endif %}
        </div></div>
      </div>
    </div>
    <div class="row g-3">
      <div class="col-12">
        <div class="card h-100"><div class="card-body">
          <h5 class="card-title">業務のつながり（前の業務 → 業務 → 後の業務）</h5>
          {% if flow_rows %}
          <div class="table-responsive">
          <table class="table table-sm mb-0">
            <thead><tr><th>領域</th><th>前の業務</th><th>業務</th><th>後の業務</th></tr></thead>
            <tbody>
            {% for t in flow_rows %}
            <tr>
              <td class="small">{{ t.domain.name if t.domain else '' }}</td>
              <td class="small text-muted">{{ t.prev_task or '－' }}</td>
              <td><a href="{{ url_for('cases.detail', task_id=t.id) }}">{{ t.name }}</a></td>
              <td class="small text-muted">{{ t.next_task or '－' }}</td>
            </tr>
            {% endfor %}
            </tbody>
          </table>
          </div>
          {% else %}
          <p class="text-muted small mb-0">前後の業務が登録されるとここに表示されます。</p>
          {% endif %}
        </div></div>
      </div>
    </div>
  </div>
</div>
{% endblock %}
{% block scripts %}
<script src="https://cdn.jsdelivr.net/npm/mermaid@11.4.1/dist/mermaid.min.js"></script>
<script src="{{ url_for('static', filename='flows.js') }}"></script>
{% endblock %}
''',
    'board/_board.html': r'''{# 活動状況の本体 #}
{% from "board/_macros.html" import row with context %}

{% set n_done, n_active, n_hold, n_wait = b.done|length, b.active|length, b.hold|length, b.waiting_total %}
<header class="pb-head">
  <div class="pb-head-main">
  <form method="get" class="pb-scope-form">
    <h1 class="pb-h1">
      <select name="workplace_id" class="pb-scope-select" aria-label="職場で絞り込む" onchange="this.form.submit()">
        <option value="">全体</option>
        {% for w in workplaces %}<option value="{{ w.id }}" {% if selected_wp and selected_wp.id == w.id %}selected{% endif %}>{{ w.name }}</option>{% endfor %}
      </select>活動状況
    </h1>
  </form>
  <div class="pb-summary">
    <div class="pb-bar" role="img" aria-label="完了 {{ n_done }}、進行中 {{ n_active }}、保留 {{ n_hold }}、順番待ち {{ n_wait }}">
      {% if n_done %}<span class="pb-seg seg-done" style="flex-grow: {{ n_done }}"></span>{% endif %}
      {% if n_active %}<span class="pb-seg seg-active" style="flex-grow: {{ n_active }}"></span>{% endif %}
      {% if n_hold %}<span class="pb-seg seg-hold" style="flex-grow: {{ n_hold }}"></span>{% endif %}
      {% if n_wait %}<span class="pb-seg seg-wait" style="flex-grow: {{ n_wait }}"></span>{% endif %}
    </div>
    <ul class="pb-legend">
      <li><i class="pb-dotc seg-done"></i>完了<b>{{ n_done }}</b></li>
      <li><i class="pb-dotc seg-active"></i>進行中<b>{{ n_active }}</b></li>
      {% if n_hold %}<li><i class="pb-dotc seg-hold"></i>保留<b>{{ n_hold }}</b></li>{% endif %}
      <li><i class="pb-dotc seg-wait"></i>順番待ち<b>{{ n_wait }}</b></li>
      <li class="pb-total">全<b>{{ n_done + n_active + n_hold + n_wait }}</b>件</li>
    </ul>
  </div>
  </div>
  <div class="pb-cta">
    <div class="pb-cta-text">あなたの定型業務も、依頼してみませんか</div>
    <div class="pb-cta-actions">
      <a class="pb-cta-btn" href="{{ url_for('survey.index') }}">自動化依頼を出す</a>
      <a class="pb-cta-link" href="{{ url_for('monthly.form') }}">月次アンケートに答える</a>
    </div>
  </div>
</header>

{% if hl %}
<section class="pb-wins" aria-label="これまでの成果と依頼">
  <div class="pb-strip">
    {% if hl.tasks %}
    <div class="card pb-stat"><div class="card-body">
      <div class="stat-head"><span class="stat-tile tile-blue"><i class="bi bi-kanban"></i></span><span class="stat-label">累計改善案件</span></div>
      <div class="stat-value">{{ hl.tasks }}<span class="stat-unit">件</span></div>
      <div class="stat-sub">受付 {{ hl.status_counts.get('受付', 0) }} ・ 対応中 {{ hl.status_counts.get('対応中', 0) }} ・ 完了 {{ hl.completed }} ・ 保留 {{ hl.status_counts.get('保留', 0) }}</div>
    </div></div>
    <div class="card pb-stat"><div class="card-body">
      <div class="stat-head"><span class="stat-tile tile-green"><i class="bi bi-hourglass-split"></i></span><span class="stat-label">年間削減工数</span></div>
      <div class="stat-value">{{ '%.0f'|format(hl.saved_annual_hours) }}<span class="stat-unit">時間</span></div>
      <div class="stat-sub">月 {{ '%.0f'|format(hl.saved_hours) }} 時間 ・ 削減率 {{ '%.1f'|format(hl.reduction_rate) }}%（棚卸し月間 {{ '%.1f'|format(hl.total_hours) }} 時間に対して）</div>
    </div></div>
    {% endif %}
    {% for t in hl.recent %}
    <div class="card pb-case"><div class="card-body">
      <div class="stat-head"><span class="stat-tile tile-teal"><i class="bi bi-trophy"></i></span><span class="stat-label">改善事例</span></div>
      <div class="pb-case-name">{{ t.name }}</div>
      <div class="pb-case-meta">{{ t.workplace.name if t.workplace else '' }}{% if show_names and (t.assignee or t.respondent) %} · {{ t.assignee or t.respondent.name }}{% endif %}{% if t.decision and t.decision != '未定' %} · {{ t.decision }}{% endif %}</div>
      <div class="pb-case-num">月 {{ '%.0f'|format(t.monthly_minutes or 0) }} 分 → {% if t.after_minutes is not none %}{{ '%.0f'|format(t.after_minutes) }} 分{% else %}－{% endif %}{% if t.saved_minutes %}<b class="pb-case-save">−{{ '%.0f'|format(t.saved_minutes) }} 分</b>{% endif %}</div>
      {% if t.result_note %}<div class="pb-case-note">{{ t.result_note }}</div>{% endif %}
    </div></div>
    {% endfor %}
  </div>
</section>
{% endif %}

<section>
  <h2 class="pb-section">進行中 <span class="pb-n">{{ b.active|length }}</span></h2>
  {% for t in b.active %}{{ row(t, show_names) }}{% else %}<div class="pb-empty">進行中の案件はありません</div>{% endfor %}
</section>

{% if b.hold %}
<section>
  <h2 class="pb-section">保留中 <span class="pb-n">{{ b.hold|length }}</span></h2>
  {% for t in b.hold %}{{ row(t, show_names) }}{% endfor %}
</section>
{% endif %}

<section>
  <h2 class="pb-section">完了 <span class="pb-n">{{ b.done|length }}</span>{% if b.done and current_user %}<a class="pb-link" href="{{ url_for('cases.index', stage='done') }}">完了した案件を見る</a>{% endif %}</h2>
  {% for t in b.done_shown %}{{ row(t, show_names) }}{% else %}<div class="pb-empty">完了した案件はまだありません</div>{% endfor %}
</section>

<section>
  <h2 class="pb-section">順番待ち <span class="pb-n">{{ b.waiting_total }}</span>{% if b.waiting_total and current_user %}<a class="pb-link" href="{{ url_for('cases.index', status='受付') }}">一覧で見る</a>{% endif %}</h2>
  {% if b.waiting %}
  <ul class="pb-waiting">
    {% for t in b.waiting %}
    <li>
      <span class="text-truncate">{% if current_user %}<a href="{{ url_for('cases.detail', task_id=t.id) }}">{{ t.name }}</a>{% else %}{{ t.name }}{% endif %}</span>
      <span class="pb-dim">{{ t.workplace.name if t.workplace else '' }}{% if show_names and t.respondent %} {{ t.respondent.name }}{% endif %}</span>
    </li>
    {% endfor %}
  </ul>
  {% else %}<div class="pb-empty">順番待ちの案件はありません</div>{% endif %}
</section>
''',
    'board/_macros.html': r'''{# 活動状況の部品（{% from "board/_macros.html" import stepper, row with context %} で使う） #}
{% macro stepper(t, small=False) %}
{% set dates = t.stage_dates %}
<ol class="pb-steps{% if small %} pb-steps-sm{% endif %}{% if t.is_on_hold %} pb-hold{% endif %}{% if t.stage_key == 'done' %} pb-complete{% endif %}" aria-label="進捗 {{ t.stage_name }}">
  {% for key, name, desc, steps in C.STAGES %}
  {% set i = loop.index0 %}
  {% if t.stage_key == 'done' %}{% set cls = 'is-done is-current' if loop.last else 'is-done' %}
  {% elif i < t.stage_index %}{% set cls = 'is-done' %}
  {% elif i == t.stage_index %}{% set cls = 'is-current' %}
  {% else %}{% set cls = '' %}{% endif %}
  <li class="pb-step {{ cls }}" title="{{ t.stage_label(key) }}：{{ desc }}">
    <span class="pb-track"></span>
    <span class="pb-label">{{ t.stage_label(key) }}</span>
    <span class="pb-date">{% if cls and dates.get(key) %}{{ dates.get(key)|jpdate('%m/%d') }}{% endif %}</span>
  </li>
  {% endfor %}
</ol>
{% endmacro %}

{% macro row(t, show_names=True) %}
<div class="pb-row">
  <div class="pb-from">
    <div class="pb-wp">{{ t.workplace.name if t.workplace else '職場未設定' }}</div>
    {% if show_names and t.respondent %}<div class="pb-name">{{ t.respondent.name }}</div>{% endif %}
  </div>
  <div class="pb-what">
    <div class="pb-title">{% if current_user %}<a href="{{ url_for('cases.detail', task_id=t.id) }}">{{ t.name }}</a>{% else %}{{ t.name }}{% endif %}{% if t.is_on_hold %}<span class="pb-pill">保留</span>{% endif %}</div>
    {% set note = t.progress_note or (t.result_note if t.status == '完了' else '') %}
    {% if note %}<div class="pb-note">{{ note }}</div>{% endif %}
  </div>
  <div class="pb-progress">{{ stepper(t) }}</div>
</div>
{% endmacro %}
''',
    'board/index.html': r'''{% extends "base.html" %}
{% block title %}活動状況{% endblock %}
{% block content %}
{% include "board/_board.html" %}
{% endblock %}
''',
    'cases/detail.html': r'''{% extends "base.html" %}
{% from "board/_macros.html" import stepper with context %}
{% block title %}{{ task.name }}{% endblock %}
{% block content %}
<nav aria-label="breadcrumb"><ol class="breadcrumb small mb-2">
  <li class="breadcrumb-item"><a href="{{ url_for('cases.index') }}">改善案件</a></li>
  <li class="breadcrumb-item active">{{ task.name }}</li>
</ol></nav>

<div class="d-flex flex-wrap justify-content-between align-items-start gap-2 mb-3">
  <div>
    <h2 class="mb-1">{{ task.name }}</h2>
    <div class="text-muted">
      {{ task.workplace.name if task.workplace else '' }}／{{ task.respondent.name if task.respondent else '' }}
      ・ 登録 {{ task.created_at|jpdate }}
      ・ 月間 <strong>{{ task.monthly_minutes|hours }} 時間</strong>（{{ task.frequency }}／{{ task.duration }}／{{ task.headcount }}人）
    </div>
  </div>
  <div>
    <span class="badge fs-6 text-bg-{{ {'受付':'secondary','対応中':'primary','完了':'success','保留':'warning'}.get(task.status, 'secondary') }}">{{ task.status }}</span>
    {% if task.decision and task.decision != '未定' %}<span class="badge fs-6 text-bg-dark">{{ task.decision }}</span>{% endif %}
    {% if task.verified %}<span class="badge fs-6 text-bg-success">担当者確認済み</span>{% else %}<span class="badge fs-6 text-bg-light text-muted">未確認</span>{% endif %}
  </div>
</div>

<div class="row g-3">
  <div class="col-lg-7">
    <div class="card mb-3"><div class="card-body">
      <h5 class="card-title">依頼内容（登録された業務情報）</h5>
      <table class="table table-sm mb-0">
        <tr><th class="w-25">業務領域</th><td>{{ task.domain.name if task.domain else '未選択' }}{% if task.ai_domain and (not task.domain or task.ai_domain != task.domain.name) %} <span class="badge badge-flag">AI推定: {{ task.ai_domain }}</span>{% endif %}</td></tr>
        <tr><th>目的</th><td>{{ task.purpose }}{% if task.purpose_unknown %} <span class="badge text-bg-danger">目的不明</span>{% endif %}</td></tr>
        <tr><th>アウトプット</th><td>{{ task.output }}</td></tr>
        <tr><th>利用者</th><td>{{ task.output_users_list|join('・') }}{% if task.output_users_note %}<div class="small text-muted">{{ task.output_users_note }}</div>{% endif %}</td></tr>
        <tr><th>入力情報</th><td>{{ task.inputs_list|join('・') }}{% if task.inputs_note %}<div class="small text-muted">{{ task.inputs_note }}</div>{% endif %}</td></tr>
        <tr><th>前後の業務</th><td>{{ task.prev_task or '－' }} → <strong>{{ task.name }}</strong> → {{ task.next_task or '－' }}</td></tr>
        <tr><th>関わる人・部署</th><td>{{ task.actors_list|join('・') or '－' }}</td></tr>
        <tr><th>情報の受け取り元</th><td>{{ task.info_from or '－' }}</td></tr>
        <tr><th>使用ツール</th><td>{{ task.tools_list|join('・') }}</td></tr>
        <tr><th>作業の性質</th><td>{{ task.nature_list|join('・') }}</td></tr>
        <tr><th>困りごと</th><td>{{ task.pains_list|join('・') }}{% if task.pains_note %}<div class="small text-muted">{{ task.pains_note }}</div>{% endif %}</td></tr>
        <tr><th>改善の希望</th><td>{{ task.wish or '' }}</td></tr>
        <tr><th>自分で改善</th><td>{{ task.self_improve or '' }}</td></tr>
        {% if task.note %}<tr><th>補足</th><td>{{ task.note }}</td></tr>{% endif %}
      </table>
    </div></div>

    <div class="card mb-3"><div class="card-body">
      <div class="d-flex justify-content-between align-items-start flex-wrap gap-2">
        <div>
          <h5 class="mb-1">この業務の流れ</h5>
          <p class="small text-muted mb-0">情報源 → 業務 → アウトプット → 利用者、関わる人。<a href="{{ url_for('bizmap.index', workplace_id=task.workplace_id) }}">職場全体の流れ図</a></p>
        </div>
        <div class="btn-group btn-group-sm flow-toolbar" data-target="task" role="group">
          <button type="button" class="btn btn-outline-secondary" data-zoom="out">－</button>
          <button type="button" class="btn btn-outline-secondary" data-zoom="reset">100%</button>
          <button type="button" class="btn btn-outline-secondary" data-zoom="in">＋</button>
        </div>
      </div>
      <div class="flow-canvas"><div class="flow-inner" id="flow-task" data-def="def-task"><span class="text-muted small">描画中…</span></div></div>
      <script type="text/plain" id="def-task">{{ task_flow|safe }}</script>
    </div></div>

    <div class="card mb-3"><div class="card-body">
      <div class="d-flex justify-content-between align-items-center">
        <h5 class="card-title mb-0">AI分析
          {% if task.ai_source == 'claude' %}<span class="badge text-bg-primary">Claude API</span>{% elif task.ai_source %}<span class="badge text-bg-secondary">ルールベース</span>{% endif %}
          {% if task.ai_priority %}<span class="badge text-bg-{{ {'高':'danger','中':'warning','低':'light text-muted'}.get(task.ai_priority, 'light') }}">優先度 {{ task.ai_priority }}</span>{% endif %}
        </h5>
        <form method="post" action="{{ url_for('cases.reanalyze', task_id=task.id) }}">
          <button class="btn btn-sm btn-outline-primary">{% if ai_enabled %}AIで再分析{% else %}再分析（ルール）{% endif %}</button>
        </form>
      </div>
      {% if task.ai_summary %}
      <p class="mt-2 mb-2">{{ task.ai_summary }}</p>
      {% endif %}
      {% if task.ai_flags_list %}
      <div class="mb-2">{% for f in task.ai_flags_list %}<span class="badge badge-flag me-1">{{ f }}</span>{% endfor %}</div>
      {% endif %}
      {% set ecrs = task.ai_ecrs_dict %}
      {% if ecrs %}
      <table class="table table-sm mb-2">
        <tbody>
          {% for key, label, question in C.ECRS %}
          {% if ecrs.get(key) %}<tr><th class="w-25">{{ label }}</th><td>{{ ecrs.get(key) }}</td></tr>{% endif %}
          {% endfor %}
          {% if ecrs.get('automation') %}<tr><th>自動化（最適化後）</th><td>{{ ecrs.get('automation') }}</td></tr>{% endif %}
        </tbody>
      </table>
      {% if ecrs.get('questions') %}
      <div class="small"><strong>担当者に確認すること</strong>
        <ul class="mb-0">{% for q in ecrs.get('questions') %}<li>{{ q }}</li>{% endfor %}</ul>
      </div>
      {% endif %}
      {% endif %}
      {% if not task.ai_source %}
      <p class="text-muted small mb-0">まだ分析されていません。「再分析」を押してください。</p>
      {% endif %}
      <p class="small text-muted mt-2 mb-0">AIの推定は参考情報です。担当者確認（右側）を経て正式な情報になります。</p>
    </div></div>

    <div class="card"><div class="card-body">
      <h5 class="card-title">関連しそうな業務</h5>
      {% set shown = [] %}
      <ul class="mb-0">
        {% for r in related_by_ai %}{% if r.id not in shown %}{% set _ = shown.append(r.id) %}
        <li><a href="{{ url_for('cases.detail', task_id=r.id) }}">{{ r.name }}</a> <span class="badge badge-flag">AI</span> <span class="small text-muted">{{ r.workplace.name if r.workplace else '' }}／{{ r.monthly_minutes|hours }} h</span></li>
        {% endif %}{% endfor %}
        {% for r in related %}{% if r.id not in shown %}{% set _ = shown.append(r.id) %}
        <li><a href="{{ url_for('cases.detail', task_id=r.id) }}">{{ r.name }}</a> <span class="small text-muted">{{ r.workplace.name if r.workplace else '' }}／{{ r.monthly_minutes|hours }} h</span></li>
        {% endif %}{% endfor %}
        {% if not shown %}<li class="text-muted">関連業務は見つかっていません</li>{% endif %}
      </ul>
    </div></div>
  </div>

  <div class="col-lg-5">
    <div class="card mb-3"><div class="card-body">
      <div class="d-flex justify-content-between align-items-center flex-wrap gap-2 mb-2">
        <h5 class="card-title mb-0">進捗（製造部全体に公開）</h5>
        <a class="small" href="{{ url_for('board.index') }}">活動状況を見る</a>
      </div>
      {{ stepper(task) }}
      <div class="d-flex justify-content-between align-items-center mt-2 gap-2 flex-wrap">
        <div class="small text-muted">いま <strong class="text-dark">{{ task.stage_name }}</strong>（{{ task.stage_since|jpdate }}〜）{% if task.is_on_hold %} <span class="badge text-bg-warning">保留</span>{% endif %}</div>
        <div class="d-flex gap-1">
          <form method="post" action="{{ url_for('cases.toggle_hold', task_id=task.id) }}"><button class="btn btn-sm btn-outline-warning" {% if task.stage_key == 'done' %}disabled{% endif %}>{% if task.is_on_hold %}保留を解除{% else %}保留にする{% endif %}</button></form>
          <form method="post" action="{{ url_for('cases.set_stage', task_id=task.id) }}"><input type="hidden" name="direction" value="prev"><button class="btn btn-sm btn-outline-secondary" {% if task.stage_index == 0 %}disabled{% endif %}>◀ 戻す</button></form>
          <form method="post" action="{{ url_for('cases.set_stage', task_id=task.id) }}"><input type="hidden" name="direction" value="next"><button class="btn btn-sm btn-success" {% if task.stage_key == 'done' %}disabled{% endif %}>次の段階へ ▶</button></form>
        </div>
      </div>
      {% for key, name, desc, steps in C.STAGES %}{% if key == task.stage_key %}
      <div class="small text-muted mt-2">{{ task.stage_label(key) }}：{{ desc }}{% if steps %}（{{ steps }}）{% endif %}</div>
      {% endif %}{% endfor %}
      <div class="small text-muted mt-1">ステータス（受付・対応中・完了）は段階から自動で決まります。</div>
      {% if task.progress_note %}<div class="pb-note mt-1">{{ task.progress_note }}</div>{% endif %}
    </div></div>

    <form method="post" action="{{ url_for('cases.update', task_id=task.id) }}" class="card mb-3"><div class="card-body">
      <h5 class="card-title">案件管理</h5>
      <div class="row g-2 mb-2">
        <div class="col-6">
          <label class="form-label small mb-0">改善方針</label>
          <select name="decision" class="form-select form-select-sm">
            {% for d in C.DECISIONS %}<option value="{{ d }}" {% if task.decision == d %}selected{% endif %}>{{ d }}</option>{% endfor %}
          </select>
        </div>
        <div class="col-6">
          <label class="form-label small mb-0">担当区分</label>
          <select name="assignment" class="form-select form-select-sm">
            {% for a in C.ASSIGNMENTS %}<option value="{{ a }}" {% if task.assignment == a %}selected{% endif %}>{{ a }}</option>{% endfor %}
          </select>
        </div>
        <div class="col-6">
          <label class="form-label small mb-0">担当者</label>
          <input type="text" name="assignee" class="form-control form-control-sm" value="{{ task.assignee or '' }}">
        </div>
        <div class="col-12">
          <label class="form-label small mb-0">活動状況のひとこと（今の状況・次の予定）</label>
          <input type="text" name="progress_note" class="form-control form-control-sm" value="{{ task.progress_note or '' }}" placeholder="例：11月中旬にテスト予定" maxlength="120">
        </div>
      </div>

      <label class="form-label small mb-1 mt-2"><strong>標準手順</strong>（自動化は最後）</label>
      <div class="step-list mb-2">
        {% for no, name, desc in C.STEPS %}
        <div class="form-check">
          <input class="form-check-input" type="checkbox" name="steps" value="{{ no }}" id="step-{{ no }}" {% if no in task.steps_done_list %}checked{% endif %}>
          <label class="form-check-label {% if no in task.steps_done_list %}step-done{% endif %}" for="step-{{ no }}">{{ no }}. {{ name }} <span class="text-muted small">— {{ desc }}</span></label>
        </div>
        {% endfor %}
      </div>

      <label class="form-label small mb-1 mt-2"><strong>ECRSの検討メモ</strong></label>
      {% for key, label, question in C.ECRS %}
      <div class="mb-2">
        <label class="form-label small mb-0">{{ label }} <span class="text-muted">— {{ question }}</span></label>
        <textarea name="ecrs_{{ key }}" class="form-control form-control-sm" rows="2">{{ task['ecrs_' ~ key] or '' }}</textarea>
      </div>
      {% endfor %}

      <label class="form-label small mb-1 mt-2"><strong>効果測定</strong></label>
      <div class="row g-2 mb-2">
        <div class="col-6">
          <label class="form-label small mb-0">改善前 月間（分）</label>
          <input type="text" class="form-control form-control-sm" value="{{ '%.0f'|format(task.monthly_minutes or 0) }}" readonly>
        </div>
        <div class="col-6">
          <label class="form-label small mb-0">改善後 月間（分）</label>
          <input type="number" name="after_minutes" class="form-control form-control-sm" min="0" step="1" value="{% if task.after_minutes is not none %}{{ '%.0f'|format(task.after_minutes) }}{% endif %}" placeholder="廃止なら 0">
        </div>
      </div>
      {% if task.saved_minutes %}
      <div class="small text-success mb-2">削減 {{ '%.0f'|format(task.saved_minutes) }} 分/月（年間 {{ '%.1f'|format(task.saved_minutes * 12 / 60) }} 時間）</div>
      {% endif %}
      <div class="mb-2">
        <label class="form-label small mb-0">改善内容・結果</label>
        <textarea name="result_note" class="form-control form-control-sm" rows="3">{{ task.result_note or '' }}</textarea>
      </div>
      <button class="btn btn-primary w-100">保存する</button>
    </div></form>

    <form method="post" action="{{ url_for('cases.verify', task_id=task.id) }}" class="card"><div class="card-body">
      <h5 class="card-title">担当者確認（推進メンバー）</h5>
      <p class="small text-muted">必須項目の内容とAIの推定を本人・推進メンバーが確認すると、業務情報の充足率に算入されます。</p>
      {% if not task.required_filled %}
      <div class="alert alert-warning py-1 small">必須項目に未入力があります。ヒアリングで補完してください。</div>
      {% endif %}
      {% if task.verified %}
        <div class="small mb-2">確認者: {{ task.verified_by or '－' }}</div>
        <button class="btn btn-outline-secondary btn-sm w-100">確認を解除する</button>
      {% else %}
        <div class="mb-2">
          <input type="text" name="verified_by" class="form-control form-control-sm" placeholder="確認者名（推進メンバー）">
        </div>
        {% if task.ai_domain and task.domain and task.ai_domain != task.domain.name %}
        <div class="form-check mb-2">
          <input class="form-check-input" type="checkbox" name="adopt_ai_domain" id="adopt" value="1">
          <label class="form-check-label small" for="adopt">業務領域をAI推定（{{ task.ai_domain }}）に変更する</label>
        </div>
        {% endif %}
        <button class="btn btn-success btn-sm w-100">内容を確認した</button>
      {% endif %}
    </div></form>
  </div>
</div>
{% endblock %}
{% block scripts %}
<script src="https://cdn.jsdelivr.net/npm/mermaid@11.4.1/dist/mermaid.min.js"></script>
<script src="{{ url_for('static', filename='flows.js') }}"></script>
{% endblock %}
''',
    'cases/index.html': r'''{% extends "base.html" %}
{% from "board/_macros.html" import stepper with context %}
{% block title %}改善案件一覧{% endblock %}
{% block content %}
<div class="d-flex flex-wrap justify-content-between align-items-center mb-2 gap-2">
  <h2 class="mb-0">改善案件一覧</h2>
  <div class="d-flex gap-2">
    <a class="btn btn-outline-secondary btn-sm" href="{{ url_for('cases.export_csv') }}">CSVエクスポート</a>
    <a class="btn btn-primary btn-sm" href="{{ url_for('survey.index') }}">自動化依頼を出す</a>
  </div>
</div>
<p class="text-muted small">登録された業務はすべて改善案件の候補です。</p>

<form method="get" class="card mb-3"><div class="card-body py-2">
  <div class="row g-2 align-items-end">
    <div class="col-6 col-md-4 col-xl">
      <label class="form-label small mb-0">ステータス</label>
      <select name="status" class="form-select form-select-sm">
        <option value="">すべて</option>
        {% for s in C.STATUSES %}<option value="{{ s }}" {% if args.get('status') == s %}selected{% endif %}>{{ s }}</option>{% endfor %}
      </select>
    </div>
    <div class="col-6 col-md-4 col-xl">
      <label class="form-label small mb-0">進捗の段階</label>
      <select name="stage" class="form-select form-select-sm">
        <option value="">すべて</option>
        {% for key, name, desc, steps in C.STAGES %}<option value="{{ key }}" {% if args.get('stage') == key %}selected{% endif %}>{{ loop.index }}. {{ '作成／伴走' if key == 'build' else name }}</option>{% endfor %}
      </select>
    </div>
    <div class="col-6 col-md-4 col-xl">
      <label class="form-label small mb-0">職場</label>
      <select name="workplace_id" class="form-select form-select-sm">
        <option value="">すべて</option>
        {% for w in workplaces %}<option value="{{ w.id }}" {% if args.get('workplace_id')|string == w.id|string %}selected{% endif %}>{{ w.name }}</option>{% endfor %}
      </select>
    </div>
    <div class="col-6 col-md-4 col-xl">
      <label class="form-label small mb-0">業務領域</label>
      <select name="domain_id" class="form-select form-select-sm">
        <option value="">すべて</option>
        {% for d in domains %}<option value="{{ d.id }}" {% if args.get('domain_id')|string == d.id|string %}selected{% endif %}>{{ d.name }}</option>{% endfor %}
      </select>
    </div>
    <div class="col-6 col-md-4 col-xl">
      <label class="form-label small mb-0">絞り込み</label>
      <select name="filter" class="form-select form-select-sm">
        <option value="">すべて</option>
        <option value="candidates" {% if args.get('filter') == 'candidates' %}selected{% endif %}>改善希望・困りごとあり</option>
        <option value="elimination" {% if args.get('filter') == 'elimination' %}selected{% endif %}>廃止・統合候補</option>
        <option value="self" {% if args.get('filter') == 'self' %}selected{% endif %}>自主改善希望</option>
        <option value="unverified" {% if args.get('filter') == 'unverified' %}selected{% endif %}>担当者未確認</option>
      </select>
    </div>
    <div class="col-6 col-md-4 col-xl">
      <label class="form-label small mb-0">キーワード</label>
      <input type="text" name="q" class="form-control form-control-sm" value="{{ args.get('q', '') }}">
    </div>
    <div class="col-6 col-md-4 col-xl d-flex gap-1">
      <select name="sort" class="form-select form-select-sm">
        <option value="minutes" {% if args.get('sort', 'minutes') == 'minutes' %}selected{% endif %}>月間時間順</option>
        <option value="new" {% if args.get('sort') == 'new' %}selected{% endif %}>新しい順</option>
        <option value="name" {% if args.get('sort') == 'name' %}selected{% endif %}>名前順</option>
      </select>
      <button class="btn btn-sm btn-primary text-nowrap">表示</button>
    </div>
  </div>
</div></form>

<div class="small text-muted mb-2">{{ tasks|length }} 件、月間 合計 {{ total_minutes|hours }} 時間</div>
<div class="table-responsive">
<table class="table table-sm table-hover align-middle bg-white">
  <thead><tr>
    <th>業務名</th><th>職場／依頼者</th><th>領域</th><th class="text-end">月間</th><th>改善の希望</th><th>自分で</th><th>ステータス</th><th>進捗（活動状況）</th><th>手順</th><th>確認</th><th>AI</th>
  </tr></thead>
  <tbody>
  {% for t in tasks %}
  <tr>
    <td><a href="{{ url_for('cases.detail', task_id=t.id) }}">{{ t.name }}</a><div class="small text-muted">{{ t.output }}</div></td>
    <td class="small">{{ t.workplace.name if t.workplace else '' }}<br>{{ t.respondent.name if t.respondent else '' }}</td>
    <td class="small">{{ t.domain.name if t.domain else '' }}</td>
    <td class="text-end">{{ t.monthly_minutes|hours }} h</td>
    <td class="small">{{ t.wish or '' }}</td>
    <td class="small">{% if t.self_improve in C.SELF_IMPROVE_POSITIVE %}<span class="badge text-bg-success">希望</span>{% else %}{{ t.self_improve or '' }}{% endif %}</td>
    <td><span class="badge text-bg-{{ {'受付':'secondary','対応中':'primary','完了':'success','保留':'warning'}.get(t.status, 'secondary') }}">{{ t.status }}</span>{% if t.decision and t.decision != '未定' %}<div class="small">{{ t.decision }}</div>{% endif %}</td>
    <td class="small" style="min-width: 150px"><div class="fw-semibold">{{ t.stage_index + 1 }}. {{ t.stage_name }}</div>{{ stepper(t, small=True) }}</td>
    <td class="small">{{ t.steps_done_list|length }}/5</td>
    <td>{% if t.verified %}<span class="badge text-bg-success">済</span>{% else %}<span class="badge text-bg-light text-muted">未</span>{% endif %}</td>
    <td class="small">{% if t.ai_priority %}<span class="badge text-bg-{{ {'高':'danger','中':'warning','低':'light text-muted'}.get(t.ai_priority, 'light') }}">{{ t.ai_priority }}</span>{% endif %}
      {% for f in t.ai_flags_list[:2] %}<span class="badge badge-flag">{{ f }}</span>{% endfor %}</td>
  </tr>
  {% else %}
  <tr><td colspan="11" class="text-muted">該当する案件はありません</td></tr>
  {% endfor %}
  </tbody>
</table>
</div>
{% endblock %}
''',
    'errors/404.html': r'''{% extends "base.html" %}
{% block title %}ページが見つかりません{% endblock %}
{% block content %}
<div class="login-wrap">
  <div class="card login-card"><div class="card-body text-center">
    <div class="login-mark"><i class="bi bi-signpost-split"></i></div>
    <h1 class="login-title">ページが見つかりません</h1>
    <p class="login-sub">URLが変わったか、削除された可能性があります。</p>
    <a class="btn btn-primary" href="{{ url_for('main.index') }}">トップへ</a>
  </div></div>
</div>
{% endblock %}
''',
    'errors/500.html': r'''{% extends "base.html" %}
{% block title %}エラーが発生しました{% endblock %}
{% block content %}
<div class="login-wrap">
  <div class="card login-card"><div class="card-body text-center">
    <div class="login-mark"><i class="bi bi-exclamation-triangle"></i></div>
    <h1 class="login-title">エラーが発生しました</h1>
    <p class="login-sub">しばらくしてからもう一度お試しください。続くときは推進メンバーに知らせてください。</p>
    <a class="btn btn-primary" href="{{ url_for('main.index') }}">トップへ</a>
  </div></div>
</div>
{% endblock %}
''',
    'index.html': r'''{% extends "base.html" %}
{% block title %}ダッシュボード{% endblock %}
{% block content %}
{% if k.task_count == 0 %}
<div class="alert alert-info">
  まだ業務が登録されていません。<a href="{{ url_for('survey.index') }}">アンケート・業務登録</a>から登録するか、
  <a href="{{ url_for('admin.index') }}">管理</a>画面からデモデータを投入すると画面の動きを確認できます。
</div>
{% endif %}

<h2 class="section-title">活動の成果</h2>
<div class="row g-3 mb-4">
  <div class="col-6 col-lg-3">
    <div class="card stat h-100"><div class="card-body">
      <div class="stat-head"><span class="stat-tile tile-blue"><i class="bi bi-kanban"></i></span><span class="stat-label">累計改善案件</span></div>
      <div class="stat-value">{{ k.task_count }}<span class="stat-unit">件</span></div>
      <div class="stat-sub">受付 {{ k.status_counts.get('受付', 0) }} ・ 対応中 {{ k.status_counts.get('対応中', 0) }} ・ 完了 {{ k.completed_count }} ・ 保留 {{ k.status_counts.get('保留', 0) }}</div>
    </div></div>
  </div>
  <div class="col-6 col-lg-3">
    <div class="card stat h-100"><div class="card-body">
      <div class="stat-head"><span class="stat-tile tile-green"><i class="bi bi-hourglass-split"></i></span><span class="stat-label">年間削減工数</span></div>
      <div class="stat-value">{{ '%.0f'|format(k.saved_annual_hours) }}<span class="stat-unit">時間</span></div>
      <div class="stat-sub">削減率 {{ '%.1f'|format(k.reduction_rate) }}%（棚卸し月間 {{ '%.1f'|format(k.total_hours) }} 時間に対して）</div>
    </div></div>
  </div>
  <div class="col-6 col-lg-3">
    <div class="card stat h-100"><div class="card-body">
      <div class="stat-head"><span class="stat-tile tile-indigo"><i class="bi bi-people"></i></span><span class="stat-label">自主改善者</span></div>
      <div class="stat-value">{{ k.self_improver_count }}<span class="stat-unit">人</span></div>
      <div class="stat-sub">自ら改善に取り組みたい人（回答者 {{ k.respondent_count }} 人中）</div>
    </div></div>
  </div>
  <div class="col-6 col-lg-3">
    <div class="card stat h-100"><div class="card-body">
      <div class="stat-head"><span class="stat-tile tile-orange"><i class="bi bi-lightning-charge"></i></span><span class="stat-label">自動化・廃止業務</span></div>
      <div class="stat-value">{{ k.automated_count }}<span class="stat-unit">件</span></div>
      <div class="stat-sub">完了案件のうち廃止・自動化したもの</div>
    </div></div>
  </div>
</div>

<h2 class="section-title">現状の把握 <span class="section-sub">自動化依頼（棚卸し）から</span></h2>
<div class="row g-3 mb-4">
  <div class="col-6 col-lg-3">
    <div class="card stat h-100"><div class="card-body">
      <div class="stat-head"><span class="stat-tile tile-teal"><i class="bi bi-list-check"></i></span><span class="stat-label">PC定型業務の月間時間</span></div>
      <div class="stat-value">{{ '%.1f'|format(k.total_hours) }}<span class="stat-unit">時間/月</span></div>
      <div class="stat-sub">年間換算 {{ '%.0f'|format(k.annual_hours) }} 時間 ・ 案件候補 {{ k.candidate_count }} 件</div>
    </div></div>
  </div>
  <div class="col-6 col-lg-3">
    <div class="card stat h-100"><div class="card-body">
      <div class="stat-head"><span class="stat-tile tile-pink"><i class="bi bi-envelope-paper"></i></span><span class="stat-label">アンケート回答率</span></div>
      <div class="stat-value">{% if k.response_rate is not none %}{{ '%.0f'|format(k.response_rate) }}<span class="stat-unit">%</span>{% else %}－{% endif %}</div>
      <div class="stat-sub">回答者 {{ k.respondent_count }} 人 ／ 対象 {{ k.target_respondents or '未設定' }} 人</div>
    </div></div>
  </div>
  <div class="col-6 col-lg-3">
    <div class="card stat h-100"><div class="card-body">
      <div class="stat-head"><span class="stat-tile tile-purple"><i class="bi bi-check2-circle"></i></span><span class="stat-label">業務情報の充足率</span></div>
      <div class="stat-value">{{ '%.0f'|format(k.sufficient_rate) }}<span class="stat-unit">%</span></div>
      <div class="stat-sub">必須項目あり {{ '%.0f'|format(k.filled_rate) }}% ・ 担当者確認済み {{ '%.0f'|format(k.verified_rate) }}%</div>
    </div></div>
  </div>
  <div class="col-6 col-lg-3">
    <div class="card stat h-100"><div class="card-body">
      <div class="stat-head"><span class="stat-tile tile-red"><i class="bi bi-exclamation-triangle"></i></span><span class="stat-label">分析不足の領域</span></div>
      <div class="stat-value">{{ k.insufficient_domains|length }}<span class="stat-unit">/ {{ k.domain_rows|length }} 領域</span></div>
      <div class="stat-sub">{% if k.insufficient_domains %}{{ k.insufficient_domains|map(attribute='domain.name')|join('、') }}{% else %}すべての領域で情報が集まっています{% endif %} <a href="{{ url_for('bizmap.index') }}#tab-domains">領域別の把握状況</a></div>
    </div></div>
  </div>
</div>

<p class="dash-link"><a href="{{ url_for('board.index') }}"><i class="bi bi-kanban"></i>活動状況で案件の進み具合を見る（進行中 {{ active_count }} 件）<i class="bi bi-arrow-right-short"></i></a></p>

<h2 class="section-title">アンケートと職場の状況</h2>
<div class="row g-3 mb-4">
  <div class="col-lg-5">
    <div class="card h-100"><div class="card-body">
      <div class="d-flex justify-content-between align-items-center">
        <h5 class="card-title mb-0">月次アンケート（時系列）</h5>
        <a class="small" href="{{ url_for('monthly.results') }}">月次集計へ</a>
      </div>
      {% if monthly_latest %}
      <div class="small text-muted mb-2">最新 {{ monthly_latest.label }}：回答 {{ monthly_latest.n }} 人{% if monthly_latest.response_rate is not none %}（{{ '%.0f'|format(monthly_latest.response_rate) }}%）{% endif %}</div>
      <div class="row g-2 text-center mb-2">
        <div class="col-4"><div class="kpi-label">繰り返し業務あり</div><div class="mini-value">{{ '%.0f'|format(monthly_latest.has_rate) }}%</div></div>
        <div class="col-4"><div class="kpi-label">先月より減った</div><div class="mini-value text-green">{{ '%.0f'|format(monthly_latest.decreased_rate) }}%</div></div>
        <div class="col-4"><div class="kpi-label">自分でやりたい</div><div class="mini-value text-blue">{{ '%.0f'|format(monthly_latest.self_rate) }}%</div></div>
      </div>
      <div class="chart-box" style="height: 150px"><canvas id="chart-monthly"></canvas></div>
      {% else %}
      <p class="text-muted small mb-2">毎月同じ設問で繰り返し業務の変化を追います。まだ回答がありません。</p>
      <a class="btn btn-sm btn-outline-secondary" href="{{ url_for('monthly.form') }}">回答フォーム</a>
      {% endif %}
    </div></div>
  </div>
  <div class="col-lg-7">
    <div class="card h-100"><div class="card-body">
      <h5 class="card-title">職場別の回答状況</h5>
      <div class="table-responsive">
      <table class="table table-sm mb-0">
        <thead><tr><th>職場</th><th class="text-end">回答者</th><th class="text-end">業務</th><th class="text-end">月間時間</th><th class="text-end">案件候補</th><th class="text-end">自主改善</th></tr></thead>
        <tbody>
        {% for w in workplaces %}
        <tr class="{% if w.respondents == 0 %}text-muted{% endif %}">
          <td>{{ w.workplace.name }}</td>
          <td class="text-end">{{ w.respondents }}</td>
          <td class="text-end">{{ w.tasks }}</td>
          <td class="text-end">{{ w.minutes|hours }} h</td>
          <td class="text-end">{{ w.candidates }}</td>
          <td class="text-end">{{ w.self_improvers }}</td>
        </tr>
        {% endfor %}
        </tbody>
      </table>
      </div>
    </div></div>
  </div>
</div>

<div class="row g-3 mb-4">
  <div class="col-lg-5">
    <div class="card h-100"><div class="card-body">
      <h5 class="card-title">改善の希望</h5>
      <div class="chart-box"><canvas id="chart-wish"></canvas></div>
    </div></div>
  </div>
  <div class="col-lg-7">
    <div class="card h-100"><div class="card-body">
      <h5 class="card-title">困りごと（ムリ・ムダ・ムラ）</h5>
      <div class="chart-box"><canvas id="chart-pains"></canvas></div>
    </div></div>
  </div>
</div>

<h2 class="section-title">次のアクション</h2>
<div class="row g-3">
  <div class="col-lg-5">
    <div class="card h-100"><div class="card-body">
      <h5 class="card-title">次のアクション提案</h5>
      <p class="text-muted small mb-2">蓄積された案件と業務情報から、次に調査・改善すべきテーマを提案します。</p>
      {% if suggestions %}
      <ul class="list-group list-group-flush">
        {% for s in suggestions %}
        <li class="list-group-item px-0">
          <span class="badge text-bg-secondary suggestion-kind">{{ s.kind }}</span>
          {% if s.href %}<a href="{{ s.href }}" class="fw-semibold">{{ s.title }}</a>{% else %}<span class="fw-semibold">{{ s.title }}</span>{% endif %}
          <div class="small text-muted">{{ s.detail }}</div>
        </li>
        {% endfor %}
      </ul>
      {% else %}
      <p class="text-muted">業務が登録されると提案が表示されます。</p>
      {% endif %}
      <hr>
      <div class="d-flex justify-content-between align-items-center">
        <h6 class="mb-0">AIによる提案</h6>
        <form method="post" action="{{ url_for('main.ai_suggest') }}">
          <button class="btn btn-sm btn-outline-primary" {% if not ai_enabled %}disabled title="ANTHROPIC_API_KEY を設定すると使えます"{% endif %}>AIで提案を生成</button>
        </form>
      </div>
      {% if ai_suggestions %}
      <ul class="list-group list-group-flush mt-2">
        {% for s in ai_suggestions %}
        <li class="list-group-item px-0">
          <span class="badge text-bg-primary suggestion-kind">{{ s.kind }}</span> <span class="fw-semibold">{{ s.title }}</span>
          <div class="small text-muted">{{ s.detail }}</div>
        </li>
        {% endfor %}
        <li class="list-group-item px-0 small text-muted">生成日時: {{ ai_suggestions[0].generated_at }}</li>
      </ul>
      {% elif not ai_enabled %}
      <p class="small text-muted mt-2 mb-0">AI連携は未設定です。config.py の ANTHROPIC_API_KEY を設定して起動すると、案件登録時のAI分析とここでの提案生成が有効になります。</p>
      {% endif %}
    </div></div>
  </div>
  <div class="col-lg-7">
    <div class="card h-100"><div class="card-body">
      <div class="d-flex justify-content-between align-items-center">
        <h5 class="card-title mb-0">最近登録された業務</h5>
        <a class="small" href="{{ url_for('cases.index', sort='new') }}">すべて見る</a>
      </div>
      <div class="table-responsive">
      <table class="table table-sm mt-2 mb-0">
        <thead><tr><th>業務名</th><th>職場</th><th>領域</th><th class="text-end">月間</th><th>希望</th><th>状態</th></tr></thead>
        <tbody>
        {% for t in recent %}
        <tr>
          <td><a href="{{ url_for('cases.detail', task_id=t.id) }}">{{ t.name }}</a></td>
          <td>{{ t.workplace.name if t.workplace else '' }}</td>
          <td>{{ t.domain.name if t.domain else '' }}</td>
          <td class="text-end">{{ t.monthly_minutes|hours }} h</td>
          <td>{{ t.wish or '' }}</td>
          <td><span class="badge text-bg-{{ {'受付':'secondary','対応中':'primary','完了':'success','保留':'warning'}.get(t.status, 'secondary') }}">{{ t.status }}</span></td>
        </tr>
        {% else %}
        <tr><td colspan="6" class="text-muted">まだ登録がありません</td></tr>
        {% endfor %}
        </tbody>
      </table>
      </div>
    </div></div>
  </div>
</div>
{% endblock %}

{% block scripts %}
<script id="chart-data" type="application/json">{{ charts_json|safe }}</script>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<script>
(function () {
  var data = JSON.parse(document.getElementById("chart-data").textContent);
  var C = { blue: "#0a84ff", green: "#30d158", orange: "#ff9f0a", purple: "#bf5af2", red: "#ff453a", teal: "#5ac8fa", indigo: "#5e5ce6", pink: "#ff375f", yellow: "#ffd60a", gray: "#8e8e93" };
  var palette = [C.blue, C.green, C.orange, C.purple, C.red, C.teal, C.indigo, C.yellow];
  Chart.defaults.font.family = getComputedStyle(document.body).fontFamily;
  Chart.defaults.font.size = 12;
  Chart.defaults.color = "#6e6e73";
  Chart.defaults.borderColor = "#ececf0";
  Chart.defaults.plugins.legend.labels.usePointStyle = true;
  Chart.defaults.plugins.legend.labels.boxWidth = 8;
  Chart.defaults.plugins.legend.labels.boxHeight = 8;
  Chart.defaults.plugins.tooltip.backgroundColor = "#1d1d1f";
  Chart.defaults.plugins.tooltip.padding = 10;
  Chart.defaults.plugins.tooltip.cornerRadius = 8;
  var axisX = { grid: { display: false }, border: { display: false } };
  var axisY = function (extra) { return Object.assign({ beginAtZero: true, grid: { color: "#ececf0" }, border: { display: false, dash: [4, 4] } }, extra || {}); };
  new Chart(document.getElementById("chart-wish"), {
    type: "doughnut",
    data: { labels: data.wish.labels, datasets: [{ data: data.wish.values, backgroundColor: palette, borderWidth: 2, borderColor: "#fff", hoverOffset: 6 }] },
    options: { responsive: true, maintainAspectRatio: false, cutout: "68%", plugins: { legend: { position: "bottom" } } }
  });
  new Chart(document.getElementById("chart-pains"), {
    type: "bar",
    data: { labels: data.pains.labels, datasets: [{ label: "件数", data: data.pains.values, backgroundColor: C.red, borderRadius: 6, borderSkipped: false, barPercentage: 0.65 }] },
    options: { indexAxis: "y", responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } },
      scales: { x: axisY({ ticks: { precision: 0 } }), y: { grid: { display: false }, border: { display: false } } } }
  });
  var mc = document.getElementById("chart-monthly");
  if (mc && data.monthly && data.monthly.labels.length) {
    new Chart(mc, {
      type: "line",
      data: { labels: data.monthly.labels, datasets: [
        { label: "繰り返し業務あり", data: data.monthly.has_rate, borderColor: C.red, backgroundColor: C.red, tension: 0.35, pointRadius: 3, borderWidth: 2.5 },
        { label: "減った", data: data.monthly.decreased_rate, borderColor: C.green, backgroundColor: C.green, tension: 0.35, pointRadius: 3, borderWidth: 2.5 },
        { label: "自分でやりたい", data: data.monthly.self_rate, borderColor: C.blue, backgroundColor: C.blue, tension: 0.35, pointRadius: 3, borderWidth: 2.5 }
      ] },
      options: { responsive: true, maintainAspectRatio: false,
        scales: { x: axisX, y: axisY({ max: 100, ticks: { callback: function (v) { return v + "%"; }, maxTicksLimit: 5 } }) },
        plugins: { legend: { display: false } } }
    });
  }
})();
</script>
{% endblock %}
''',
    'monthly/admin.html': r'''{% extends "base.html" %}
{% block title %}月次アンケートの設問{% endblock %}
{% block content %}
<nav aria-label="breadcrumb"><ol class="breadcrumb small mb-2">
  <li class="breadcrumb-item"><a href="{{ url_for('monthly.results', period=period) }}">月次アンケート 集計</a></li>
  <li class="breadcrumb-item active">設問の編集</li>
</ol></nav>
<div class="d-flex flex-wrap justify-content-between align-items-end gap-2 mb-3">
  <div>
    <h2 class="mb-0">月次アンケートの設問</h2>
    <div class="text-muted small">月度ごとに設問を決めます。基本設問（鍵マーク）は毎月の推移に使うため、文面だけ変えられます。選択肢と必須は固定です。</div>
  </div>
  <form method="get" class="d-flex gap-2 align-items-center">
    <input type="month" name="period" class="form-control form-control-sm" value="{{ period }}" onchange="this.form.submit()">
  </form>
</div>

<div class="row g-3">
  <div class="col-lg-8">
    <div class="card"><div class="card-body">
      <div class="d-flex flex-wrap justify-content-between align-items-center gap-2 mb-3">
        <h5 class="card-title mb-0">{{ period_label }}の設問 {% if response_count %}<span class="badge text-bg-secondary ms-1">回答 {{ response_count }} 件</span>{% endif %}</h5>
        <div class="d-flex gap-1 flex-wrap">
          <button type="button" class="btn btn-sm btn-outline-secondary" id="qe-copy-prev">前月（{{ prev_period }}）をコピー</button>
          <button type="button" class="btn btn-sm btn-outline-secondary" id="qe-reset">基本設問に戻す</button>
        </div>
      </div>
      {% if response_count %}<div class="alert alert-warning py-2 small">この月度には回答があります。設問を変えると、変更前の回答と対応が取れなくなることがあります。</div>{% endif %}
      <div id="qe-list"></div>
      <button type="button" class="btn btn-outline-primary w-100 mt-2" id="qe-add"><i class="bi bi-plus-lg me-1"></i>設問を追加</button>
      <hr>
      <div class="d-flex flex-wrap justify-content-between align-items-center gap-2">
        <div class="form-check form-switch mb-0">
          <input class="form-check-input" type="checkbox" id="qe-active" {% if survey.is_active %}checked{% endif %}>
          <label class="form-check-label" for="qe-active">この月度の回答を受け付ける</label>
        </div>
        <div class="d-flex align-items-center gap-2">
          <span id="qe-msg" class="small"></span>
          <button type="button" class="btn btn-primary" id="qe-save">保存する</button>
        </div>
      </div>
    </div></div>
  </div>
  <div class="col-lg-4">
    <div class="card"><div class="card-body">
      <h5 class="card-title">月度ごとの設問セット</h5>
      <table class="table table-sm mb-0">
        <thead><tr><th>月度</th><th class="text-end">設問</th><th class="text-end">回答</th><th>状態</th><th></th></tr></thead>
        <tbody>
        {% for s in surveys %}
        <tr class="{% if s.period == period %}table-primary{% endif %}">
          <td><a href="{{ url_for('monthly.survey_admin', period=s.period) }}">{{ s.period_label }}</a>{% if s.period == current_period %} <span class="badge text-bg-primary">今月</span>{% endif %}</td>
          <td class="text-end">{{ s.questions|length }}</td>
          <td class="text-end">{{ counts.get(s.period, 0) }}</td>
          <td>{% if s.is_active %}<span class="badge text-bg-success">受付中</span>{% else %}<span class="badge text-bg-secondary">停止</span>{% endif %}</td>
          <td class="text-end">
            {% if not counts.get(s.period) %}
            <form method="post" action="{{ url_for('monthly.survey_admin_delete', survey_id=s.id) }}" data-confirm="{{ s.period_label }}の設問セットを削除しますか？"><button class="btn btn-sm btn-outline-danger">削除</button></form>
            {% endif %}
          </td>
        </tr>
        {% endfor %}
        </tbody>
      </table>
      <p class="small text-muted mt-2 mb-0">月度の設問セットは、回答フォームが初めて開かれたときに直近の月度からコピーして自動で作られます。</p>
    </div></div>
  </div>
</div>

<template id="qe-item">
  <div class="qe-q card mb-2">
    <div class="card-body py-3">
      <div class="d-flex align-items-start gap-2">
        <div class="qe-handle text-muted small pt-2"><span class="qe-index"></span></div>
        <div class="flex-grow-1">
          <div class="row g-2">
            <div class="col-md-8"><input type="text" class="form-control qe-text" placeholder="設問の文章" maxlength="200"></div>
            <div class="col-6 col-md-2">
              <select class="form-select qe-type"><option value="choice">選択（1つ）</option><option value="text">記述</option></select>
            </div>
            <div class="col-6 col-md-2 d-flex align-items-center">
              <div class="form-check mb-0"><input class="form-check-input qe-required" type="checkbox"><label class="form-check-label small">必須</label></div>
            </div>
          </div>
          <div class="qe-options mt-2">
            <label class="form-label small mb-1">選択肢（1行に1つ、2〜7個。上から順に「良い→悪い」などの順で）</label>
            <textarea class="form-control form-control-sm qe-opts" rows="3"></textarea>
          </div>
          <div class="qe-placeholder mt-2 d-none">
            <input type="text" class="form-control form-control-sm qe-ph" placeholder="記入欄の例文（任意）" maxlength="100">
          </div>
          <div class="small text-muted mt-1 qe-core d-none"><i class="bi bi-lock-fill me-1"></i>基本設問（毎月の推移に使います。選択肢の変更と削除はできません）</div>
        </div>
        <div class="btn-group-vertical btn-group-sm">
          <button type="button" class="btn btn-outline-secondary qe-up" title="上へ">▲</button>
          <button type="button" class="btn btn-outline-secondary qe-down" title="下へ">▼</button>
          <button type="button" class="btn btn-outline-danger qe-del" title="削除">×</button>
        </div>
      </div>
    </div>
  </div>
</template>
{% endblock %}
{% block scripts %}
<script id="qe-data" type="application/json">{{ questions_json|safe }}</script>
<script>
(function () {
  var CORE = {{ core_keys|tojson }};
  var LABELS = {{ C.MONTHLY_KPI_LABELS|tojson }};
  var list = document.getElementById("qe-list");
  var tpl = document.getElementById("qe-item");
  var questions = JSON.parse(document.getElementById("qe-data").textContent);
  var msg = document.getElementById("qe-msg");

  function render() {
    list.innerHTML = "";
    questions.forEach(function (q, i) {
      var node = tpl.content.firstElementChild.cloneNode(true);
      node.dataset.index = i;
      node.querySelector(".qe-index").textContent = (i + 1) + ".";
      node.querySelector(".qe-text").value = q.text || "";
      node.querySelector(".qe-type").value = q.type || "choice";
      node.querySelector(".qe-required").checked = !!q.required;
      node.querySelector(".qe-opts").value = (q.options || []).join("\n");
      node.querySelector(".qe-ph").value = q.placeholder || "";
      var isCore = CORE.indexOf(q.key) >= 0;
      if (isCore) {
        node.querySelector(".qe-core").classList.remove("d-none");
        node.querySelector(".qe-core").insertAdjacentText("beforeend", LABELS[q.key] ? "　→ " + LABELS[q.key] : "");
        node.querySelector(".qe-del").disabled = true;
        node.querySelector(".qe-type").disabled = true;
        node.querySelector(".qe-opts").disabled = true;
        node.querySelector(".qe-required").disabled = true;
      }
      if (q.depends && q.depends.key) {
        var parentQ = questions.filter(function (x) { return x.key === q.depends.key; })[0];
        var cond = (q.depends.show_if || q.depends.required_if || []).join("／");
        var coreNote = node.querySelector(".qe-core");
        coreNote.classList.remove("d-none");
        coreNote.insertAdjacentText("beforeend", "　条件つき: 「" + (parentQ ? parentQ.text : q.depends.key) + "」が「" + cond + "」のときだけ出します（出たら必須）");
      }
      toggleType(node, q.type);
      node.querySelector(".qe-type").addEventListener("change", function (e) { q.type = e.target.value; toggleType(node, q.type); });
      node.querySelector(".qe-up").addEventListener("click", function () { if (i > 0) { sync(); questions.splice(i - 1, 0, questions.splice(i, 1)[0]); render(); } });
      node.querySelector(".qe-down").addEventListener("click", function () { if (i < questions.length - 1) { sync(); questions.splice(i + 1, 0, questions.splice(i, 1)[0]); render(); } });
      node.querySelector(".qe-del").addEventListener("click", function () { if (isCore) return; sync(); questions.splice(i, 1); render(); });
      list.appendChild(node);
    });
  }
  function toggleType(node, type) {
    node.querySelector(".qe-options").classList.toggle("d-none", type === "text");
    node.querySelector(".qe-placeholder").classList.toggle("d-none", type !== "text");
  }
  function sync() {
    list.querySelectorAll(".qe-q").forEach(function (node, i) {
      var q = questions[i];
      q.text = node.querySelector(".qe-text").value.trim();
      q.type = node.querySelector(".qe-type").value;
      q.required = node.querySelector(".qe-required").checked;
      q.options = node.querySelector(".qe-opts").value.split("\n").map(function (s) { return s.trim(); }).filter(Boolean);
      q.placeholder = node.querySelector(".qe-ph").value.trim();
    });
  }
  function note(text, ok) { msg.textContent = text; msg.className = "small " + (ok ? "text-success" : "text-danger"); }

  document.getElementById("qe-add").addEventListener("click", function () {
    sync(); questions.push({ key: "", type: "choice", text: "", options: ["そう思う", "どちらともいえない", "そう思わない"], required: true }); render();
    var last = list.lastElementChild; if (last) { last.querySelector(".qe-text").focus(); last.scrollIntoView({ behavior: "smooth", block: "center" }); }
  });
  async function load(which, label) {
    var r = await fetch("{{ url_for('monthly.survey_admin_questions') }}?period=" + encodeURIComponent(which), { headers: { Accept: "application/json" } });
    var j = await r.json();
    if (!j.questions) { note(j.message || "読み込めませんでした", false); return; }
    questions = j.questions; render(); note(label + "を読み込みました。保存ボタンで確定します。", true);
  }
  document.getElementById("qe-copy-prev").addEventListener("click", function () { load("{{ prev_period }}", "前月の設問"); });
  document.getElementById("qe-reset").addEventListener("click", function () { if (confirm("基本設問に戻しますか？追加した設問は消えます（保存するまで確定しません）。")) load("default", "基本設問"); });
  document.getElementById("qe-save").addEventListener("click", async function () {
    sync();
    var btn = this; btn.disabled = true; note("保存中…", true);
    try {
      var r = await fetch("{{ url_for('monthly.survey_admin_save') }}", { method: "POST", headers: { "Content-Type": "application/json", Accept: "application/json" },
        body: JSON.stringify({ period: "{{ period }}", questions: questions, is_active: document.getElementById("qe-active").checked }) });
      var j = await r.json();
      note(j.message || (j.success ? "保存しました" : "保存できませんでした"), !!j.success);
      if (j.success) setTimeout(function () { window.location.reload(); }, 600);
    } catch (e) { note("保存できませんでした。通信状態を確認してください。", false); }
    btn.disabled = false;
  });
  render();
})();
</script>
{% endblock %}
''',
    'monthly/done.html': r'''{% extends "base.html" %}
{% block title %}回答ありがとうございました{% endblock %}
{% block content %}
<div class="mq-wrap">
  <div class="card mq-card mq-done"><div class="card-body text-center">
    <div class="mq-check"><i class="bi bi-check-lg"></i></div>
    <h2 class="h4 mb-1">回答を受け付けました</h2>
    <p class="text-muted mb-3">{{ period_label }}のアンケートにご協力ありがとうございました。</p>
    {% if row and row.answers.get('wants') == 'ある' %}
    <a class="btn btn-outline-primary" href="{{ url_for('survey.start', emp=row.emp_id) }}">なくしたい業務を自動化依頼に出す</a>
    {% endif %}
  </div></div>
  <p class="text-center small mt-3"><a href="{{ url_for('board.index') }}">活動状況へ</a></p>
</div>
{% endblock %}
''',
    'monthly/form.html': r'''{% extends "base.html" %}
{% block title %}月次アンケート {{ period_label }}{% endblock %}
{% block content %}
<div class="mq-wrap">
  <div class="mq-head">
    <span class="mq-badge">{{ period_label }}</span>
    <h1 class="mq-title">今月のアンケート</h1>
    <p class="mq-sub">所要時間は４分です。回答は改善推進のためだけに使い評価には使いません。</p>
  </div>

  {% if not survey.is_active %}
  <div class="card mq-card"><div class="card-body text-center text-muted py-5">この月度のアンケートは準備中です。</div></div>
  {% else %}
  <div class="card mq-card mq-done {% if not existing %}d-none{% endif %}" id="mq-answered"><div class="card-body text-center">
    <div class="mq-check"><i class="bi bi-check-lg"></i></div>
    <h2 class="h4 mb-1">回答済み</h2>
    <p class="text-muted mb-0">{{ period_label }}のアンケートはすでに回答いただいています。ご協力ありがとうございました。</p>
  </div></div>

  <form id="mq-form" class="card mq-card {% if existing %}d-none{% endif %}" method="post" action="{{ url_for('monthly.submit') }}" novalidate>
    <div class="card-body">
      <input type="hidden" name="period" value="{{ period }}">
      <div class="mq-identity">
        <div class="row g-2">
          <div class="col-sm-6">
            <label class="form-label" for="emp_id">社員ID</label>
            <input type="text" name="emp_id" id="emp_id" class="form-control" value="{{ emp_id }}" autocomplete="off" required>
          </div>
          <div class="col-sm-6">
            <label class="form-label" for="workplace_id">職場</label>
            <select name="workplace_id" id="workplace_id" class="form-select" required>
              <option value="">選択</option>
              {% for w in workplaces %}<option value="{{ w.id }}" {% if info and info.workplace_id == w.id %}selected{% endif %}>{{ w.name }}</option>{% endfor %}
            </select>
          </div>
          <div class="col-sm-6 d-none" id="mq-fallback">
            <label class="form-label" for="name">氏名</label>
            <input type="text" name="name" id="name" class="form-control" value="{{ info.name if info else '' }}">
          </div>
        </div>
        <div id="emp-msg" class="form-text"></div>
      </div>

      {% set ns = namespace(main=0, nos={}, subs={}) %}
      {% for q in questions %}
      {% if q.depends %}
        {% set _ = ns.subs.update({q.depends.key: ns.subs.get(q.depends.key, 0) + 1}) %}
        {% set label = ns.nos.get(q.depends.key, '?') ~ '-' ~ ns.subs[q.depends.key] %}
      {% else %}
        {% set ns.main = ns.main + 1 %}
        {% set _ = ns.nos.update({q.key: ns.main}) %}
        {% set label = ns.main %}
      {% endif %}
      <div class="mq-q{% if q.depends %} mq-sub d-none{% endif %}" data-key="{{ q.key }}" data-type="{{ q.type }}" data-required="{{ 1 if q.required else 0 }}"{% if q.depends %} data-dep-key="{{ q.depends.key }}" data-dep-show='{{ (q.depends.show_if or q.depends.required_if or [])|tojson }}' data-dep-required="{{ 1 if (q.depends.required or q.depends.required_if) else 0 }}"{% endif %}>
        <div class="mq-q-head">
          <span class="mq-no">{{ label }}</span>
          <span class="mq-text">{{ q.text }}</span>
          {% if q.depends %}<span class="mq-required d-none" data-badge="req">必須</span><span class="mq-optional d-none" data-badge="opt">任意</span>{% elif not q.required %}<span class="mq-optional">任意</span>{% endif %}
        </div>
        {% if q.type == 'choice' %}
        <div class="mq-options">
          {% for opt in q.options %}
          <label class="mq-option"><input type="radio" name="ans-{{ q.key }}" value="{{ opt }}"><span>{{ opt }}</span></label>
          {% endfor %}
        </div>
        {% else %}
        <textarea name="ans-{{ q.key }}" class="form-control" rows="2" placeholder="{{ q.placeholder or '' }}"></textarea>
        {% endif %}
        <div class="mq-err small text-danger"></div>
      </div>
      {% endfor %}

      <div class="mq-q">
        <div class="mq-q-head"><span class="mq-no"><i class="bi bi-chat-left-text"></i></span><span class="mq-text">ご意見・気づいたこと</span><span class="mq-optional">任意</span></div>
        <textarea name="comment" class="form-control" rows="3" maxlength="500" placeholder="困っていること、やってみたいこと、何でも"></textarea>
      </div>

      <div id="mq-error" class="alert alert-danger py-2 small d-none"></div>
      <button class="btn btn-primary btn-lg w-100" id="mq-submit">回答を送信する</button>
    </div>
  </form>

  <div class="card mq-card mq-done d-none" id="mq-success"><div class="card-body text-center">
    <div class="mq-check"><i class="bi bi-check-lg"></i></div>
    <h2 class="h4 mb-1">回答を受け付けました</h2>
    <p class="text-muted" id="mq-success-msg"></p>
    <a class="btn btn-outline-primary d-none" id="mq-register">なくしたい業務を自動化依頼に出す</a>
  </div></div>
  {% endif %}
</div>
{% endblock %}
{% block scripts %}
{% if survey.is_active %}
<script>
(function () {
  var form = document.getElementById("mq-form");
  var answered = document.getElementById("mq-answered");
  var empInput = document.getElementById("emp_id");
  var nameInput = document.getElementById("name");
  var wpSelect = document.getElementById("workplace_id");
  var fallback = document.getElementById("mq-fallback");
  var msg = document.getElementById("emp-msg");
  var lookupTimer = null;
  async function lookupEmp() {
    var id = empInput.value.trim();
    if (!id) { fallback.classList.add("d-none"); msg.textContent = ""; return; }
    var r = await fetch("{{ url_for('monthly.api_employee', emp_id='') }}" + encodeURIComponent(id) + "?period={{ period }}", { headers: { Accept: "application/json" } });
    var j = await r.json();
    if (empInput.value.trim() !== id) return;
    if (j.answered) {
      form.classList.add("d-none");
      answered.classList.remove("d-none");
      return;
    }
    answered.classList.add("d-none");
    form.classList.remove("d-none");
    if (j.found) {
      nameInput.value = j.name || "";
      if (j.workplace_id) wpSelect.value = String(j.workplace_id);
      fallback.classList.toggle("d-none", !!j.name);
      var noName = !j.name, noWp = !wpSelect.value;
      msg.textContent = noName && noWp ? "氏名と職場を入れてください（初回だけ）" : noName ? "氏名を入れてください（初回だけ）" : noWp ? "職場を選んでください（初回だけ）" : "";
    } else {
      nameInput.value = "";
      fallback.classList.remove("d-none");
      msg.textContent = "見つかりません。氏名と職場を入れてください";
    }
  }
  // 条件つきの設問: 親の答えで「必須」になったり、入力欄ごと消えたりする
  function applyDeps() {
    form.querySelectorAll(".mq-q[data-dep-key]").forEach(function (q) {
      var parent = form.querySelector('.mq-q[data-key="' + q.dataset.depKey + '"] input:checked');
      var val = parent ? parent.value : "";
      var shown = JSON.parse(q.dataset.depShow || "[]").indexOf(val) >= 0;
      var required = shown && q.dataset.depRequired === "1";
      q.classList.toggle("d-none", !shown);
      q.dataset.required = required ? "1" : "0";
      if (!shown) {
        var t = q.querySelector("textarea");
        if (t) t.value = "";
        q.querySelectorAll('input[type="radio"]').forEach(function (i) { i.checked = false; });
      }
      var req = q.querySelector('[data-badge="req"]'), opt = q.querySelector('[data-badge="opt"]');
      if (req) req.classList.toggle("d-none", !required);
      if (opt) opt.classList.toggle("d-none", required || !shown);
      var err = q.querySelector(".mq-err");
      if (err && !shown) err.textContent = "";
    });
  }
  form.addEventListener("change", function (e) { if (e.target.matches('input[type="radio"]')) applyDeps(); });
  applyDeps();
  empInput.addEventListener("input", function () { clearTimeout(lookupTimer); lookupTimer = setTimeout(lookupEmp, 500); });
  empInput.addEventListener("change", lookupEmp);
  {% if emp_id and not existing %}lookupEmp();{% endif %}
  var submitBtn = document.getElementById("mq-submit");
  var errBox = document.getElementById("mq-error");
  form.addEventListener("submit", async function (e) {
    e.preventDefault();
    errBox.classList.add("d-none");
    var answers = {}, firstMissing = null;
    form.querySelectorAll(".mq-q[data-key]").forEach(function (q) {
      var key = q.dataset.key, val = "";
      if (q.classList.contains("d-none")) { answers[key] = ""; return; }
      if (q.dataset.type === "choice") { var c = q.querySelector("input:checked"); val = c ? c.value : ""; }
      else { var t = q.querySelector("textarea"); val = t ? t.value.trim() : ""; }
      var err = q.querySelector(".mq-err"); err.textContent = "";
      if (q.dataset.required === "1" && !val) { err.textContent = q.dataset.type === "choice" ? "選択してください" : "入力してください"; if (!firstMissing) firstMissing = q; }
      answers[key] = val;
    });
    var missing = !form.emp_id.value.trim() ? "社員IDを入力してください" : !form.name.value.trim() ? "氏名を入力してください" : !form.workplace_id.value ? "職場を選んでください" : "";
    if (missing) {
      if (form.emp_id.value.trim() && !form.name.value.trim()) fallback.classList.remove("d-none");
      errBox.textContent = missing; errBox.classList.remove("d-none");
      document.querySelector(".mq-identity").scrollIntoView({ behavior: "smooth", block: "center" });
      return;
    }
    if (firstMissing) { firstMissing.scrollIntoView({ behavior: "smooth", block: "center" }); return; }
    var payload = { period: form.period.value, emp_id: form.emp_id.value.trim(), name: form.name.value.trim(), workplace_id: form.workplace_id.value, answers: answers, comment: form.comment.value.trim() };
    submitBtn.disabled = true; submitBtn.textContent = "送信中…";
    try {
      var r = await fetch(form.action, { method: "POST", headers: { "Content-Type": "application/json", Accept: "application/json" }, body: JSON.stringify(payload) });
      var j = await r.json();
      if (j.success) {
        form.classList.add("d-none");
        document.getElementById("mq-success-msg").textContent = j.message;
        if (j.wants) { var a = document.getElementById("mq-register"); a.href = j.register_url; a.classList.remove("d-none"); }
        document.getElementById("mq-success").classList.remove("d-none");
        window.scrollTo({ top: 0, behavior: "smooth" });
      } else {
        errBox.textContent = j.message || "送信できませんでした"; errBox.classList.remove("d-none");
        if (j.errors) Object.keys(j.errors).forEach(function (k) { var el = form.querySelector('.mq-q[data-key="' + k + '"] .mq-err'); if (el) el.textContent = j.errors[k]; });
        submitBtn.disabled = false; submitBtn.textContent = "回答を送信する";
      }
    } catch (err) {
      errBox.textContent = "送信できませんでした。通信状態を確認してください。"; errBox.classList.remove("d-none");
      submitBtn.disabled = false; submitBtn.textContent = "回答を送信する";
    }
  });
})();
</script>
{% endif %}
{% endblock %}
''',
    'monthly/results.html': r'''{% extends "base.html" %}
{% block title %}月次アンケート 集計{% endblock %}
{% block content %}
<div class="d-flex flex-wrap justify-content-between align-items-end gap-2 mb-1">
  <div>
    <h2 class="mb-0">月次アンケート 集計</h2>
    <div class="text-muted small">毎月同じ基本設問で、変化を追います。</div>
  </div>
  <div class="d-flex gap-2 flex-wrap">
    <a class="btn btn-outline-secondary btn-sm" href="{{ url_for('monthly.export_csv') }}">CSV</a>
    <a class="btn btn-outline-primary btn-sm" href="{{ url_for('monthly.survey_admin', period=selected) }}">設問を編集</a>
  </div>
</div>

{% if rows %}
<h2 class="section-title">推移 <span class="section-sub">基本設問の割合と回答者数</span></h2>
<div class="row g-3 mb-2">
  <div class="col-lg-7">
    <div class="card h-100"><div class="card-body">
      <h5 class="card-title">割合の推移（%）</h5>
      <div class="chart-box" style="height: 280px"><canvas id="chart-rates"></canvas></div>
    </div></div>
  </div>
  <div class="col-lg-5">
    <div class="card h-100"><div class="card-body">
      <h5 class="card-title">回答者数と推計月間時間</h5>
      <div class="chart-box" style="height: 280px"><canvas id="chart-volume"></canvas></div>
    </div></div>
  </div>
  <div class="col-12">
    <div class="card"><div class="card-body">
      <div class="table-responsive">
      <table class="table table-sm align-middle mb-0">
        <thead><tr>
          <th>月度</th><th class="text-end">回答者</th><th class="text-end">回答率</th><th class="text-end">繰り返し業務あり</th><th class="text-end">推計 月間時間</th>
          <th class="text-end">減った</th><th class="text-end">増えた</th><th class="text-end">自動化したい</th><th class="text-end">依頼したい</th><th class="text-end">自分でやりたい</th><th class="text-end">新規業務</th><th class="text-end">コメント</th>
        </tr></thead>
        <tbody>
        {% for r in rows %}
        <tr class="{% if r.period == selected %}table-primary{% endif %}">
          <td><a href="{{ url_for('monthly.results', period=r.period) }}#detail">{{ r.label }}</a></td>
          <td class="text-end">{{ r.n }}</td>
          <td class="text-end">{% if r.response_rate is not none %}{{ '%.0f'|format(r.response_rate) }}%{% else %}－{% endif %}</td>
          <td class="text-end">{{ '%.0f'|format(r.has_rate) }}%</td>
          <td class="text-end">{{ '%.0f'|format(r.hours_month) }} h</td>
          <td class="text-end">{{ '%.0f'|format(r.decreased_rate) }}%</td>
          <td class="text-end">{{ '%.0f'|format(r.increased_rate) }}%</td>
          <td class="text-end">{{ '%.0f'|format(r.wants_rate) }}%</td>
          <td class="text-end">{{ '%.0f'|format(r.ask_rate) }}%</td>
          <td class="text-end">{{ '%.0f'|format(r.self_rate) }}%</td>
          <td class="text-end">{{ r.new_count }}</td>
          <td class="text-end">{{ r.comment_count }}</td>
        </tr>
        {% endfor %}
        </tbody>
      </table>
      </div>
      <p class="small text-muted mt-2 mb-0">回答率の分母は社員マスタの有効人数（未登録なら管理画面の対象者数）。推計月間時間 ＝ 1シフトあたり時間の換算値 × {{ C.SHIFTS_PER_MONTH }}（1か月の出勤日数の目安）。</p>
    </div></div>
  </div>
</div>

<h2 class="section-title">設問ごとの推移 <span class="section-sub">選択肢の割合（%）が月度ごとにどう動いたか</span></h2>
<div class="row g-3 mb-2">
  {% for t in trends %}
  <div class="col-lg-6">
    <div class="card h-100"><div class="card-body">
      <h5 class="card-title mb-0">{{ t.q.text }}</h5>
      <div class="small text-muted mb-2">{{ t.label }}</div>
      <div class="chart-box" style="height: 200px"><canvas id="chart-q-{{ t.q.key }}"></canvas></div>
      <div class="table-responsive mt-2">
      <table class="table table-sm trend-table mb-0">
        <thead><tr><th></th>{% for l in t.labels %}<th class="text-end">{{ l|replace('年', '/')|replace('月度', '') }}</th>{% endfor %}</tr></thead>
        <tbody>
        {% for s in t.series %}
        <tr><td>{{ s.option }}</td>{% for v in s.percents %}<td class="text-end">{{ v }}%</td>{% endfor %}</tr>
        {% endfor %}
        <tr class="text-muted"><td>回答数</td>{% for n in t.n %}<td class="text-end">{{ n }}</td>{% endfor %}</tr>
        </tbody>
      </table>
      </div>
    </div></div>
  </div>
  {% endfor %}
</div>

<h2 class="section-title">職場別の推移 <span class="section-sub">繰り返し業務あり率と推計月間時間</span></h2>
<div class="card mb-2"><div class="card-body">
  <div class="table-responsive">
  <table class="table table-sm trend-table mb-0">
    <thead><tr><th>職場</th>{% for l in wp_trends.labels %}<th class="text-end">{{ l|replace('年', '/')|replace('月度', '') }}</th>{% endfor %}</tr></thead>
    <tbody>
    {% for row in wp_trends.rows %}
    <tr>
      <td>{{ row.workplace.name }}</td>
      {% for c in row.cells %}
      <td class="text-end">{% if c.n %}<div>{{ '%.0f'|format(c.has_rate) }}%</div><div class="cell-sub">{{ '%.0f'|format(c.hours_month) }} h・{{ c.n }}人</div>{% else %}<span class="text-muted">－</span>{% endif %}</td>
      {% endfor %}
    </tr>
    {% else %}
    <tr><td colspan="{{ 1 + wp_trends.labels|length }}" class="text-muted">回答はまだありません</td></tr>
    {% endfor %}
    </tbody>
  </table>
  </div>
  <p class="small text-muted mt-2 mb-0">上段が繰り返し業務あり率、下段が推計月間時間と回答者数です。</p>
</div></div>
{% endif %}

<h2 class="section-title" id="detail">月度の詳細
  <form method="get" class="d-inline-flex align-items-center ms-1">
    <select name="period" class="form-select form-select-sm" onchange="this.form.submit()" style="width: auto">
      {% for s in surveys %}<option value="{{ s.period }}" {% if s.period == selected %}selected{% endif %}>{{ s.period_label }}</option>{% endfor %}
      {% if not surveys %}<option value="{{ selected }}">{{ selected_label }}</option>{% endif %}
    </select>
  </form>
  <span class="section-sub">{% if summary %}回答 {{ summary.n }} 人{% if summary.response_rate is not none %}（回答率 {{ '%.0f'|format(summary.response_rate) }}%）{% endif %}{% else %}回答はまだありません{% endif %}</span>
</h2>
{% if survey %}
<div class="row g-3 mb-3">
  {% for s in stats %}
  <div class="col-lg-6">
    <div class="card h-100"><div class="card-body">
      <div class="mq-stat-head"><span class="mq-no">{{ loop.index }}</span><span class="fw-semibold">{{ s.q.text }}</span></div>
      {% if s.q.type == 'choice' %}
        {% if s.n %}
        <div class="mq-bars">
          {% for opt in s.q.options %}
          <div class="mq-bar-row">
            <div class="mq-bar-label">{{ opt }}</div>
            <div class="mq-bar"><span style="width: {{ s.percents[loop.index0] }}%"></span></div>
            <div class="mq-bar-val">{{ s.counts[loop.index0] }}<span class="text-muted"> 人・{{ s.percents[loop.index0] }}%</span></div>
          </div>
          {% endfor %}
        </div>
        {% else %}<p class="text-muted small mb-0">回答はまだありません</p>{% endif %}
      {% else %}
        {% if s['items'] %}
        <ul class="mq-list">
          {% for name, wp, text in s['items'] %}<li><span class="text-muted small">{{ wp }} {{ name }}</span><div>{{ text|replace('\n', '／') }}</div></li>{% endfor %}
        </ul>
        {% else %}<p class="text-muted small mb-0">記入はありません</p>{% endif %}
      {% endif %}
    </div></div>
  </div>
  {% endfor %}
  <div class="col-lg-6">
    <div class="card h-100"><div class="card-body">
      <div class="mq-stat-head"><span class="mq-no"><i class="bi bi-chat-left-text"></i></span><span class="fw-semibold">コメント</span><span class="text-muted small ms-1">{{ comments|length }} 件</span></div>
      {% if comments %}
      <ul class="mq-list">
        {% for r in comments %}<li><span class="text-muted small">{{ r.workplace.name if r.workplace else '' }} {{ r.name }}</span><div>{{ r.comment }}</div></li>{% endfor %}
      </ul>
      {% else %}<p class="text-muted small mb-0">コメントはありません</p>{% endif %}
    </div></div>
  </div>
</div>

<div class="card"><div class="card-body">
  <h5 class="card-title">回答一覧（{{ responses|length }} 件）</h5>
  <div class="table-responsive">
  <table class="table table-sm mb-0">
    <thead><tr><th>社員ID</th><th>氏名</th><th>職場</th>{% for q in survey.questions if q.type == 'choice' %}<th class="small">{{ q.text|truncate(14, True, '…') }}</th>{% endfor %}<th>回答日</th></tr></thead>
    <tbody>
    {% for r in responses %}
    <tr>
      <td class="small">{{ r.emp_id }}</td><td>{{ r.name }}</td><td class="small">{{ r.workplace.name if r.workplace else '' }}</td>
      {% for q in survey.questions if q.type == 'choice' %}<td class="small">{{ r.answers.get(q.key, '') }}</td>{% endfor %}
      <td class="small text-muted">{{ r.created_at|jpdate('%m/%d') }}</td>
    </tr>
    {% else %}
    <tr><td colspan="{{ 4 + choice_keys|length }}" class="text-muted">回答はありません</td></tr>
    {% endfor %}
    </tbody>
  </table>
  </div>
</div></div>
{% else %}
<div class="alert alert-info">この月度の設問セットはまだありません。<a href="{{ url_for('monthly.survey_admin', period=selected) }}">設問を編集</a>で作成すると回答を受け付けられます。</div>
{% endif %}
{% endblock %}
{% block scripts %}
{% if rows %}
<script id="monthly-chart" type="application/json">{{ chart_json|safe }}</script>
<script id="monthly-trends" type="application/json">{{ trends_json|safe }}</script>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<script>
(function () {
  var d = JSON.parse(document.getElementById("monthly-chart").textContent);
  var trends = JSON.parse(document.getElementById("monthly-trends").textContent);
  Chart.defaults.font.family = getComputedStyle(document.body).fontFamily;
  Chart.defaults.color = "#6e6e73";
  Chart.defaults.plugins.legend.labels.usePointStyle = true;
  Chart.defaults.plugins.legend.labels.boxWidth = 8;
  var line = function (label, data, color) { return { label: label, data: data, borderColor: color, backgroundColor: color, tension: 0.35, pointRadius: 3, borderWidth: 2.5 }; };
  new Chart(document.getElementById("chart-rates"), {
    type: "line",
    data: { labels: d.labels, datasets: [
      line("繰り返し業務あり", d.has_rate, "#ff453a"), line("減った", d.decreased_rate, "#30d158"), line("自動化したい", d.wants_rate, "#ff9f0a"),
      line("改善チームに依頼したい", d.ask_rate, "#bf5af2"), line("自分でやりたい", d.self_rate, "#0a84ff")
    ] },
    options: { responsive: true, maintainAspectRatio: false, scales: { x: { grid: { display: false }, border: { display: false } }, y: { beginAtZero: true, max: 100, grid: { color: "#ececf0" }, border: { display: false }, ticks: { callback: function (v) { return v + "%"; } } } }, plugins: { legend: { position: "bottom" } } }
  });
  new Chart(document.getElementById("chart-volume"), {
    type: "bar",
    data: { labels: d.labels, datasets: [
      { label: "回答者数", data: d.n, backgroundColor: "#0a84ff", yAxisID: "y", borderRadius: 6, borderSkipped: false, barPercentage: 0.6 },
      { label: "推計 月間時間 (h)", data: d.hours_month, backgroundColor: "#ff9f0a", yAxisID: "y1", borderRadius: 6, borderSkipped: false, barPercentage: 0.6 }
    ] },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: "bottom" } },
      scales: { x: { grid: { display: false }, border: { display: false } }, y: { beginAtZero: true, position: "left", grid: { color: "#ececf0" }, border: { display: false } }, y1: { beginAtZero: true, position: "right", grid: { display: false }, border: { display: false } } } }
  });
  var palette = ["#0a84ff", "#4da3ff", "#8ec3ff", "#b5d6ff", "#d2e5fb", "#e4edf8", "#eef3fa"];
  trends.forEach(function (q) {
    var el = document.getElementById("chart-q-" + q.key);
    if (!el) return;
    new Chart(el, {
      type: "bar",
      data: { labels: q.labels, datasets: q.options.map(function (o, i) { return { label: o, data: q.percents[i], backgroundColor: palette[i % palette.length], stack: "s", borderRadius: 3, borderSkipped: false, barPercentage: 0.55 }; }) },
      options: { responsive: true, maintainAspectRatio: false,
        plugins: { legend: { position: "bottom" }, tooltip: { callbacks: { label: function (c) { return c.dataset.label + ": " + c.parsed.y + "%"; } } } },
        scales: { x: { stacked: true, grid: { display: false }, border: { display: false } }, y: { stacked: true, beginAtZero: true, suggestedMax: 100, ticks: { callback: function (v) { return v + "%"; } }, grid: { color: "#ececf0" }, border: { display: false } } } }
    });
  });
})();
</script>
{% endif %}
{% endblock %}
''',
    'people/index.html': r'''{% extends "base.html" %}
{% block title %}DX人財情報{% endblock %}
{% block content %}
<h2 class="mb-1">DX人財情報</h2>
<p class="text-muted">自動化依頼の「改善への関わり方」と月次アンケートの回答から、一人ひとりの意欲・使える技術・学びたいことをまとめます。</p>

<div class="row g-3 mb-3">
  <div class="col-6 col-lg-3"><div class="card kpi-card h-100"><div class="card-body py-2">
    <div class="kpi-label">回答した人</div>
    <div class="kpi-value fs-3">{{ summary.total }}<span class="fs-6 text-muted"> 人</span></div>
    <div class="kpi-sub">自動化依頼または月次アンケート</div>
  </div></div></div>
  <div class="col-6 col-lg-3"><div class="card kpi-card h-100"><div class="card-body py-2">
    <div class="kpi-label">自分で改善したい</div>
    <div class="kpi-value fs-3">{{ summary.positive }}<span class="fs-6 text-muted"> 人</span></div>
    <div class="kpi-sub">関わり方・業務ごとの回答・月次のいずれかで「自分で取り組みたい」</div>
  </div></div></div>
  <div class="col-6 col-lg-3"><div class="card kpi-card h-100"><div class="card-body py-2">
    <div class="kpi-label">直近の月次で「自分で取り組みたい」</div>
    <div class="kpi-value fs-3">{{ summary.self_now }}<span class="fs-6 text-muted"> 人</span></div>
    <div class="kpi-sub">{{ summary.period_label }}</div>
  </div></div></div>
  <div class="col-6 col-lg-3"><div class="card kpi-card h-100"><div class="card-body py-2">
    <div class="kpi-label">事例の発表・共有ができる</div>
    <div class="kpi-value fs-3">{{ summary.share }}<span class="fs-6 text-muted"> 人</span></div>
    <div class="kpi-sub">関わり方の「発表・共有」で「できる」</div>
  </div></div></div>
</div>

<form class="card mb-3" method="get"><div class="card-body d-flex flex-wrap align-items-end gap-3 py-3">
  <div>
    <label class="form-label small mb-1" for="dx-wp">職場</label>
    <select name="workplace_id" id="dx-wp" class="form-select form-select-sm">
      <option value="">すべて</option>
      {% for w in workplaces %}<option value="{{ w.id }}" {% if args.workplace_id == w.id %}selected{% endif %}>{{ w.name }}</option>{% endfor %}
    </select>
  </div>
  <div>
    <label class="form-label small mb-1" for="dx-wish">意欲</label>
    <select name="wish" id="dx-wish" class="form-select form-select-sm">
      <option value="">すべて</option>
      <option value="positive" {% if args.wish == 'positive' %}selected{% endif %}>自分で改善したい人</option>
      {% for opt in C.C1_OPTIONS %}<option value="{{ opt }}" {% if args.wish == opt %}selected{% endif %}>{{ opt }}</option>{% endfor %}
    </select>
  </div>
  <button class="btn btn-sm btn-primary">表示</button>
  <span class="small text-muted ms-auto">{{ people|length }} 人</span>
</div></form>

<div class="card mb-3"><div class="card-body">
  <h5 class="card-title">人財一覧</h5>
  <div class="table-responsive">
  <table class="table table-sm align-middle mb-0 dx-table">
    <thead><tr>
      <th>氏名</th><th>職場</th><th>意欲（自動化依頼）</th><th>月次：改善の進め方</th>
      <th>使える技術</th><th>学びたい技術</th><th>割ける時間</th><th>発表・共有</th><th>業務</th>
    </tr></thead>
    <tbody>
    {% for p in people %}
    <tr>
      <td class="text-nowrap">{{ p.name }}{% if p.emp_id %}<div class="small text-muted">{{ p.emp_id }}</div>{% endif %}</td>
      <td class="small text-nowrap">{{ p.workplace.name if p.workplace else '' }}</td>
      <td class="small text-nowrap">
        {% if p.wish %}<span class="badge {{ 'text-bg-success' if p.wish_positive else 'text-bg-light border' }}">{{ p.wish }}</span>{% else %}<span class="text-muted">－</span>{% endif %}
      </td>
      <td class="small text-nowrap">
        {% if p.stance_latest %}
        <span class="badge {{ 'text-bg-success' if p.stance_latest in C.MONTHLY_KPI['self_positive'] else ('text-bg-primary' if p.stance_latest in C.MONTHLY_KPI['ask_positive'] else 'text-bg-light border') }}">{{ p.stance_latest }}</span>
        <div class="text-muted">{{ p.latest_label }}{% if p.self_positive_count %}・自分で {{ p.self_positive_count }}回{% endif %}</div>
        {% else %}<span class="text-muted">－</span>{% endif %}
      </td>
      <td class="small dx-wide">{{ p.skills|join('・') or '－' }}</td>
      <td class="small dx-wide">{{ p.learn|join('・') or '－' }}</td>
      <td class="small text-nowrap">{{ p.hours or '－' }}</td>
      <td class="small text-nowrap">{{ p.share or '－' }}</td>
      <td class="small dx-tasks">
        {% if p.task_count %}<div>{{ p.task_count }} 件</div>{% endif %}
        {% for t in p.target_tasks %}<div><a href="{{ url_for('cases.detail', task_id=t.id) }}">{{ t.name }}</a> <span class="text-muted">自分で</span></div>{% endfor %}
        {% if p.targets %}<div class="text-muted">{{ p.targets }}</div>{% endif %}
        {% if not p.task_count and not p.targets %}<span class="text-muted">－</span>{% endif %}
      </td>
    </tr>
    {% else %}
    <tr><td colspan="9" class="text-muted">該当する人はいません</td></tr>
    {% endfor %}
    </tbody>
  </table>
  </div>
</div></div>

<div class="row g-3">
  <div class="col-lg-4">
    <div class="card h-100"><div class="card-body">
      <h5 class="card-title">意欲（改善への関わり方）</h5>
      {% set maxc = (c1_counts|map(attribute=1)|max) if c1_counts else 1 %}
      {% for opt, n in c1_counts %}
      <div class="d-flex align-items-center small mb-1">
        <div style="width: 11rem">{{ opt }}</div>
        <div class="progress flex-grow-1" style="height: 14px"><div class="progress-bar" style="width: {{ (n / (maxc or 1) * 100)|round }}%"></div></div>
        <div style="width: 3rem" class="text-end">{{ n }} 人</div>
      </div>
      {% endfor %}
    </div></div>
  </div>
  <div class="col-lg-4">
    <div class="card h-100"><div class="card-body">
      <h5 class="card-title">現在使える技術</h5>
      {% set maxh = (have[0][1] if have else 1) %}
      {% for skill, n in have %}
      <div class="d-flex align-items-center small mb-1">
        <div style="width: 9rem">{{ skill }}</div>
        <div class="progress flex-grow-1" style="height: 14px"><div class="progress-bar" style="width: {{ (n / maxh * 100)|round }}%"></div></div>
        <div style="width: 3rem" class="text-end">{{ n }} 人</div>
      </div>
      {% else %}
      <p class="text-muted small mb-0">データなし</p>
      {% endfor %}
    </div></div>
  </div>
  <div class="col-lg-4">
    <div class="card h-100"><div class="card-body">
      <h5 class="card-title">学びたい技術</h5>
      {% set maxl = (learn[0][1] if learn else 1) %}
      {% for skill, n in learn %}
      <div class="d-flex align-items-center small mb-1">
        <div style="width: 9rem">{{ skill }}</div>
        <div class="progress flex-grow-1" style="height: 14px"><div class="progress-bar bg-success" style="width: {{ (n / maxl * 100)|round }}%"></div></div>
        <div style="width: 3rem" class="text-end">{{ n }} 人</div>
      </div>
      {% else %}
      <p class="text-muted small mb-0">データなし</p>
      {% endfor %}
    </div></div>
  </div>
</div>
{% endblock %}
''',
    'survey/index.html': r'''{% extends "base.html" %}
{% block title %}自動化依頼{% endblock %}
{% block content %}
<div class="rq" id="rq">
  <div class="rq-top">
    <h1 class="rq-title">自動化依頼</h1>
  </div>

  <section class="card rq-card rq-sec" id="sec-person">
    <div class="card-body">
      <div class="rq-sec-head"><span class="rq-no">1</span><span class="rq-sec-title">あなたのこと</span></div>
      <div id="rq-person"></div>
    </div>
  </section>

  <section class="card rq-card rq-sec" id="sec-task">
    <div class="card-body">
      <div class="rq-sec-head"><span class="rq-no">2</span><span class="rq-sec-title">依頼する業務</span></div>
      <div id="rq-task"></div>
      <div class="rq-actions">
        <button type="button" class="btn btn-primary btn-lg" id="rq-save-task"><i class="bi bi-plus-lg"></i> この業務を登録</button>
      </div>
    </div>
  </section>

  <section class="card rq-card rq-sec" id="sec-c">
    <div class="card-body">
      <div class="rq-sec-head"><span class="rq-no">3</span><span class="rq-sec-title">改善への関わり方</span><span class="rq-opt">任意</span></div>
      <div id="rq-c"></div>
    </div>
  </section>

  <section class="card rq-card rq-sec" id="sec-d">
    <div class="card-body">
      <div class="rq-sec-head"><span class="rq-no">4</span><span class="rq-sec-title">部全体について</span><span class="rq-opt">任意</span></div>
      <div id="rq-d"></div>
    </div>
  </section>

  <div class="rq-finish">
    <button type="button" class="btn btn-primary btn-lg" id="rq-finish"><i class="bi bi-send"></i> 送信して完了</button>
  </div>

  <section class="card rq-card rq-done d-none" id="rq-done">
    <div class="card-body text-center">
      <div class="rq-check"><i class="bi bi-check-lg"></i></div>
      <h2 class="h4 mb-1">依頼を受け付けました</h2>
      <p class="text-muted mb-3" id="rq-done-sum"></p>
      <a class="btn btn-primary" href="{{ url_for('board.index') }}">活動状況を見る</a>
    </div>
  </section>
</div>
<div class="rq-toast" id="rq-toast"></div>
{% endblock %}
{% block scripts %}
<script id="rq-options" type="application/json">{{ options|tojson }}</script>
<script id="rq-state" type="application/json">{{ state|tojson }}</script>
<script>
  window.RQ_CONFIG = {
    emp: {{ emp|tojson }},
    api: {
      lookup: "{{ url_for('monthly.api_employee', emp_id='') }}",
      respondent: "{{ url_for('survey.api_respondent') }}",
      reset: "{{ url_for('survey.api_reset') }}",
      tasks: "{{ url_for('survey.api_task_create') }}",
      partC: "{{ url_for('survey.api_part_c') }}",
      partD: "{{ url_for('survey.api_part_d') }}"
    }
  };
</script>
<script src="{{ url_for('static', filename='survey.js') }}"></script>
{% endblock %}
''',
}


# ==============================================================================
# 静的ファイル
# ==============================================================================
# 静的ファイル（CSS / JS）。/static/<名前> で配信する
STATIC = {
    'app.js': r'''// 共通スクリプト
document.addEventListener("DOMContentLoaded", function () {
  // data-confirm 付きフォームは送信前に確認する
  document.querySelectorAll("form[data-confirm]").forEach(function (form) {
    form.addEventListener("submit", function (e) {
      if (!window.confirm(form.getAttribute("data-confirm"))) {
        e.preventDefault();
      }
    });
  });

  // 月間時間の目安をフォーム上で即時計算する（パートB）
  var freq = document.getElementById("frequency");
  var dur = document.getElementById("duration");
  var head = document.getElementById("headcount");
  var out = document.getElementById("monthly-estimate");
  if (freq && dur && out) {
    var coef = JSON.parse(out.getAttribute("data-coef") || "{}");
    var mins = JSON.parse(out.getAttribute("data-minutes") || "{}");
    var update = function () {
      var c = coef[freq.value] || 0;
      var m = mins[dur.value] || 0;
      var h = parseInt(head && head.value ? head.value : "1", 10) || 1;
      var total = c * m * h;
      if (!freq.value || !dur.value) {
        out.textContent = "頻度と所要時間を選ぶと月間時間の目安を表示します";
      } else if (freq.value === "不定期") {
        out.textContent = "不定期の業務は集計時に個別確認します";
      } else {
        out.textContent = "月間時間の目安: 約 " + (total / 60).toFixed(1) + " 時間（" + Math.round(total) + " 分）";
      }
    };
    [freq, dur, head].forEach(function (el) {
      if (el) el.addEventListener("change", update);
    });
    update();
  }
});

// 指標の数字をカウントアップで表示する（視差効果を減らす設定ではそのまま表示）
(function () {
  if (window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches) return;
  var targets = document.querySelectorAll(".stat-value, .mini-value, .kpi-value, .pb-fcount");
  var ease = function (t) { return 1 - Math.pow(1 - t, 3); };
  targets.forEach(function (el, i) {
    var node = el.firstChild;
    if (!node || node.nodeType !== 3) return;
    var text = node.nodeValue.trim();
    var m = text.match(/^-?[\d,]+(\.\d+)?$/);
    if (!m) return;
    var hasComma = text.indexOf(",") >= 0;
    var decimals = m[1] ? m[1].length - 1 : 0;
    var end = parseFloat(text.replace(/,/g, ""));
    if (isNaN(end)) return;
    var duration = 900, start = null, delay = Math.min(i, 8) * 60;
    var fmt = function (v) {
      var s = v.toFixed(decimals);
      return hasComma ? Number(s).toLocaleString("ja-JP", { minimumFractionDigits: decimals, maximumFractionDigits: decimals }) : s;
    };
    node.nodeValue = fmt(0);
    var step = function (ts) {
      if (start === null) start = ts + delay;
      var t = Math.min(Math.max((ts - start) / duration, 0), 1);
      node.nodeValue = fmt(end * ease(t));
      if (t < 1) requestAnimationFrame(step); else node.nodeValue = text;
    };
    requestAnimationFrame(step);
    // 画面が隠れていて requestAnimationFrame が止まっていても、最後は本来の値にする
    setTimeout(function () { if (node.nodeValue !== text) node.nodeValue = text; }, delay + duration + 300);
  });
})();

// 活動状況の見出しのプルダウン（「全体」「製造1課」…）を選んだ文字の幅に合わせる
(function () {
  var sel = document.querySelector(".pb-scope-select");
  if (!sel) return;
  if (window.CSS && CSS.supports && CSS.supports("field-sizing", "content")) return;  // ブラウザが文字幅に合わせてくれる
  var probe = document.createElement("span");
  probe.style.cssText = "position:absolute;left:-9999px;top:0;visibility:hidden;white-space:nowrap;";
  document.body.appendChild(probe);
  function fit() {
    var cs = getComputedStyle(sel);
    probe.style.fontFamily = cs.fontFamily;
    probe.style.fontSize = cs.fontSize;
    probe.style.fontWeight = cs.fontWeight;
    probe.style.letterSpacing = cs.letterSpacing;
    var opt = sel.options[sel.selectedIndex];
    probe.textContent = opt ? opt.text : "";
    sel.style.width = Math.ceil(probe.offsetWidth + (parseFloat(cs.paddingRight) || 0) + 6) + "px";
  }
  window.addEventListener("resize", fit);
  window.addEventListener("load", fit);
  if (document.fonts && document.fonts.ready) document.fonts.ready.then(fit);
  sel.addEventListener("change", fit);
  fit();
})();
''',
    'flowedit.js': r'''// 業務マップ 編集エディタ
// mindmap（Flex Diagram）の操作感を流れ図向けに応用: ドラッグ移動、ダブルクリック編集、🔗でリンク作成、
// 矢印クリックで向き反転、Shift範囲選択、Undo/Redo、名前付き保存。
// 自動生成したノード（source: auto）と手で足したノード（source: manual）を区別して保存し、
// 再生成時はサーバー側で手修正をマージする。
(function () {
  var app = document.getElementById("fe-app");
  if (!app) return;

  var cfg = {
    kind: app.dataset.kind,
    workplace: app.dataset.workplace,
    domain: app.dataset.domain,
    api: app.dataset.api
  };
  var els = {
    container: document.getElementById("fe-container"),
    canvas: document.getElementById("fe-canvas"),
    nodes: document.getElementById("fe-nodes"),
    edges: document.getElementById("fe-edges"),
    temp: document.getElementById("fe-temp"),
    selection: document.getElementById("fe-selection"),
    status: document.getElementById("fe-status"),
    mode: document.getElementById("fe-mode"),
    zoomLabel: document.getElementById("fe-zoom-label"),
    nodePop: document.getElementById("fe-node-popover"),
    edgePop: document.getElementById("fe-edge-popover"),
    edgeLabel: document.getElementById("fe-edge-label"),
    clsSelect: document.getElementById("fe-cls-select"),
    clsRow: document.getElementById("fe-cls-row"),
    undo: document.getElementById("fe-undo"),
    redo: document.getElementById("fe-redo")
  };

  var G = { nodes: [], edges: [], suppressed_nodes: [], suppressed_edges: [] };
  var meta = {};
  var view = { x: 40, y: 40, zoom: 1 };
  var history = [], historyIndex = -1, MAX_HISTORY = 60;
  var savedSnapshot = null, dirty = false;
  var selected = [];
  var drag = null, pan = null, selecting = null, linkSource = null, editingId = null;
  var popNodeId = null, popEdgeId = null;
  var lastClick = { x: 120, y: 120 };

  // ---------- 基本 ----------
  function nodeById(id) { for (var i = 0; i < G.nodes.length; i++) if (G.nodes[i].id === id) return G.nodes[i]; return null; }
  function edgeById(id) { for (var i = 0; i < G.edges.length; i++) if (G.edges[i].id === id) return G.edges[i]; return null; }
  function nodeEl(id) { return els.nodes.querySelector('[data-id="' + CSS.escape(id) + '"]'); }
  function toCanvas(clientX, clientY) {
    var r = els.container.getBoundingClientRect();
    return { x: (clientX - r.left - view.x) / view.zoom, y: (clientY - r.top - view.y) / view.zoom };
  }
  function applyView() {
    els.canvas.style.transform = "translate(" + view.x + "px, " + view.y + "px) scale(" + view.zoom + ")";
    els.zoomLabel.textContent = Math.round(view.zoom * 100) + "%";
  }
  function setMode(text) { els.mode.textContent = text || ""; els.mode.style.display = text ? "block" : "none"; }
  function setStatus(text) { els.status.textContent = text || ""; }
  function newId(prefix) { return prefix + Date.now().toString(36) + "_" + Math.random().toString(36).slice(2, 7); }

  function snapshot() {
    return JSON.stringify({ nodes: G.nodes, edges: G.edges, suppressed_nodes: G.suppressed_nodes, suppressed_edges: G.suppressed_edges });
  }
  function pushHistory() {
    history = history.slice(0, historyIndex + 1);
    history.push(snapshot());
    if (history.length > MAX_HISTORY) history.shift(); else historyIndex++;
    updateButtons();
    updateDirty();
  }
  function restore(snap) {
    var s = JSON.parse(snap);
    G.nodes = s.nodes; G.edges = s.edges; G.suppressed_nodes = s.suppressed_nodes; G.suppressed_edges = s.suppressed_edges;
    selected = [];
    closePopovers();
    renderAll();
    updateDirty();
  }
  function undo() { if (historyIndex <= 0) return; historyIndex--; restore(history[historyIndex]); updateButtons(); }
  function redo() { if (historyIndex >= history.length - 1) return; historyIndex++; restore(history[historyIndex]); updateButtons(); }
  function updateButtons() { els.undo.disabled = historyIndex <= 0; els.redo.disabled = historyIndex >= history.length - 1; }
  function updateDirty() {
    dirty = savedSnapshot !== null && snapshot() !== savedSnapshot;
    document.title = (dirty ? "* " : "") + document.title.replace(/^\* /, "");
    if (dirty) setStatus("未保存の変更があります");
    else if (meta.saved_at) setStatus("保存済み（" + meta.saved_at + "）");
    else setStatus("自動生成（未保存）");
  }

  // ---------- 描画 ----------
  function renderAll() {
    els.nodes.innerHTML = "";
    G.nodes.forEach(function (n) { els.nodes.appendChild(createNodeEl(n)); });
    drawEdges();
  }

  function createNodeEl(n) {
    var div = document.createElement("div");
    div.className = "fe-node fe-cls-" + (n.cls || "manual") + (n.source === "manual" ? " fe-manual" : "");
    div.dataset.id = n.id;
    if (n.color) div.dataset.color = n.color;
    div.style.left = (n.x || 0) + "px";
    div.style.top = (n.y || 0) + "px";
    if (selected.indexOf(n.id) !== -1) div.classList.add("fe-selected");
    if (linkSource === n.id) div.classList.add("fe-link-source");

    var btns = document.createElement("div");
    btns.className = "fe-btns";
    if (n.url) btns.appendChild(btn("↗", "案件詳細を開く", "fe-btn-open", function () { window.open(n.url, "_blank"); }));
    btns.appendChild(btn("🔗", "この箱から矢印を引く", "fe-btn-link", function () { startLink(n.id); }));
    btns.appendChild(btn("➕", "つながる箱を追加", "fe-btn-add", function () { addChild(n.id); }));
    btns.appendChild(btn("🎨", "色・種類", "fe-btn-color", function (e) { openNodePopover(n.id, e); }));
    btns.appendChild(btn("✕", n.source === "manual" ? "削除" : "非表示にする", "fe-btn-del", function () { deleteNode(n.id); }));
    div.appendChild(btns);

    var text = document.createElement("div");
    text.className = "fe-text";
    text.textContent = n.text || "";
    div.appendChild(text);
    if (n.group_title) {
      var sub = document.createElement("div");
      sub.className = "fe-sub";
      sub.textContent = n.group_title;
      div.appendChild(sub);
    }
    if (n.text_edited) {
      var mark = document.createElement("span");
      mark.className = "fe-edited";
      mark.title = "名前を手で修正";
      mark.textContent = "✎";
      div.appendChild(mark);
    }

    text.addEventListener("dblclick", function (e) { e.stopPropagation(); startEdit(n.id); });
    div.addEventListener("mousedown", function (e) {
      if (editingId === n.id) return;
      if (e.button !== 0) return;
      if (linkSource) return;
      e.preventDefault();
      e.stopPropagation();
      if (e.shiftKey) { toggleSelect(n.id); return; }
      if (selected.indexOf(n.id) === -1) selectOnly(n.id);
      var start = toCanvas(e.clientX, e.clientY);
      var origin = {};
      selected.forEach(function (id) { var m = nodeById(id); if (m) origin[id] = { x: m.x, y: m.y }; });
      drag = { startX: start.x, startY: start.y, origin: origin, moved: false };
    });
    div.addEventListener("click", function (e) {
      e.stopPropagation();
      if (linkSource && linkSource !== n.id) { createLink(linkSource, n.id); }
      else if (linkSource === n.id) { cancelLink(); }
    });
    div.addEventListener("mouseenter", function () { if (linkSource && linkSource !== n.id) drawTempLine(n.id); });
    return div;
  }

  function btn(label, title, cls, onClick) {
    var b = document.createElement("button");
    b.type = "button";
    b.className = "fe-btn " + cls;
    b.title = title;
    b.textContent = label;
    b.addEventListener("mousedown", function (e) { e.stopPropagation(); e.preventDefault(); });
    b.addEventListener("click", function (e) { e.stopPropagation(); onClick(e); });
    return b;
  }

  function bounds(id) {
    var n = nodeById(id), el = nodeEl(id);
    if (!n || !el) return null;
    var w = el.offsetWidth, h = el.offsetHeight;
    return { cx: n.x + w / 2, cy: n.y + h / 2, rx: w / 2, ry: h / 2, w: w, h: h };
  }

  function endpoints(aId, bId) {
    var A = bounds(aId), B = bounds(bId);
    if (!A || !B) return null;
    var ang = Math.atan2(B.cy - A.cy, B.cx - A.cx);
    return {
      x1: A.cx + A.rx * Math.cos(ang), y1: A.cy + A.ry * Math.sin(ang),
      x2: B.cx - B.rx * Math.cos(ang), y2: B.cy - B.ry * Math.sin(ang)
    };
  }

  function curve(p) {
    var k = Math.max(30, Math.abs(p.x2 - p.x1) / 2);
    var c1x = p.x1 + k, c2x = p.x2 - k;
    return { d: "M " + p.x1 + " " + p.y1 + " C " + c1x + " " + p.y1 + ", " + c2x + " " + p.y2 + ", " + p.x2 + " " + p.y2,
      mx: 0.125 * p.x1 + 0.375 * c1x + 0.375 * c2x + 0.125 * p.x2,
      my: 0.125 * p.y1 + 0.375 * p.y1 + 0.375 * p.y2 + 0.125 * p.y2 };
  }

  function svgEl(tag, attrs) {
    var el = document.createElementNS("http://www.w3.org/2000/svg", tag);
    Object.keys(attrs || {}).forEach(function (k) { el.setAttribute(k, attrs[k]); });
    return el;
  }

  function drawEdges() {
    els.edges.innerHTML = "";
    G.edges.forEach(function (e) {
      var a = e.reversed ? e.to : e.from, b = e.reversed ? e.from : e.to;
      var p = endpoints(a, b);
      if (!p) return;
      var c = curve(p);
      var kind = e.source === "manual" ? "manual" : (e.kind || "info");
      var g = svgEl("g", { "class": "fe-edge-group", "data-id": e.id });
      var hit = svgEl("path", { d: c.d, "class": "fe-hit" });
      var path = svgEl("path", { d: c.d, "class": "fe-edge fe-kind-" + kind + (e.source === "manual" ? " fe-src-manual" : "") });
      if (kind !== "relation") path.setAttribute("marker-end", "url(#fe-arrow-" + kind + ")");
      g.appendChild(hit);
      g.appendChild(path);
      if (e.label) {
        var t = svgEl("text", { x: c.mx, y: c.my - 4, "class": "fe-edge-label", "text-anchor": "middle" });
        t.textContent = e.label;
        g.appendChild(t);
      }
      g.addEventListener("click", function (ev) { ev.stopPropagation(); openEdgePopover(e.id, ev); });
      els.edges.appendChild(g);
    });
  }

  function updateNodePositions(ids) {
    (ids || G.nodes.map(function (n) { return n.id; })).forEach(function (id) {
      var n = nodeById(id), el = nodeEl(id);
      if (n && el) { el.style.left = n.x + "px"; el.style.top = n.y + "px"; }
    });
  }

  // ---------- 選択 ----------
  function selectOnly(id) {
    selected = id ? [id] : [];
    els.nodes.querySelectorAll(".fe-node").forEach(function (el) { el.classList.toggle("fe-selected", selected.indexOf(el.dataset.id) !== -1); });
  }
  function toggleSelect(id) {
    var i = selected.indexOf(id);
    if (i === -1) selected.push(id); else selected.splice(i, 1);
    var el = nodeEl(id);
    if (el) el.classList.toggle("fe-selected", i === -1);
  }
  function selectInRect(r) {
    selected = [];
    G.nodes.forEach(function (n) {
      var b = bounds(n.id);
      if (!b) return;
      var left = n.x * view.zoom + view.x, top = n.y * view.zoom + view.y;
      var right = left + b.w * view.zoom, bottom = top + b.h * view.zoom;
      if (left < r.right && right > r.left && top < r.bottom && bottom > r.top) selected.push(n.id);
    });
    els.nodes.querySelectorAll(".fe-node").forEach(function (el) { el.classList.toggle("fe-selected", selected.indexOf(el.dataset.id) !== -1); });
  }

  // ---------- 編集操作 ----------
  function startEdit(id) {
    var n = nodeById(id), el = nodeEl(id);
    if (!n || !el) return;
    var text = el.querySelector(".fe-text");
    editingId = id;
    text.contentEditable = "true";
    text.classList.add("fe-editing");
    text.focus();
    var range = document.createRange();
    range.selectNodeContents(text);
    var sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
    var done = false;
    function finish(cancel) {
      if (done) return;
      done = true;
      text.contentEditable = "false";
      text.classList.remove("fe-editing");
      editingId = null;
      var value = text.textContent.replace(/\s+$/, "").trim();
      if (!cancel && value && value !== n.text) {
        n.text = value;
        n.text_edited = true;
        renderAll();
        pushHistory();
      } else {
        text.textContent = n.text;
        drawEdges();
      }
    }
    text.addEventListener("blur", function () { finish(false); }, { once: true });
    text.addEventListener("keydown", function (e) {
      if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); text.blur(); }
      if (e.key === "Escape") { e.preventDefault(); finish(true); }
    });
  }

  function addNode(x, y, text, cls) {
    var n = { id: newId("m_"), text: text || "新しいノード", cls: cls || "manual", shape: null, group: null, group_title: null, url: null,
      source: "manual", x: x, y: y, color: null, text_edited: false };
    G.nodes.push(n);
    els.nodes.appendChild(createNodeEl(n));
    selectOnly(n.id);
    drawEdges();
    pushHistory();
    setTimeout(function () { startEdit(n.id); }, 50);
    return n;
  }

  function addChild(parentId) {
    var p = nodeById(parentId), b = bounds(parentId);
    if (!p || !b) return;
    var n = addNode(p.x + b.w + 90, p.y, "新しいノード", "manual");
    G.edges.push({ id: parentId + "->" + n.id, from: parentId, to: n.id, label: "", kind: "manual", style: "solid", source: "manual", reversed: false, label_edited: false });
    drawEdges();
    pushHistory();
  }

  function deleteNode(id) {
    var n = nodeById(id);
    if (!n) return;
    if (n.source !== "manual" && G.suppressed_nodes.indexOf(id) === -1) G.suppressed_nodes.push(id);
    G.nodes = G.nodes.filter(function (m) { return m.id !== id; });
    G.edges = G.edges.filter(function (e) { return e.from !== id && e.to !== id; });
    selected = selected.filter(function (s) { return s !== id; });
    closePopovers();
    renderAll();
    pushHistory();
  }

  function deleteSelected() {
    if (!selected.length) return;
    selected.slice().forEach(function (id) {
      var n = nodeById(id);
      if (!n) return;
      if (n.source !== "manual" && G.suppressed_nodes.indexOf(id) === -1) G.suppressed_nodes.push(id);
      G.nodes = G.nodes.filter(function (m) { return m.id !== id; });
      G.edges = G.edges.filter(function (e) { return e.from !== id && e.to !== id; });
    });
    selected = [];
    closePopovers();
    renderAll();
    pushHistory();
  }

  function deleteEdge(id) {
    var e = edgeById(id);
    if (!e) return;
    if (e.source !== "manual" && G.suppressed_edges.indexOf(id) === -1) G.suppressed_edges.push(id);
    G.edges = G.edges.filter(function (x) { return x.id !== id; });
    closePopovers();
    drawEdges();
    pushHistory();
  }

  function reverseEdge(id) {
    var e = edgeById(id);
    if (!e) return;
    e.reversed = !e.reversed;
    drawEdges();
    pushHistory();
  }

  function setEdgeLabel(id, label) {
    var e = edgeById(id);
    if (!e || e.label === label) return;
    e.label = label;
    e.label_edited = true;
    drawEdges();
    pushHistory();
  }

  function startLink(id) {
    if (linkSource === id) { cancelLink(); return; }
    linkSource = id;
    els.canvas.classList.add("fe-linking");
    els.nodes.querySelectorAll(".fe-node").forEach(function (el) { el.classList.toggle("fe-link-source", el.dataset.id === id); });
    var n = nodeById(id);
    setMode("🔗 「" + (n ? n.text : "") + "」から矢印を引きます。相手の箱をクリック（Esc で中止）");
  }

  function cancelLink() {
    linkSource = null;
    els.canvas.classList.remove("fe-linking");
    els.nodes.querySelectorAll(".fe-link-source").forEach(function (el) { el.classList.remove("fe-link-source"); });
    els.temp.innerHTML = "";
    setMode("");
  }

  function drawTempLine(targetId) {
    if (!linkSource) return;
    var p = targetId ? endpoints(linkSource, targetId) : null;
    els.temp.innerHTML = "";
    if (!p) return;
    els.temp.appendChild(svgEl("path", { d: curve(p).d, "class": "fe-temp-line" }));
  }

  function drawTempLineTo(canvasPt) {
    var A = bounds(linkSource);
    if (!A) return;
    els.temp.innerHTML = "";
    els.temp.appendChild(svgEl("path", { d: curve({ x1: A.cx, y1: A.cy, x2: canvasPt.x, y2: canvasPt.y }).d, "class": "fe-temp-line" }));
  }

  function createLink(from, to) {
    var id = from + "->" + to;
    if (edgeById(id)) { setMode("この矢印はすでにあります"); setTimeout(cancelLink, 1200); return; }
    var sup = G.suppressed_edges.indexOf(id);
    if (sup !== -1) G.suppressed_edges.splice(sup, 1);
    G.edges.push({ id: id, from: from, to: to, label: "", kind: "manual", style: "solid", source: "manual", reversed: false, label_edited: false });
    cancelLink();
    drawEdges();
    pushHistory();
    setMode("✅ 矢印を作成しました");
    setTimeout(function () { if (!linkSource) setMode(""); }, 1500);
  }

  // ---------- ポップオーバー ----------
  function placePopover(pop, ev) {
    var r = els.container.getBoundingClientRect();
    var x = Math.min(ev.clientX - r.left + 8, r.width - 240), y = Math.min(ev.clientY - r.top + 8, r.height - 200);
    pop.style.left = Math.max(4, x) + "px";
    pop.style.top = Math.max(4, y) + "px";
    pop.classList.add("fe-show");
  }
  function closePopovers() {
    els.nodePop.classList.remove("fe-show");
    els.edgePop.classList.remove("fe-show");
    var wb = document.getElementById("fe-wb-popover");
    if (wb) wb.classList.remove("fe-show");
    popNodeId = null; popEdgeId = null;
  }
  function openNodePopover(id, ev) {
    closePopovers();
    var n = nodeById(id);
    if (!n) return;
    popNodeId = id;
    els.clsRow.style.display = n.source === "manual" ? "block" : "none";
    els.clsSelect.value = n.cls || "manual";
    placePopover(els.nodePop, ev);
  }
  function openEdgePopover(id, ev) {
    closePopovers();
    var e = edgeById(id);
    if (!e) return;
    popEdgeId = id;
    els.edgeLabel.value = e.label || "";
    placePopover(els.edgePop, ev);
    els.edgeLabel.focus();
  }

  els.nodePop.querySelectorAll(".fe-color").forEach(function (sw) {
    sw.addEventListener("click", function () {
      var n = nodeById(popNodeId);
      if (!n) return;
      n.color = sw.dataset.color || null;
      var el = nodeEl(n.id);
      if (el) { if (n.color) el.dataset.color = n.color; else delete el.dataset.color; }
      pushHistory();
    });
  });
  els.clsSelect.addEventListener("change", function () {
    var n = nodeById(popNodeId);
    if (!n || n.source !== "manual") return;
    n.cls = els.clsSelect.value;
    n.shape = null;
    renderAll();
    pushHistory();
  });
  document.getElementById("fe-node-delete").addEventListener("click", function () { if (popNodeId) deleteNode(popNodeId); });
  document.getElementById("fe-node-close").addEventListener("click", closePopovers);
  document.getElementById("fe-edge-reverse").addEventListener("click", function () { if (popEdgeId) { reverseEdge(popEdgeId); } });
  document.getElementById("fe-edge-delete").addEventListener("click", function () { if (popEdgeId) deleteEdge(popEdgeId); });
  document.getElementById("fe-edge-close").addEventListener("click", closePopovers);
  els.edgeLabel.addEventListener("change", function () { if (popEdgeId) setEdgeLabel(popEdgeId, els.edgeLabel.value.trim()); });
  els.edgeLabel.addEventListener("keydown", function (e) { if (e.key === "Enter") { e.preventDefault(); setEdgeLabel(popEdgeId, els.edgeLabel.value.trim()); closePopovers(); } });
  [els.nodePop, els.edgePop].forEach(function (p) { p.addEventListener("mousedown", function (e) { e.stopPropagation(); }); });

  // ---------- レイアウト ----------
  function nodeSizes() {
    var sizes = {};
    G.nodes.forEach(function (n) { var b = bounds(n.id); sizes[n.id] = b ? { w: b.w, h: b.h } : { w: 120, h: 40 }; });
    return sizes;
  }

  function autoLayout() {
    if (!window.dagre) { alert("自動レイアウト用のライブラリ（dagre）を読み込めませんでした。"); return; }
    var g = new dagre.graphlib.Graph();
    g.setGraph({ rankdir: "LR", nodesep: 28, ranksep: 90, marginx: 20, marginy: 20 });
    g.setDefaultEdgeLabel(function () { return {}; });
    var sizes = nodeSizes();
    G.nodes.forEach(function (n) { g.setNode(n.id, { width: sizes[n.id].w, height: sizes[n.id].h }); });
    G.edges.forEach(function (e) {
      var a = e.reversed ? e.to : e.from, b = e.reversed ? e.from : e.to;
      if (nodeById(a) && nodeById(b)) g.setEdge(a, b);
    });
    dagre.layout(g);
    G.nodes.forEach(function (n) {
      var p = g.node(n.id);
      if (p) { n.x = Math.round(p.x - sizes[n.id].w / 2); n.y = Math.round(p.y - sizes[n.id].h / 2); }
    });
    updateNodePositions();
    drawEdges();
  }

  function placeNewNodes(ids) {
    var placed = G.nodes.filter(function (n) { return n.x !== null && n.x !== undefined && ids.indexOf(n.id) === -1; });
    var maxX = 0, minY = 0;
    placed.forEach(function (n) { var b = bounds(n.id); maxX = Math.max(maxX, n.x + (b ? b.w : 120)); minY = Math.min(minY, n.y); });
    var y = minY;
    ids.forEach(function (id) {
      var n = nodeById(id);
      if (!n) return;
      n.x = maxX + 140;
      n.y = y;
      var b = bounds(id);
      y += (b ? b.h : 40) + 16;
    });
    updateNodePositions(ids);
    drawEdges();
  }

  function fitView() {
    if (!G.nodes.length) return;
    var minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
    G.nodes.forEach(function (n) {
      var b = bounds(n.id);
      if (!b) return;
      minX = Math.min(minX, n.x); minY = Math.min(minY, n.y);
      maxX = Math.max(maxX, n.x + b.w); maxY = Math.max(maxY, n.y + b.h);
    });
    var cw = els.container.clientWidth, ch = els.container.clientHeight;
    var bw = Math.max(1, maxX - minX), bh = Math.max(1, maxY - minY);
    var z = Math.min(1.2, Math.max(0.2, Math.min((cw - 60) / bw, (ch - 60) / bh)));
    view.zoom = z;
    view.x = (cw - bw * z) / 2 - minX * z;
    view.y = (ch - bh * z) / 2 - minY * z;
    applyView();
  }

  function zoomAt(factor, clientX, clientY) {
    var r = els.container.getBoundingClientRect();
    var cx = (clientX === undefined) ? r.width / 2 : clientX - r.left;
    var cy = (clientY === undefined) ? r.height / 2 : clientY - r.top;
    var nz = Math.min(3, Math.max(0.2, view.zoom * factor));
    view.x = cx - (cx - view.x) * (nz / view.zoom);
    view.y = cy - (cy - view.y) * (nz / view.zoom);
    view.zoom = nz;
    applyView();
  }

  // ---------- マウス・キーボード ----------
  els.container.addEventListener("mousedown", function (e) {
    if (e.button !== 0) return;
    var empty = e.target === els.container || e.target === els.canvas || e.target === els.nodes || e.target.id === "fe-svg";
    if (!empty) return;
    closePopovers();
    lastClick = toCanvas(e.clientX, e.clientY);
    if (linkSource) { cancelLink(); return; }
    var r = els.container.getBoundingClientRect();
    if (e.shiftKey) {
      selecting = { x0: e.clientX - r.left, y0: e.clientY - r.top };
      els.selection.style.display = "block";
      els.selection.style.left = selecting.x0 + "px"; els.selection.style.top = selecting.y0 + "px";
      els.selection.style.width = "0px"; els.selection.style.height = "0px";
      e.preventDefault();
    } else {
      selectOnly(null);
      pan = { startX: e.clientX, startY: e.clientY, vx: view.x, vy: view.y };
      els.container.classList.add("fe-panning");
    }
  });

  document.addEventListener("mousemove", function (e) {
    if (drag) {
      var p = toCanvas(e.clientX, e.clientY);
      var dx = p.x - drag.startX, dy = p.y - drag.startY;
      if (Math.abs(dx) > 1 || Math.abs(dy) > 1) drag.moved = true;
      Object.keys(drag.origin).forEach(function (id) {
        var n = nodeById(id);
        if (n) { n.x = Math.round(drag.origin[id].x + dx); n.y = Math.round(drag.origin[id].y + dy); }
      });
      updateNodePositions(Object.keys(drag.origin));
      drawEdges();
    } else if (pan) {
      view.x = pan.vx + (e.clientX - pan.startX);
      view.y = pan.vy + (e.clientY - pan.startY);
      applyView();
    } else if (selecting) {
      var r = els.container.getBoundingClientRect();
      var x = e.clientX - r.left, y = e.clientY - r.top;
      var rect = { left: Math.min(selecting.x0, x), top: Math.min(selecting.y0, y), right: Math.max(selecting.x0, x), bottom: Math.max(selecting.y0, y) };
      els.selection.style.left = rect.left + "px"; els.selection.style.top = rect.top + "px";
      els.selection.style.width = (rect.right - rect.left) + "px"; els.selection.style.height = (rect.bottom - rect.top) + "px";
      selectInRect(rect);
    } else if (linkSource && !e.target.closest(".fe-node")) {
      drawTempLineTo(toCanvas(e.clientX, e.clientY));
    }
  });

  document.addEventListener("mouseup", function () {
    if (drag) { if (drag.moved) pushHistory(); drag = null; }
    if (pan) { pan = null; els.container.classList.remove("fe-panning"); }
    if (selecting) { selecting = null; els.selection.style.display = "none"; }
  });

  els.container.addEventListener("wheel", function (e) {
    e.preventDefault();
    zoomAt(e.deltaY < 0 ? 1.1 : 1 / 1.1, e.clientX, e.clientY);
  }, { passive: false });

  document.addEventListener("keydown", function (e) {
    var typing = e.target.tagName === "INPUT" || e.target.tagName === "TEXTAREA" || e.target.isContentEditable;
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "s") { e.preventDefault(); save(); return; }
    if (typing) return;
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "z" && !e.shiftKey) { e.preventDefault(); undo(); }
    else if ((e.ctrlKey || e.metaKey) && (e.key.toLowerCase() === "y" || (e.key.toLowerCase() === "z" && e.shiftKey))) { e.preventDefault(); redo(); }
    else if (e.key === "Delete" || e.key === "Backspace") { if (selected.length) { e.preventDefault(); deleteSelected(); } }
    else if (e.key === "Escape") { cancelLink(); closePopovers(); selectOnly(null); }
  });

  window.addEventListener("beforeunload", function (e) {
    if (dirty) { e.preventDefault(); e.returnValue = ""; }
  });

  // ---------- サーバー ----------
  function apiUrl() {
    var q = new URLSearchParams({ kind: cfg.kind });
    if (cfg.workplace) q.set("workplace_id", cfg.workplace);
    if (cfg.domain) q.set("domain_id", cfg.domain);
    return cfg.api + "?" + q.toString();
  }

  async function load(opts) {
    opts = opts || {};
    setStatus("読み込み中…");
    var res = await fetch(apiUrl(), { headers: { Accept: "application/json" } });
    var data = await res.json();
    G = { nodes: data.nodes || [], edges: data.edges || [], suppressed_nodes: data.suppressed_nodes || [], suppressed_edges: data.suppressed_edges || [] };
    meta = data.meta || {};
    selected = [];
    cancelLink();
    closePopovers();
    G.nodes.forEach(function (n) { if (n.x === null || n.x === undefined) { n.x = 0; n.y = 0; } });
    renderAll();
    var newIds = data.new_node_ids || [];
    if (data.needs_layout) {
      autoLayout();
      fitView();
    } else {
      if (newIds.length) placeNewNodes(newIds);
      if (data.view && !opts.keepView) { view = { x: data.view.x || 0, y: data.view.y || 0, zoom: data.view.zoom || 1 }; applyView(); }
      else if (!opts.keepView) fitView();
    }
    history = []; historyIndex = -1;
    savedSnapshot = snapshot();
    pushHistory();
    if (!meta.saved_at) setStatus("自動生成（未保存）。配置や名前を直して保存すると、再生成しても修正が残ります");
    if (newIds.length) setMode("新しく登録された業務など " + newIds.length + " 件を右端に追加しました。配置を整えて保存してください");
    else setMode("");
    if (meta.truncated) setMode("業務が " + meta.task_count + " 件を超えるため一部のみ表示しています。職場や領域で絞り込んでください");
  }

  async function save() {
    setStatus("保存中…");
    var payload = {
      kind: cfg.kind, workplace_id: cfg.workplace || null, domain_id: cfg.domain || null,
      data: { nodes: G.nodes, edges: G.edges, suppressed_nodes: G.suppressed_nodes, suppressed_edges: G.suppressed_edges, view: view }
    };
    try {
      var res = await fetch(cfg.api, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) });
      var data = await res.json();
      if (!res.ok || !data.ok) throw new Error(data.error || res.statusText);
      meta.saved_at = data.saved_at;
      savedSnapshot = snapshot();
      updateDirty();
      setMode("");
    } catch (err) {
      setStatus("保存に失敗しました: " + err.message);
    }
  }

  async function regenerate() {
    if (dirty && !confirm("未保存の変更があります。再生成すると失われます（保存済みの手修正は残ります）。続けますか？")) return;
    await load({ keepView: true });
  }

  async function restoreHidden() {
    if (!G.suppressed_nodes.length && !G.suppressed_edges.length) { setMode("非表示にしたものはありません"); return; }
    G.suppressed_nodes = []; G.suppressed_edges = [];
    await save();
    await load({ keepView: true });
  }

  async function resetAuto() {
    if (!confirm("手修正（配置・名前・追加・削除・向き）をすべて捨てて、自動生成の図に戻します。よろしいですか？")) return;
    await fetch(apiUrl(), { method: "DELETE" });
    await load();
  }

  // ---------- 業務データへの書き戻し ----------
  var wbPop = document.getElementById("fe-wb-popover");
  var wbList = document.getElementById("fe-wb-list");
  function writebackUrl() { return cfg.api.replace(/\/graph$/, "/writeback") + apiUrl().slice(cfg.api.length); }

  async function openWriteback() {
    if (dirty) {
      if (!confirm("未保存の変更があります。先に保存してから、業務データに反映する内容を取り出します。よろしいですか？")) return;
      await save();
    }
    closePopovers();
    wbList.innerHTML = '<div class="text-muted small">取り出し中…</div>';
    wbPop.classList.add("fe-show");
    wbPop.style.left = "16px"; wbPop.style.top = "16px";
    try {
      var res = await fetch(writebackUrl(), { headers: { Accept: "application/json" } });
      var data = await res.json();
      var items = data.candidates || [];
      if (!items.length) {
        wbList.innerHTML = '<div class="text-muted small">反映できる変更はありません。業務同士の矢印、情報源や部署から業務への矢印、業務名の修正、消した前後の矢印が対象です。</div>';
        document.getElementById("fe-wb-apply").disabled = true;
        return;
      }
      document.getElementById("fe-wb-apply").disabled = false;
      wbList.innerHTML = "";
      items.forEach(function (c) {
        var row = document.createElement("label");
        row.className = "fe-wb-item";
        var cb = document.createElement("input");
        cb.type = "checkbox"; cb.checked = true; cb.value = c.id; cb.className = "form-check-input me-1";
        row.appendChild(cb);
        row.appendChild(document.createTextNode(c.text));
        wbList.appendChild(row);
      });
    } catch (err) {
      wbList.innerHTML = '<div class="text-danger small">取り出しに失敗しました: ' + err.message + "</div>";
    }
  }

  async function applyWriteback() {
    var ids = Array.prototype.map.call(wbList.querySelectorAll("input:checked"), function (cb) { return cb.value; });
    if (!ids.length) { setMode("反映する変更を選んでください"); return; }
    try {
      var res = await fetch(cfg.api.replace(/\/graph$/, "/writeback"), {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ kind: cfg.kind, workplace_id: cfg.workplace || null, domain_id: cfg.domain || null, ids: ids })
      });
      var data = await res.json();
      if (!res.ok || !data.ok) throw new Error(data.error || res.statusText);
      wbPop.classList.remove("fe-show");
      setMode("✅ 業務データに " + data.applied + " 件反映しました。図を再生成します");
      await load({ keepView: true });
      setMode("✅ 業務データに " + data.applied + " 件反映しました");
    } catch (err) {
      setMode("反映に失敗しました: " + err.message);
    }
  }
  document.getElementById("fe-writeback").addEventListener("click", openWriteback);
  document.getElementById("fe-wb-apply").addEventListener("click", applyWriteback);
  document.getElementById("fe-wb-close").addEventListener("click", function () { wbPop.classList.remove("fe-show"); });
  wbPop.addEventListener("mousedown", function (e) { e.stopPropagation(); });

  // ---------- ツールバー ----------
  document.getElementById("fe-save").addEventListener("click", save);
  els.undo.addEventListener("click", undo);
  els.redo.addEventListener("click", redo);
  document.getElementById("fe-add").addEventListener("click", function () {
    var r = els.container.getBoundingClientRect();
    var c = toCanvas(r.left + r.width / 2, r.top + r.height / 2);
    addNode(Math.round(c.x - 60), Math.round(c.y - 20), "新しいノード", "manual");
  });
  document.getElementById("fe-layout").addEventListener("click", function () { autoLayout(); fitView(); pushHistory(); });
  document.getElementById("fe-fit").addEventListener("click", fitView);
  document.getElementById("fe-zoom-in").addEventListener("click", function () { zoomAt(1.25); });
  document.getElementById("fe-zoom-out").addEventListener("click", function () { zoomAt(1 / 1.25); });
  els.zoomLabel.addEventListener("click", function () { view.zoom = 1; applyView(); });
  document.getElementById("fe-reload").addEventListener("click", regenerate);
  document.getElementById("fe-restore").addEventListener("click", restoreHidden);
  document.getElementById("fe-reset").addEventListener("click", resetAuto);

  applyView();
  load();
})();
''',
    'flows.js': r'''// 業務マップ（Mermaid）の描画・拡大縮小
(function () {
  if (!window.mermaid) return;
  mermaid.initialize({
    startOnLoad: false,
    securityLevel: "loose",
    theme: "neutral",
    fontFamily: "inherit",
    flowchart: { htmlLabels: true, curve: "basis", nodeSpacing: 28, rankSpacing: 48, padding: 8 }
  });

  var scales = {};

  function keyOf(el) {
    return el.id.replace(/^flow-/, "");
  }

  function applyScale(el) {
    var svg = el.querySelector("svg");
    if (!svg) return;
    var vb = svg.viewBox && svg.viewBox.baseVal;
    if (!vb || !vb.width) return;
    var s = scales[keyOf(el)] || 1;
    svg.style.maxWidth = "none";
    svg.style.width = Math.round(vb.width * s) + "px";
    svg.style.height = Math.round(vb.height * s) + "px";
    var tb = document.querySelector('.flow-toolbar[data-target="' + keyOf(el) + '"]');
    var pct = tb && tb.querySelector("[data-zoom='reset']");
    if (pct) pct.textContent = Math.round(s * 100) + "%";
  }

  async function render(el) {
    if (!el || el.dataset.rendered) return;
    var defEl = document.getElementById(el.dataset.def);
    if (!defEl) return;
    var def = defEl.textContent;
    el.dataset.rendered = "1";
    try {
      var out = await mermaid.render("svg-" + el.id, def);
      el.innerHTML = out.svg;
      if (out.bindFunctions) out.bindFunctions(el);
      var svg = el.querySelector("svg");
      var vb = svg && svg.viewBox && svg.viewBox.baseVal;
      var canvas = el.parentElement;
      if (vb && vb.width && canvas) {
        // 初期表示は横幅に収まるように縮小（最小 35%）
        var fit = (canvas.clientWidth - 32) / vb.width;
        scales[keyOf(el)] = Math.max(0.25, Math.min(1, fit));
      }
      applyScale(el);
    } catch (e) {
      el.innerHTML = '<div class="text-danger small">図の描画に失敗しました: ' + (e && e.message ? e.message : e) + "</div>" +
        '<pre class="small text-muted">' + def.replace(/</g, "&lt;") + "</pre>";
    }
  }

  function visible(el) {
    return el && el.offsetParent !== null;
  }

  document.querySelectorAll(".flow-inner").forEach(function (el) {
    if (visible(el)) render(el);
  });

  document.querySelectorAll('[data-bs-toggle="tab"]').forEach(function (btn) {
    btn.addEventListener("shown.bs.tab", function () {
      var pane = document.querySelector(btn.getAttribute("data-bs-target"));
      if (pane) pane.querySelectorAll(".flow-inner").forEach(render);
    });
  });

  document.querySelectorAll(".flow-toolbar").forEach(function (tb) {
    var key = tb.dataset.target;
    tb.addEventListener("click", function (e) {
      var b = e.target.closest("button");
      if (!b) return;
      var el = document.getElementById("flow-" + key);
      if (b.dataset.zoom && el) {
        var s = scales[key] || 1;
        if (b.dataset.zoom === "in") s = Math.min(s * 1.25, 4);
        else if (b.dataset.zoom === "out") s = Math.max(s / 1.25, 0.2);
        else s = 1;
        scales[key] = s;
        applyScale(el);
      }
    });
  });
})();
''',
    'style.css': r'''body {
  background: #f5f6f8;
}
.kpi-card .kpi-value {
  font-size: 2rem;
  font-weight: 700;
  line-height: 1.1;
}
.kpi-card .kpi-label {
  font-size: 0.85rem;
  color: #6c757d;
}
.kpi-card .kpi-sub {
  font-size: 0.8rem;
  color: #6c757d;
}
.step-list .step-done {
  text-decoration: line-through;
  color: #6c757d;
}
.suggestion-kind {
  font-size: 0.75rem;
}
.chart-box {
  position: relative;
  height: 260px;
}
.table-sm td, .table-sm th {
  vertical-align: middle;
}
.badge-flag {
  background: #e7f1ff;
  color: #0a58ca;
  border: 1px solid #b6d4fe;
  font-weight: 500;
}
.flow-canvas {
  overflow: auto;
  border: 1px solid #dee2e6;
  border-radius: 0.5rem;
  background: #fff;
  min-height: 240px;
  max-height: 72vh;
  margin-top: 0.75rem;
}
.flow-inner {
  display: inline-block;
  padding: 12px;
  transform-origin: 0 0;
  transition: transform 0.15s ease;
}
.flow-inner svg {
  max-width: none !important;
  height: auto;
}
.flow-inner .node {
  cursor: pointer;
}
.legend-chip {
  display: inline-block;
  padding: 0.05rem 0.5rem;
  border-radius: 0.3rem;
  border: 1px solid #adb5bd;
  margin-right: 0.3rem;
  font-size: 0.78rem;
}
.lg-task { background: #e7f1ff; border-color: #0d6efd; }
.lg-done { background: #d1e7dd; border-color: #198754; }
.lg-source { background: #fff3cd; border-color: #ffc107; }
.lg-system { background: #f8f9fa; border-color: #6c757d; }
.lg-output { background: #fff; border-color: #0d6efd; border-style: dashed; }
.lg-wp { background: #e7f1ff; border-color: #0d6efd; border-radius: 1rem; }
.lg-ext { background: #fff3cd; border-color: #fd7e14; border-radius: 1rem; }
.lg-goods { background: #e2f0d9; border-color: #198754; }
.lg-place { background: #f1f3f5; border-color: #adb5bd; border-radius: 1rem; }
.lg-warn { background: #f8d7da; border-color: #dc3545; }
/* ---- 業務マップ 編集エディタ ---- */
.fe-wrap { display: flex; flex-direction: column; height: calc(100vh - 150px); min-height: 520px; }
.fe-toolbar { display: flex; flex-wrap: wrap; justify-content: space-between; align-items: center; gap: 0.5rem; margin-bottom: 0.4rem; }
.fe-buttons { display: flex; flex-wrap: wrap; align-items: center; gap: 0.3rem; }
.fe-sep { width: 1px; height: 1.4rem; background: #dee2e6; margin: 0 0.2rem; }
.fe-status { font-size: 0.8rem; margin-left: 0.6rem; }
.fe-hint { font-size: 0.78rem; color: #6c757d; margin-bottom: 0.3rem; }
.fe-legend { display: inline-block; margin-left: 0.5rem; }
.fe-edge-legend { margin-left: 0.4rem; }
.fe-line { display: inline-block; width: 22px; height: 0; border-top: 2px solid #0d6efd; vertical-align: middle; margin: 0 0.2rem 0 0.4rem; }
.fe-line-goods { border-top-width: 4px; border-top-color: #198754; }
.fe-line-seq { border-top-style: dashed; border-top-color: #6c757d; }
.fe-line-manual { border-top-style: dashed; border-top-color: #6f42c1; }
.fe-mode { display: none; font-size: 0.85rem; background: #e7f1ff; color: #052c65; border: 1px solid #b6d4fe; border-radius: 0.4rem; padding: 0.25rem 0.6rem; margin-bottom: 0.3rem; }
.fe-container { position: relative; flex: 1; overflow: hidden; border: 1px solid #dee2e6; border-radius: 0.5rem; background: #fff;
  background-image: linear-gradient(#f1f3f5 1px, transparent 1px), linear-gradient(90deg, #f1f3f5 1px, transparent 1px); background-size: 24px 24px; cursor: grab; }
.fe-container.fe-panning { cursor: grabbing; }
.fe-canvas { position: absolute; left: 0; top: 0; transform-origin: 0 0; }
.fe-canvas.fe-linking .fe-node { cursor: crosshair; }
.fe-svg { position: absolute; left: 0; top: 0; width: 1px; height: 1px; overflow: visible; pointer-events: none; }
.fe-svg .fe-edge-group { pointer-events: auto; cursor: pointer; }
.fe-hit { fill: none; stroke: transparent; stroke-width: 16px; }
.fe-edge { fill: none; stroke: #0d6efd; stroke-width: 2px; }
.fe-edge-group:hover .fe-edge { stroke-width: 3.5px; }
.fe-kind-goods { stroke: #198754; stroke-width: 3.5px; }
.fe-kind-seq { stroke: #6c757d; stroke-dasharray: 6 4; }
.fe-kind-relation { stroke: #adb5bd; stroke-width: 1.5px; stroke-dasharray: 2 4; }
.fe-kind-manual, .fe-src-manual { stroke: #6f42c1; stroke-dasharray: 8 4; }
.fe-edge-label { font-size: 11px; fill: #343a40; paint-order: stroke; stroke: #fff; stroke-width: 3px; pointer-events: none; }
.fe-temp-line { fill: none; stroke: #60a5fa; stroke-width: 2px; stroke-dasharray: 4 4; }
#fe-nodes { position: absolute; left: 0; top: 0; }
.fe-node { position: absolute; min-width: 90px; max-width: 230px; padding: 6px 10px; border: 2px solid #0d6efd; border-radius: 8px; background: #fff;
  font-size: 13px; line-height: 1.3; text-align: center; cursor: move; user-select: none; box-shadow: 0 2px 4px rgba(0,0,0,0.08); }
.fe-node:hover { box-shadow: 0 4px 10px rgba(0,0,0,0.15); z-index: 5; }
.fe-node.fe-selected { outline: 3px solid rgba(25, 135, 84, 0.45); outline-offset: 1px; z-index: 6; }
.fe-node.fe-link-source { outline: 3px solid rgba(246, 173, 85, 0.8); outline-offset: 1px; }
.fe-node.fe-manual { border-style: dashed; }
.fe-node .fe-text { white-space: pre-wrap; word-break: break-word; min-width: 40px; }
.fe-node .fe-text.fe-editing { outline: 1px solid #0d6efd; background: #fff; cursor: text; user-select: text; }
.fe-node .fe-sub { font-size: 10px; color: #6c757d; margin-top: 2px; }
.fe-node .fe-edited { position: absolute; right: 3px; bottom: 1px; font-size: 9px; color: #6f42c1; }
.fe-btns { display: none; position: absolute; top: -26px; left: 50%; transform: translateX(-50%); gap: 3px; white-space: nowrap; z-index: 7; }
.fe-node:hover .fe-btns, .fe-node.fe-selected .fe-btns { display: flex; }
.fe-btn { width: 22px; height: 22px; border: 1px solid #ced4da; border-radius: 50%; background: #fff; font-size: 11px; line-height: 1; padding: 0; cursor: pointer; box-shadow: 0 1px 2px rgba(0,0,0,0.15); }
.fe-btn:hover { transform: scale(1.1); }
.fe-btn-del { color: #dc3545; }
.fe-cls-task { background: #e7f1ff; border-color: #0d6efd; color: #052c65; }
.fe-cls-done { background: #d1e7dd; border-color: #198754; color: #0f5132; }
.fe-cls-source { background: #fff3cd; border-color: #ffc107; color: #664d03; border-radius: 16px; }
.fe-cls-system { background: #f8f9fa; border-color: #6c757d; color: #343a40; border-radius: 4px; }
.fe-cls-output { background: #fff; border-color: #0d6efd; border-style: dashed; color: #052c65; transform: skewX(-6deg); }
.fe-cls-output .fe-text { transform: skewX(6deg); }
.fe-cls-person { background: #fde2e4; border-color: #dc3545; color: #842029; border-radius: 20px; }
.fe-cls-wp { background: #e7f1ff; border-color: #0d6efd; color: #052c65; border-radius: 20px; }
.fe-cls-ext { background: #fff3cd; border-color: #fd7e14; color: #664d03; border-radius: 20px; }
.fe-cls-goods { background: #e2f0d9; border-color: #198754; color: #0f5132; border-radius: 12px; }
.fe-cls-place { background: #f1f3f5; border-color: #adb5bd; color: #343a40; border-radius: 20px; }
.fe-cls-warn { background: #f8d7da; border-color: #dc3545; color: #842029; border-radius: 20px; }
.fe-cls-note { background: #fff; border-color: #adb5bd; border-style: dashed; color: #6c757d; }
.fe-cls-manual { background: #e9d5ff; border-color: #6f42c1; color: #3b0764; }
.fe-node[data-color="white"] { background: #fff; }
.fe-node[data-color="red"] { background: #fee2e2; }
.fe-node[data-color="orange"] { background: #fed7aa; }
.fe-node[data-color="yellow"] { background: #fef3c7; }
.fe-node[data-color="green"] { background: #d1fae5; }
.fe-node[data-color="blue"] { background: #dbeafe; }
.fe-node[data-color="purple"] { background: #e9d5ff; }
.fe-node[data-color="pink"] { background: #fce7f3; }
.fe-lg-task { background: #e7f1ff; border-color: #0d6efd; }
.fe-lg-source { background: #fff3cd; border-color: #ffc107; }
.fe-lg-system { background: #f8f9fa; border-color: #6c757d; }
.fe-lg-output { background: #fff; border-color: #0d6efd; border-style: dashed; }
.fe-lg-person { background: #fde2e4; border-color: #dc3545; }
.fe-lg-wp { background: #e7f1ff; border-color: #0d6efd; border-radius: 1rem; }
.fe-lg-ext { background: #fff3cd; border-color: #fd7e14; border-radius: 1rem; }
.fe-lg-goods { background: #e2f0d9; border-color: #198754; }
.fe-lg-place { background: #f1f3f5; border-color: #adb5bd; border-radius: 1rem; }
.fe-lg-warn { background: #f8d7da; border-color: #dc3545; }
.fe-lg-note { background: #fff; border-color: #adb5bd; border-style: dashed; }
.fe-lg-manual { background: #e9d5ff; border-color: #6f42c1; }
.fe-selection { display: none; position: absolute; border: 2px dashed #0d6efd; background: rgba(13,110,253,0.08); pointer-events: none; z-index: 50; }
.fe-popover { display: none; position: absolute; z-index: 100; background: #fff; border: 1px solid #ced4da; border-radius: 0.5rem; padding: 0.6rem; width: 230px; box-shadow: 0 6px 16px rgba(0,0,0,0.18); font-size: 0.85rem; }
.fe-popover.fe-show { display: block; }
.fe-popover-wide { width: 520px; max-width: calc(100% - 32px); }
.fe-wb-list { max-height: 50vh; overflow: auto; }
.fe-wb-item { display: block; padding: 0.25rem 0; border-bottom: 1px solid #f1f3f5; cursor: pointer; }
.fe-pop-title { font-weight: 600; margin-bottom: 0.4rem; }
.fe-colors { display: flex; flex-wrap: wrap; gap: 5px; }
.fe-color { width: 24px; height: 24px; border-radius: 50%; border: 2px solid #dee2e6; cursor: pointer; display: inline-flex; align-items: center; justify-content: center; font-size: 12px; color: #6c757d; }
.fe-color:hover { border-color: #0d6efd; transform: scale(1.1); }
.fe-color[data-color="white"] { background: #fff; }
.fe-color[data-color="red"] { background: #fee2e2; }
.fe-color[data-color="orange"] { background: #fed7aa; }
.fe-color[data-color="yellow"] { background: #fef3c7; }
.fe-color[data-color="green"] { background: #d1fae5; }
.fe-color[data-color="blue"] { background: #dbeafe; }
.fe-color[data-color="purple"] { background: #e9d5ff; }
.fe-color[data-color="pink"] { background: #fce7f3; }
pre.notice {
  white-space: pre-wrap;
  background: #f8f9fa;
  border: 1px solid #dee2e6;
  padding: 1rem;
  border-radius: 0.5rem;
  font-family: inherit;
  font-size: 0.95rem;
}
/* ---- 進捗掲示板 ---- */
.pb-head { margin: 0.25rem 0 0.75rem; display: flex; flex-wrap: wrap; justify-content: space-between; align-items: flex-start; gap: 0.75rem 1.5rem; }
.pb-head-main { flex: 1 1 420px; min-width: 0; }
.pb-h1 { font-size: 1.75rem; font-weight: 700; letter-spacing: -0.01em; margin: 0; white-space: nowrap; }
.pb-link { color: var(--ui-accent); text-decoration: none; font-size: 0.9rem; white-space: nowrap; }
.pb-link:hover { text-decoration: underline; }
.pb-scope-form { display: inline-block; max-width: 100%; }
.pb-h1 { display: flex; align-items: baseline; gap: 0.1em; flex-wrap: wrap; }
.pb-scope-select { font: inherit; font-weight: 700; letter-spacing: inherit; color: var(--ui-accent); background: transparent; border: 0; border-bottom: 3px solid rgba(0, 113, 227, 0.3); border-radius: 0; padding: 0 1.15em 0 0; margin: 0; appearance: none; -webkit-appearance: none; cursor: pointer; max-width: 80vw; background-image: url("data:image/svg+xml;charset=utf-8,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 16 16'%3E%3Cpath fill='none' stroke='%230071e3' stroke-width='2.2' stroke-linecap='round' stroke-linejoin='round' d='M3 6l5 5 5-5'/%3E%3C/svg%3E"); background-repeat: no-repeat; background-position: right 0.05em center; background-size: 0.6em; field-sizing: content; }
.pb-scope-select:hover { border-bottom-color: var(--ui-accent); }
.pb-scope-select:focus { outline: none; border-bottom-color: var(--ui-accent); }
.pb-section { display: flex; align-items: baseline; gap: 0.4rem; font-size: 0.85rem; font-weight: 600; color: #6c757d; letter-spacing: 0.04em; margin: 1.5rem 0 0.6rem; }
.pb-section .pb-n { font-weight: 500; }
.pb-section .pb-link { margin-left: auto; font-weight: 400; letter-spacing: 0; }
.pb-row { display: grid; grid-template-columns: 170px minmax(220px, 1fr) minmax(460px, 2.2fr); gap: 1.25rem; align-items: center;
  background: #fff; border-radius: 14px; padding: 1rem 1.25rem; margin-bottom: 0.6rem; box-shadow: var(--ui-shadow); }
.pb-wp { font-weight: 600; font-size: 0.95rem; }
.pb-name { color: #6c757d; font-size: 0.9rem; }
.pb-title { font-weight: 600; font-size: 1.05rem; line-height: 1.35; }
.pb-title a { color: inherit; text-decoration: none; }
.pb-title a:hover { text-decoration: underline; }
.pb-note { color: #6c757d; font-size: 0.88rem; margin-top: 0.2rem; }
.pb-pill { display: inline-block; vertical-align: middle; font-size: 0.72rem; font-weight: 600; padding: 0.1rem 0.5rem; border-radius: 999px; background: #fff3cd; color: #997404; margin-left: 0.4rem; }
.pb-empty { color: #adb5bd; font-size: 0.9rem; padding: 0.5rem 0; }
/* 順番待ち */
.pb-waiting { list-style: none; margin: 0 0 1rem; padding: 0; background: #fff; border-radius: 14px; box-shadow: var(--ui-shadow); }
.pb-waiting li { display: flex; justify-content: space-between; gap: 1rem; padding: 0.65rem 1.25rem; border-top: 1px solid #f1f3f5; font-size: 0.95rem; }
.pb-waiting li:first-child { border-top: 0; }
.pb-waiting a { color: inherit; text-decoration: none; }
.pb-waiting a:hover { text-decoration: underline; }
.pb-dim { color: #6c757d; white-space: nowrap; font-size: 0.88rem; }
/* タブレット以下: 1列に積む */
@media (max-width: 991.98px) {
  .pb-row { grid-template-columns: 1fr; gap: 0.6rem; padding: 0.9rem 1rem; }
  .pb-from { display: flex; gap: 0.5rem; align-items: baseline; }
}
/* スマホ: ラベルは現在の段階だけ。横スクロールなしで収める */
@media (max-width: 575.98px) {
  .pb-h1 { font-size: 1.5rem; }
  .pb-waiting li { flex-direction: column; gap: 0.1rem; }
}
/* ============================================================
   デザインの基礎（書体・色・角丸・影）。Bootstrap の後に読み込まれ、既定値を上書きする
   ============================================================ */
:root {
  --ui-font: "Inter", "Noto Sans JP", -apple-system, BlinkMacSystemFont, "Segoe UI", "Hiragino Sans", "Hiragino Kaku Gothic ProN", "Yu Gothic UI", Meiryo, sans-serif;
  --ui-bg: #f5f5f7;
  --ui-surface: #ffffff;
  --ui-text: #1d1d1f;
  --ui-text-2: #6e6e73;
  --ui-text-3: #a1a1a6;
  --ui-line: #e5e5ea;
  --ui-fill: #f2f2f5;
  --ui-accent: #0071e3;
  --ui-accent-hover: #0062c3;
  --ui-accent-soft: #e5f0ff;
  --ui-accent-text: #0a63c9;
  --ui-green: #248a3d;
  --ui-green-soft: #e3f7ea;
  --ui-green-text: #1d7a3b;
  --ui-yellow-soft: #fff1cc;
  --ui-yellow-text: #8a6500;
  --ui-red-soft: #ffe5e7;
  --ui-red-text: #c0001f;
  --c-blue: #0a84ff;
  --c-green: #30d158;
  --c-orange: #ff9f0a;
  --c-purple: #bf5af2;
  --c-red: #ff453a;
  --c-teal: #5ac8fa;
  --c-indigo: #5e5ce6;
  --c-pink: #ff375f;
  --ui-radius: 16px;
  --ui-radius-sm: 10px;
  --ui-shadow: 0 1px 2px rgba(0, 0, 0, 0.04), 0 4px 16px rgba(0, 0, 0, 0.05);
  --bs-body-font-family: var(--ui-font);
  --bs-link-color: var(--ui-accent);
  --bs-link-hover-color: var(--ui-accent-hover);
  --bs-border-radius: var(--ui-radius-sm);
}
body {
  font-family: var(--ui-font);
  background: var(--ui-bg);
  color: var(--ui-text);
  line-height: 1.6;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  font-feature-settings: "palt" 1;
  text-rendering: optimizeLegibility;
}
pre, code, kbd, samp, input, textarea, select, .fe-node { font-feature-settings: normal; }
h1, h2, h3, h4, h5, h6, .h1, .h2, .h3, .h4, .h5, .h6 { font-weight: 700; letter-spacing: -0.02em; color: var(--ui-text); line-height: 1.3; }
h2 { font-size: 1.65rem; }
h4 { font-size: 1.2rem; }
h5, .card-title { font-size: 1.05rem; font-weight: 700; letter-spacing: -0.01em; }
h6 { font-size: 0.95rem; font-weight: 700; }
a { color: var(--ui-accent); text-decoration: none; }
a:hover { color: var(--ui-accent-hover); text-decoration: underline; }
.text-muted { color: var(--ui-text-2) !important; }
.text-green { color: var(--ui-green) !important; }
.text-blue { color: var(--ui-accent) !important; }
main.container { padding-top: 2rem !important; padding-bottom: 4rem !important; }
/* 横幅に追従する（Bootstrap の段階的な固定幅をやめ、1480px まで可変） */
.container { width: 100%; max-width: 1480px; padding-left: clamp(16px, 2.5vw, 40px); padding-right: clamp(16px, 2.5vw, 40px); }
.breadcrumb { font-size: 0.85rem; }
.breadcrumb-item + .breadcrumb-item::before { color: var(--ui-text-3); }
hr { border-color: var(--ui-line); opacity: 1; }

/* ナビ（白・半透明。いまのページは青で示す） */
.app-nav {
  background: rgba(255, 255, 255, 0.86);
  -webkit-backdrop-filter: saturate(180%) blur(20px);
  backdrop-filter: saturate(180%) blur(20px);
  border-bottom: 1px solid var(--ui-line);
  padding-top: 0.45rem;
  padding-bottom: 0.45rem;
}
.app-nav .navbar-brand { display: inline-flex; align-items: center; gap: 0.55rem; font-weight: 700; letter-spacing: -0.02em; color: var(--ui-text); font-size: 1rem; }
.app-nav .navbar-brand:hover { color: var(--ui-text); text-decoration: none; }
.brand-mark { display: inline-flex; align-items: center; justify-content: center; width: 26px; height: 26px; border-radius: 8px; background: linear-gradient(135deg, var(--c-blue), var(--c-indigo)); color: #fff; font-weight: 800; font-size: 0.95rem; font-family: "Inter", sans-serif; letter-spacing: 0; box-shadow: 0 1px 2px rgba(0, 0, 0, 0.15); }
.app-nav .nav-link { color: var(--ui-text); font-weight: 500; font-size: 0.9rem; padding: 0.4rem 0.65rem; border-radius: 999px; white-space: nowrap; }
.app-nav .nav-link:hover { color: var(--ui-text); background: rgba(0, 0, 0, 0.05); text-decoration: none; }
.app-nav .nav-link.active { color: var(--ui-accent-text); background: var(--ui-accent-soft); font-weight: 600; }
.app-nav .navbar-toggler { border: 0; padding: 0.3rem 0.5rem; }
.app-nav .navbar-toggler:focus { box-shadow: none; }
@media (max-width: 1199.98px) {
  .app-nav .navbar-nav { padding: 0.5rem 0; }
  .app-nav .nav-link { border-radius: var(--ui-radius-sm); }
}

/* 見出し（セクション） */
.section-title { display: flex; align-items: baseline; gap: 0.6rem; font-size: 1.25rem; font-weight: 700; letter-spacing: -0.02em; margin: 2rem 0 0.8rem; }
.section-title:first-child { margin-top: 0; }
.section-sub { font-size: 0.9rem; font-weight: 500; color: var(--ui-text-2); letter-spacing: 0; }

/* カード */
.card { border: 0; border-radius: var(--ui-radius); box-shadow: var(--ui-shadow); background: var(--ui-surface); }
.card > .card-body { padding: 1.4rem 1.5rem; }
.card.bg-light { background: #fafafc !important; box-shadow: none; border: 1px solid var(--ui-line); }
.card .card-title { margin-bottom: 0.75rem; }
.list-group-item { border-color: #f0f0f3; padding-top: 0.7rem; padding-bottom: 0.7rem; }

/* 指標カード（色付きアイコン＋大きな数字） */
.stat .card-body { padding: 1.25rem 1.35rem 1.2rem; }
.stat-head { display: flex; align-items: center; gap: 0.65rem; margin-bottom: 0.9rem; }
.stat-tile { display: inline-flex; align-items: center; justify-content: center; width: 38px; height: 38px; border-radius: 11px; color: #fff; font-size: 1.15rem; flex: 0 0 auto; box-shadow: 0 2px 6px rgba(0, 0, 0, 0.12); }
.tile-blue { background: linear-gradient(135deg, #0a84ff, #4da3ff); }
.tile-green { background: linear-gradient(135deg, #28b24a, #34c759); }
.tile-indigo { background: linear-gradient(135deg, #5e5ce6, #7d7bff); }
.tile-orange { background: linear-gradient(135deg, #ff9500, #ffb340); }
.tile-teal { background: linear-gradient(135deg, #32ade6, #64d2ff); }
.tile-pink { background: linear-gradient(135deg, #ff2d55, #ff6482); }
.tile-purple { background: linear-gradient(135deg, #af52de, #c77dff); }
.tile-red { background: linear-gradient(135deg, #ff3b30, #ff6961); }
.stat-label { font-size: 0.9rem; font-weight: 600; color: var(--ui-text); line-height: 1.25; }
.stat-value { font-size: 2.4rem; font-weight: 700; letter-spacing: -0.04em; line-height: 1; color: var(--ui-text); font-variant-numeric: tabular-nums; }
.stat-unit { font-size: 0.95rem; font-weight: 600; color: var(--ui-text-2); margin-left: 0.3rem; letter-spacing: 0; }
.stat-sub { font-size: 0.8rem; color: var(--ui-text-2); margin-top: 0.6rem; line-height: 1.45; }
.mini-value { font-size: 1.5rem; font-weight: 700; letter-spacing: -0.03em; font-variant-numeric: tabular-nums; }
.kpi-card .kpi-value { font-size: 2.1rem; font-weight: 700; letter-spacing: -0.03em; font-variant-numeric: tabular-nums; color: var(--ui-text); }
.kpi-card .kpi-label, .kpi-label { font-size: 0.78rem; font-weight: 600; color: var(--ui-text-2); }
.kpi-card .kpi-sub { color: var(--ui-text-3); font-size: 0.78rem; }
@media (max-width: 575.98px) {
  .stat .card-body { padding: 1rem 1.05rem 0.95rem; }
  .stat-head { gap: 0.5rem; margin-bottom: 0.6rem; }
  .stat-tile { width: 32px; height: 32px; border-radius: 9px; font-size: 1rem; }
  .stat-label { font-size: 0.8rem; }
  .stat-value { font-size: 1.9rem; }
  .stat-unit { font-size: 0.8rem; }
}

/* ボタン */
.btn { border-radius: var(--ui-radius-sm); font-weight: 600; letter-spacing: 0; }
.btn-sm { border-radius: 8px; }
.btn:focus-visible { box-shadow: 0 0 0 4px rgba(0, 113, 227, 0.2); }
.btn-primary { --bs-btn-bg: var(--ui-accent); --bs-btn-border-color: var(--ui-accent); --bs-btn-hover-bg: var(--ui-accent-hover); --bs-btn-hover-border-color: var(--ui-accent-hover); --bs-btn-active-bg: #00529f; --bs-btn-active-border-color: #00529f; --bs-btn-disabled-bg: #9cc3ef; --bs-btn-disabled-border-color: #9cc3ef; }
.btn-outline-primary { --bs-btn-color: var(--ui-accent-text); --bs-btn-border-color: #bfd6f5; --bs-btn-hover-bg: var(--ui-accent-soft); --bs-btn-hover-color: var(--ui-accent-text); --bs-btn-hover-border-color: #9fc0ec; --bs-btn-active-bg: #d6e6fb; --bs-btn-active-color: var(--ui-accent-text); --bs-btn-active-border-color: #9fc0ec; }
.btn-outline-secondary { --bs-btn-color: var(--ui-text); --bs-btn-border-color: #d9d9de; --bs-btn-hover-bg: var(--ui-fill); --bs-btn-hover-color: var(--ui-text); --bs-btn-hover-border-color: #c9c9cf; --bs-btn-active-bg: #e8e8ec; --bs-btn-active-color: var(--ui-text); --bs-btn-active-border-color: #c9c9cf; }
.btn-outline-danger { --bs-btn-color: var(--ui-red-text); --bs-btn-border-color: #f3c3c9; --bs-btn-hover-bg: var(--ui-red-soft); --bs-btn-hover-color: var(--ui-red-text); --bs-btn-hover-border-color: #e9a3ab; }
.btn-success { --bs-btn-bg: var(--ui-green); --bs-btn-border-color: var(--ui-green); --bs-btn-hover-bg: #1e7a35; --bs-btn-hover-border-color: #1e7a35; --bs-btn-active-bg: #1a6b2f; --bs-btn-active-border-color: #1a6b2f; }

/* 入力欄（塗りつぶし型。フォーカスで白くなる） */
.form-control, .form-select { border-radius: var(--ui-radius-sm); border: 1px solid transparent; background-color: var(--ui-fill); color: var(--ui-text); padding: 0.55rem 0.85rem; transition: background-color 0.15s, border-color 0.15s, box-shadow 0.15s; }
.form-select { background-color: var(--ui-fill); }
.form-control:hover, .form-select:hover { background-color: #ececf0; }
.form-control:focus, .form-select:focus { background-color: #fff; border-color: var(--ui-accent); box-shadow: 0 0 0 4px rgba(0, 113, 227, 0.15); }
.form-control.is-invalid, .form-select.is-invalid { border-color: var(--c-red); background-color: #fff; }
.form-control-sm, .form-select-sm { border-radius: 8px; padding: 0.35rem 0.65rem; }
.form-control[readonly] { background-color: #fafafc; color: var(--ui-text-2); }
.form-control::placeholder { color: var(--ui-text-3); }
.form-label { font-weight: 600; color: var(--ui-text); margin-bottom: 0.4rem; }
.form-text { color: var(--ui-text-3); }
.form-check-input { border-color: #c9c9cf; width: 1.05em; height: 1.05em; margin-top: 0.22em; }
.form-check-input:checked { background-color: var(--ui-accent); border-color: var(--ui-accent); }
.form-check-input:focus { border-color: var(--ui-accent); box-shadow: 0 0 0 4px rgba(0, 113, 227, 0.15); }
.form-check-label { color: var(--ui-text); }

/* 表 */
.table { --bs-table-border-color: #ececf0; --bs-table-hover-bg: #f7f7fa; color: var(--ui-text); }
.table > thead th { font-size: 0.76rem; font-weight: 600; color: var(--ui-text-2); letter-spacing: 0.02em; border-bottom: 1px solid var(--ui-line); white-space: nowrap; }
.table > :not(caption) > * > * { padding: 0.6rem 0.65rem; }
.table-sm > :not(caption) > * > * { padding: 0.45rem 0.55rem; }
.table th { font-weight: 600; }
.table td a { font-weight: 500; }

/* バッジ（淡い塗り） */
.badge { font-weight: 600; border-radius: 999px; padding: 0.32em 0.7em; letter-spacing: 0; }
.badge.fs-6 { font-size: 0.85rem !important; padding: 0.35em 0.85em; }
.badge.text-bg-primary { background: var(--ui-accent-soft) !important; color: var(--ui-accent-text) !important; }
.badge.text-bg-success { background: var(--ui-green-soft) !important; color: var(--ui-green-text) !important; }
.badge.text-bg-secondary { background: #ececf0 !important; color: #4b4b50 !important; }
.badge.text-bg-warning { background: var(--ui-yellow-soft) !important; color: var(--ui-yellow-text) !important; }
.badge.text-bg-danger { background: var(--ui-red-soft) !important; color: var(--ui-red-text) !important; }
.badge.text-bg-dark { background: var(--ui-text) !important; color: #fff !important; }
.badge.text-bg-light { background: var(--ui-fill) !important; color: var(--ui-text-2) !important; }
.badge-flag { background: var(--ui-accent-soft); color: var(--ui-accent-text); border: 0; }
.suggestion-kind { font-weight: 600; }

/* アラート・その他 */
.alert { border: 0; border-radius: 12px; }
.alert-light { background: #fff; box-shadow: var(--ui-shadow); color: var(--ui-text); }
.alert-light.border { border: 0 !important; }
.alert-info { background: var(--ui-accent-soft); color: var(--ui-accent-text); }
.alert-warning { background: var(--ui-yellow-soft); color: var(--ui-yellow-text); }
.alert-success { background: var(--ui-green-soft); color: var(--ui-green-text); }
.alert-danger { background: var(--ui-red-soft); color: var(--ui-red-text); }
.alert-secondary { background: #ececf0; color: #4b4b50; }
td.heat-0, td.heat-1, td.heat-2, td.heat-3 { background: transparent; }
.heat-0 { background: var(--ui-fill); color: var(--ui-text-2); }
.heat-1 { background: var(--ui-red-soft); color: var(--ui-red-text); }
.heat-2 { background: var(--ui-yellow-soft); color: var(--ui-yellow-text); }
.heat-3 { background: var(--ui-green-soft); color: var(--ui-green-text); }
.flow-canvas { border-color: var(--ui-line); border-radius: var(--ui-radius-sm); }
pre.notice { background: #fafafc; border-color: var(--ui-line); border-radius: var(--ui-radius-sm); }
.btn-close:focus { box-shadow: none; }

/* 進捗掲示板（見出しの強弱） */
.pb-h1 { font-size: 2rem; letter-spacing: -0.03em; }
.pb-section { font-size: 1.1rem; font-weight: 700; color: var(--ui-text); letter-spacing: -0.02em; margin: 1.75rem 0 0.7rem; }
.pb-section .pb-n { color: var(--ui-text-3); font-weight: 600; font-size: 1rem; }
.pb-row { box-shadow: var(--ui-shadow); border-radius: var(--ui-radius); }
.pb-title { font-size: 1.08rem; }
.pb-waiting { box-shadow: var(--ui-shadow); border-radius: var(--ui-radius); }
@media (max-width: 575.98px) {
  .pb-h1 { font-size: 1.6rem; }
}

/* ============================================================
   動き（控えめに。OS の「視差効果を減らす」設定では止める）
   ============================================================ */
@keyframes ui-fade-up { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: none; } }
@keyframes ui-pop { from { opacity: 0; transform: scale(0.6); } to { opacity: 1; transform: scale(1); } }
@keyframes ui-draw { from { transform: scaleX(0); } to { transform: scaleX(1); } }
@keyframes ui-pulse { 0% { transform: scale(1); opacity: 0.55; } 70% { transform: scale(1.55); opacity: 0; } 100% { transform: scale(1.55); opacity: 0; } }
html { scroll-behavior: smooth; }
.section-title, .card, .alert, .pb-head, .pb-section, .pb-row, .pb-waiting, .pb-empty {
  animation: ui-fade-up 0.55s cubic-bezier(0.2, 0.7, 0.2, 1) both;
}
.row > div:nth-child(2) > .card { animation-delay: 0.06s; }
.row > div:nth-child(3) > .card { animation-delay: 0.12s; }
.row > div:nth-child(4) > .card { animation-delay: 0.18s; }
.pb-row:nth-of-type(2) { animation-delay: 0.05s; }
.pb-row:nth-of-type(3) { animation-delay: 0.1s; }
.pb-row:nth-of-type(4) { animation-delay: 0.15s; }
.pb-row:nth-of-type(5) { animation-delay: 0.2s; }
.pb-row:nth-of-type(6) { animation-delay: 0.25s; }
.pb-row:nth-of-type(n+7) { animation-delay: 0.3s; }
/* いまの段階はゆっくり脈打つ */

/* 触れたときの反応 */
.card, .pb-row, .pb-waiting li, .btn, .nav-link, .form-control, .form-select, .badge, .list-group-item, .table > tbody > tr {
  transition: background-color 0.18s ease, color 0.18s ease, border-color 0.18s ease, box-shadow 0.22s ease, transform 0.18s ease;
}
.pb-row:hover { transform: translateY(-2px); box-shadow: 0 8px 24px rgba(0, 0, 0, 0.09); }
.pb-waiting li:hover { background: #fafafc; }
.btn:active { transform: scale(0.97); }
.btn-primary:hover, .btn-success:hover { box-shadow: 0 4px 14px rgba(0, 0, 0, 0.14); }
.stat-tile { transition: transform 0.25s cubic-bezier(0.3, 1.4, 0.4, 1); }
.stat:hover .stat-tile { transform: scale(1.08) rotate(-3deg); }
.table-hover > tbody > tr:hover > * { transition: background-color 0.15s; }
@media (prefers-reduced-motion: reduce) {
  *, *::before, *::after { animation: none !important; transition: none !important; scroll-behavior: auto !important; }
}

/* 進捗掲示板: 割合の帯と凡例 */
.pb-summary { max-width: 580px; margin-top: 0.7rem; }
.pb-bar { display: flex; gap: 3px; height: 10px; border-radius: 999px; overflow: hidden; background: var(--ui-fill); }
.pb-seg { display: block; flex: 0 1 0%; min-width: 8px; border-radius: 999px; transform-origin: left center; animation: ui-draw 0.8s cubic-bezier(0.2, 0.7, 0.2, 1) both; }
.pb-seg:nth-child(2) { animation-delay: 0.08s; }
.pb-seg:nth-child(3) { animation-delay: 0.16s; }
.pb-seg:nth-child(4) { animation-delay: 0.24s; }
.seg-done { background: var(--c-green); }
.seg-active { background: var(--c-blue); }
.seg-hold { background: #ffcc00; }
.seg-wait { background: #d1d1d6; }
.pb-legend { display: flex; flex-wrap: wrap; align-items: baseline; gap: 0.3rem 1.2rem; list-style: none; margin: 0.55rem 0 0; padding: 0; font-size: 0.9rem; color: var(--ui-text-2); }
.pb-legend b { color: var(--ui-text); font-weight: 700; font-variant-numeric: tabular-nums; margin: 0 0.15rem 0 0.35rem; font-size: 1rem; }
.pb-dotc { display: inline-block; width: 9px; height: 9px; border-radius: 50%; margin-right: 0.45rem; vertical-align: 0; }
.pb-total { margin-left: auto; }
@media (max-width: 575.98px) { .pb-legend { gap: 0.25rem 0.9rem; font-size: 0.85rem; } .pb-total { margin-left: 0; } }

/* ============================================================
   ステッパー（開始 → … → 完了！）: 区切りのある帯。済＝緑、いま＝青（保留＝黄）、これから＝灰
   ============================================================ */
.pb-steps { display: flex; gap: 4px; list-style: none; margin: 0; padding: 0; }
.pb-step { flex: 1 1 0; min-width: 0; text-align: center; }
.pb-track { display: block; height: 7px; border-radius: 999px; background: #e3e3e8; transform-origin: left center; }
.pb-step.is-done .pb-track { background: var(--c-green); }
.pb-step.is-current .pb-track { background: var(--c-blue); box-shadow: 0 0 0 3px rgba(10, 132, 255, 0.16); }
.pb-complete .pb-step.is-current .pb-track { background: var(--c-green); box-shadow: 0 0 0 3px rgba(48, 209, 88, 0.2); }
.pb-hold .pb-step.is-current .pb-track { background: #ffcc00; box-shadow: 0 0 0 3px rgba(255, 204, 0, 0.22); }
.pb-label { display: block; position: relative; left: 50%; transform: translateX(-50%); width: max-content; max-width: 10rem; margin-top: 0.5rem; font-size: 0.78rem; line-height: 1.2; color: var(--ui-text-3); white-space: nowrap; }
.pb-step.is-done .pb-label { color: var(--ui-text-2); }
.pb-step.is-current .pb-label { color: var(--ui-text); font-weight: 700; }
.pb-complete .pb-step.is-current .pb-label { color: var(--ui-green-text); }
.pb-hold .pb-step.is-current .pb-label { color: #8a6500; }
.pb-date { display: block; margin-top: 0.1rem; font-size: 0.66rem; line-height: 1.1; min-height: 0.8rem; color: var(--ui-text-3); font-variant-numeric: tabular-nums; }
.pb-step.is-current .pb-date { color: var(--ui-text-2); }
/* 案件一覧の小さい版 */
.pb-steps-sm { gap: 3px; width: 120px; margin-top: 0.2rem; }
.pb-steps-sm .pb-track { height: 5px; }
.pb-steps-sm .pb-label, .pb-steps-sm .pb-date { display: none; }
.pb-steps-sm .pb-step.is-current .pb-track { box-shadow: none; }
/* 掲示板の行では帯を少し太く */
.pb-row .pb-track { height: 8px; }
/* 動き: 済んだ区間が左から順に伸び、いまの区間がゆっくり呼吸する */
@keyframes ui-breathe { 0%, 100% { opacity: 1; } 50% { opacity: 0.5; } }
.pb-step.is-done .pb-track, .pb-step.is-current .pb-track { animation: ui-draw 0.45s cubic-bezier(0.2, 0.7, 0.2, 1) both; }
.pb-step:nth-child(2) .pb-track { animation-delay: 0.07s; }
.pb-step:nth-child(3) .pb-track { animation-delay: 0.14s; }
.pb-step:nth-child(4) .pb-track { animation-delay: 0.21s; }
.pb-step:nth-child(5) .pb-track { animation-delay: 0.28s; }
.pb-step:nth-child(6) .pb-track { animation-delay: 0.35s; }
.pb-step:nth-child(7) .pb-track { animation-delay: 0.42s; }
.pb-steps:not(.pb-complete):not(.pb-steps-sm) .pb-step.is-current .pb-track { animation: ui-draw 0.45s cubic-bezier(0.2, 0.7, 0.2, 1) both, ui-breathe 2.8s ease-in-out 1.2s infinite; }
.pb-steps-sm .pb-track { animation: none !important; }
@media (max-width: 575.98px) {
  .pb-track { height: 6px; }
  .pb-row .pb-track { height: 6px; }
  .pb-label, .pb-date { display: none; }
  .pb-step.is-current .pb-label { display: block; font-size: 0.74rem; }
  .pb-step.is-current .pb-date { display: block; }
}
@media (prefers-reduced-motion: reduce) {
  .pb-track { animation: none !important; }
}

/* ダッシュボードの1行リンク・保留ボタン */
.dash-link { margin: -0.5rem 0 1.5rem; font-size: 0.95rem; }
.dash-link a { font-weight: 600; text-decoration: none; }
.dash-link .bi-kanban { margin-right: 0.4rem; }
.btn-outline-warning { --bs-btn-color: #8a6500; --bs-btn-border-color: #f1d27a; --bs-btn-hover-bg: var(--ui-yellow-soft); --bs-btn-hover-color: #8a6500; --bs-btn-hover-border-color: #e3bd4f; --bs-btn-active-bg: #ffe9a8; --bs-btn-active-color: #8a6500; --bs-btn-active-border-color: #e3bd4f; }

/* ---- ログイン ---- */
.login-wrap { max-width: 440px; margin: 2rem auto 0; }
.login-card .card-body { padding: 2.25rem 2rem; }
.login-mark { width: 56px; height: 56px; border-radius: 16px; background: linear-gradient(135deg, var(--c-blue), var(--c-indigo)); color: #fff; display: flex; align-items: center; justify-content: center; font-size: 1.5rem; margin: 0 auto 1rem; }
.login-title { font-size: 1.5rem; text-align: center; margin-bottom: 0.25rem; }
.login-sub { color: var(--ui-text-2); font-size: 0.9rem; text-align: center; margin-bottom: 1.25rem; }
/* ---- 月次アンケート（回答） ---- */
.mq-wrap { max-width: 720px; margin: 0 auto; }
.mq-head { text-align: center; margin: 0.5rem 0 1.25rem; }
.mq-badge { display: inline-block; padding: 0.2rem 0.8rem; border-radius: 999px; background: var(--ui-accent-soft); color: var(--ui-accent-text); font-weight: 600; font-size: 0.85rem; }
.mq-title { font-size: 1.9rem; letter-spacing: -0.03em; margin: 0.5rem 0 0.35rem; }
.mq-sub { color: var(--ui-text-2); font-size: 0.95rem; margin: 0; }
.mq-card .card-body { padding: 1.5rem 1.75rem; }
.mq-identity { padding-bottom: 1.25rem; margin-bottom: 0.5rem; border-bottom: 1px solid var(--ui-line); }
.mq-q { padding: 1.1rem 0; border-bottom: 1px solid var(--ui-line); }
.mq-q:last-of-type { border-bottom: 0; }
.mq-q-head { display: flex; align-items: baseline; gap: 0.6rem; margin-bottom: 0.7rem; }
.mq-no { display: inline-flex; align-items: center; justify-content: center; min-width: 1.7rem; height: 1.7rem; border-radius: 50%; background: var(--ui-text); color: #fff; font-size: 0.8rem; font-weight: 700; flex: 0 0 auto; align-self: center; }
.mq-text { font-weight: 600; font-size: 1.02rem; line-height: 1.4; }
.mq-optional { font-size: 0.72rem; color: var(--ui-text-3); border: 1px solid var(--ui-line); border-radius: 4px; padding: 0 0.35rem; align-self: center; white-space: nowrap; }
.mq-required { font-size: 0.72rem; color: var(--ui-red-text); background: var(--ui-red-soft); border-radius: 4px; padding: 0 0.35rem; align-self: center; white-space: nowrap; font-weight: 700; }
.mq-options { display: flex; flex-wrap: wrap; gap: 0.5rem; }
.mq-option { position: relative; cursor: pointer; margin: 0; }
.mq-option input { position: absolute; opacity: 0; width: 0; height: 0; }
.mq-option span { display: inline-block; padding: 0.55rem 1rem; border-radius: 999px; background: var(--ui-fill); color: var(--ui-text); font-size: 0.95rem; font-weight: 500; border: 1.5px solid transparent; transition: all 0.15s ease; user-select: none; }
.mq-option:hover span { background: #e6e6eb; }
.mq-option input:checked + span { background: var(--ui-accent); color: #fff; box-shadow: 0 2px 8px rgba(0, 113, 227, 0.3); }
.mq-option input:focus-visible + span { box-shadow: 0 0 0 4px rgba(0, 113, 227, 0.2); }
.mq-err:empty { display: none; }
.mq-err { margin-top: 0.4rem; }
.mq-done .card-body { padding: 2.5rem 1.75rem; }
.mq-check { width: 64px; height: 64px; border-radius: 50%; background: var(--c-green); color: #fff; display: flex; align-items: center; justify-content: center; font-size: 2rem; margin: 0 auto 1rem; animation: ui-pop 0.5s cubic-bezier(0.3, 1.4, 0.4, 1) both; }
@media (max-width: 575.98px) {
  .mq-card .card-body { padding: 1.1rem 1rem; }
  .mq-title { font-size: 1.5rem; }
  .mq-option span { padding: 0.5rem 0.85rem; font-size: 0.9rem; }
}
/* ---- 月次アンケート（集計） ---- */
.mq-stat-head { display: flex; align-items: center; gap: 0.6rem; margin-bottom: 0.9rem; }
.mq-bars { display: flex; flex-direction: column; gap: 0.5rem; }
.mq-bar-row { display: grid; grid-template-columns: minmax(90px, 32%) 1fr minmax(80px, auto); gap: 0.75rem; align-items: center; font-size: 0.9rem; }
.mq-bar-label { color: var(--ui-text); }
.mq-bar { height: 10px; border-radius: 999px; background: var(--ui-fill); overflow: hidden; }
.mq-bar span { display: block; height: 100%; border-radius: 999px; background: var(--c-blue); transform-origin: left center; animation: ui-draw 0.7s cubic-bezier(0.2, 0.7, 0.2, 1) both; }
.mq-bar-row:nth-child(2) .mq-bar span { background: #4da3ff; }
.mq-bar-row:nth-child(3) .mq-bar span { background: #8ec3ff; }
.mq-bar-row:nth-child(4) .mq-bar span { background: #b5d6ff; }
.mq-bar-row:nth-child(n+5) .mq-bar span { background: #d2e5fb; }
.mq-bar-val { text-align: right; font-variant-numeric: tabular-nums; font-weight: 600; white-space: nowrap; }
.mq-list { list-style: none; margin: 0; padding: 0; }
.mq-list li { padding: 0.5rem 0; border-top: 1px solid var(--ui-line); font-size: 0.92rem; }
.mq-list li:first-child { border-top: 0; padding-top: 0; }
/* ---- 設問エディタ ---- */
.qe-q { border: 1px solid var(--ui-line); box-shadow: none; }
.qe-handle { min-width: 1.6rem; font-weight: 700; }
/* ---- 自動化依頼（1画面） ---- */
.rq { max-width: 1080px; margin: 0 auto; }
.rq-top { display: flex; flex-wrap: wrap; align-items: center; justify-content: space-between; gap: 0.75rem 1.5rem; margin-bottom: 1rem; }
.rq-title { font-size: 1.6rem; margin: 0; }
.rq-no { display: inline-flex; align-items: center; justify-content: center; width: 1.5rem; height: 1.5rem; border-radius: 50%; background: #fff; color: var(--ui-text-2); font-size: 0.78rem; font-weight: 700; }
.rq-opt { font-size: 0.66rem; border: 1px solid currentColor; border-radius: 4px; padding: 0 0.3rem; opacity: 0.7; }
.rq-card .card-body { padding: 1.5rem 1.75rem; }
.rq-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 0 1rem; }
.rq-field { margin-bottom: 1.15rem; }
.rq-label { display: flex; align-items: center; gap: 0.45rem; font-weight: 600; margin-bottom: 0.45rem; font-size: 0.95rem; }
.rq-req { font-size: 0.66rem; color: var(--ui-red-text); background: var(--ui-red-soft); border-radius: 4px; padding: 0.05rem 0.35rem; font-weight: 700; }
.rq-pills { display: flex; flex-wrap: wrap; gap: 0.45rem; }
.rq-pill { border: 1.5px solid transparent; background: var(--ui-fill); color: var(--ui-text); border-radius: 999px; padding: 0.45rem 0.95rem; font-size: 0.92rem; font-weight: 500; cursor: pointer; transition: all 0.15s ease; }
.rq-pill:hover { background: #e3e3e8; }
.rq-pill:active { transform: scale(0.97); }
.rq-pill.is-on { background: var(--ui-accent); color: #fff; box-shadow: 0 2px 8px rgba(0, 113, 227, 0.3); }
.rq-pill:focus-visible { outline: none; box-shadow: 0 0 0 4px rgba(0, 113, 227, 0.2); }
.rq-note { margin-top: 0.5rem; max-width: 440px; }
.rq-unit-wrap { display: flex; align-items: center; gap: 0.5rem; max-width: 180px; }
.rq-unit { color: var(--ui-text-2); }
.rq-field.is-invalid .form-control, .rq-field.is-invalid .form-select { border-color: var(--c-red); background: #fff; }
.rq-field.is-invalid .rq-pill:not(.is-on) { border-color: rgba(255, 69, 58, 0.45); }
.rq-err { color: var(--ui-red-text); font-size: 0.8rem; margin-top: 0.35rem; }
.rq-err:empty { display: none; }
.rq-section { padding-top: 1.1rem; margin-top: 0.4rem; border-top: 1px solid var(--ui-line); }
.rq-section:first-child { border-top: 0; margin-top: 0; padding-top: 0; }
.rq-section-title { display: flex; align-items: center; justify-content: space-between; gap: 0.5rem; font-size: 0.8rem; font-weight: 700; color: var(--ui-text-2); letter-spacing: 0.04em; margin-bottom: 0.9rem; }
.rq-section.is-collapsed .rq-section-body { display: none; }
.rq-section.is-collapsed .rq-section-title { margin-bottom: 0; }
.rq-toggle { background: none; border: 0; color: var(--ui-accent); font-weight: 600; font-size: 0.85rem; padding: 0; letter-spacing: 0; }
.rq-actions { display: flex; justify-content: flex-end; align-items: center; gap: 0.75rem; margin-top: 1rem; padding-top: 1rem; border-top: 1px solid var(--ui-line); flex-wrap: wrap; }
.rq-done .card-body { padding: 2.5rem 1.75rem; }
.rq-check { width: 64px; height: 64px; border-radius: 50%; background: var(--c-green); color: #fff; display: flex; align-items: center; justify-content: center; font-size: 2rem; margin: 0 auto 1rem; animation: ui-pop 0.5s cubic-bezier(0.3, 1.4, 0.4, 1) both; }
.rq-toast { position: fixed; left: 50%; bottom: 24px; transform: translateX(-50%) translateY(16px); background: var(--ui-text); color: #fff; padding: 0.6rem 1.1rem; border-radius: 999px; font-size: 0.9rem; opacity: 0; pointer-events: none; transition: all 0.25s ease; z-index: 2000; box-shadow: 0 8px 24px rgba(0, 0, 0, 0.2); max-width: calc(100% - 32px); }
.rq-toast.is-show { opacity: 1; transform: translateX(-50%) translateY(0); }
@media (max-width: 575.98px) { .rq-card .card-body { padding: 1.1rem 1rem; } .rq-title { font-size: 1.35rem; } .rq-pill { padding: 0.45rem 0.8rem; font-size: 0.88rem; } }
.rq-sec { margin-bottom: 1rem; }
.rq-sec-head { display: flex; flex-wrap: wrap; align-items: center; gap: 0.4rem 0.6rem; margin-bottom: 1.1rem; }
.rq-sec-head .rq-no { width: 1.9rem; height: 1.9rem; background: var(--ui-text); color: #fff; font-size: 0.9rem; }
.rq-sec-title { font-weight: 700; font-size: 1.1rem; letter-spacing: -0.01em; white-space: nowrap; }
.rq-sec-head .rq-opt { color: var(--ui-text-3); }
.rq-finish { display: flex; justify-content: center; margin: 1.5rem 0; }
.rq-finish .btn { min-width: 260px; }
.rq-done { animation: ui-fade-up 0.4s cubic-bezier(0.2, 0.7, 0.2, 1) both; }
.rq-id-row { display: flex; align-items: flex-end; gap: 1rem; flex-wrap: wrap; }
.rq-id-row .rq-field { margin-bottom: 0.5rem; min-width: 220px; }
.rq-identity { font-weight: 700; font-size: 1.05rem; color: var(--ui-green-text); padding-bottom: 0.9rem; }
.rq-identity.is-missing { color: var(--ui-text-2); font-weight: 500; font-size: 0.9rem; }
.rq-identity:empty { display: none; }
.rq-fallback { margin-top: 0.75rem; }
/* 右上の管理者メニュー（目立たせない） */
.nav-admin { background: none; border: 0; padding: 0.4rem 0.5rem; font-size: 0.78rem; color: var(--ui-text-3); text-decoration: none; white-space: nowrap; border-radius: 999px; }
.nav-admin:hover { color: var(--ui-text-2); background: rgba(0, 0, 0, 0.04); text-decoration: none; }
/* ---- ヘッダー1段: 管理者タブは藍色で区別 ---- */
.app-header { z-index: 1030; }
.app-nav .nav-link { font-size: 0.88rem; padding: 0.38rem 0.6rem; }
.nav-sep { width: 1px; height: 1.25rem; background: #d6d6dc; margin: 0 0.4rem 0 0.5rem; align-self: center; }
.nav-lock { align-self: center; color: #4f46e5; font-size: 0.8rem; margin-right: 0.1rem; display: inline-flex; align-items: center; gap: 0.3rem; }
.app-nav .nav-link.nav-link-admin { color: #4338ca; }
.app-nav .nav-link.nav-link-admin:hover { background: rgba(79, 70, 229, 0.08); color: #312e81; }
.app-nav .nav-link.nav-link-admin.active { background: #4f46e5; color: #fff; }
.nav-admin { margin-left: 0.5rem; }
@media (min-width: 1200px) and (max-width: 1399.98px) {
  .app-nav .nav-link { font-size: 0.8rem; padding: 0.35rem 0.4rem; }
  .app-nav .navbar-brand { font-size: 0.92rem; }
  .nav-sep { margin: 0 0.25rem; }
}
@media (max-width: 1199.98px) {
  .nav-sep { width: 100%; height: 1px; margin: 0.4rem 0; }
  .nav-lock { align-self: flex-start; font-size: 0.78rem; font-weight: 700; margin: 0.1rem 0 0.1rem 0.75rem; }
  .nav-lock::after { content: "管理者"; }
}

/* ---- DX人財情報 ---- */
.dx-table { min-width: 1120px; }
.dx-table th { white-space: nowrap; }
.dx-table .dx-wide { min-width: 11rem; }
.dx-table .dx-tasks { min-width: 16rem; }

/* ---- 月次集計（推移の表） ---- */
.trend-table th, .trend-table td { white-space: nowrap; font-variant-numeric: tabular-nums; font-size: 0.85rem; }
.trend-table .cell-sub { font-size: 0.72rem; color: var(--ui-text-3); }
.section-title form .form-select { font-weight: 500; }

/* ---- 進捗掲示板: 成果と依頼への導線 ---- */
.pb-wins { margin: 0.25rem 0 1rem; display: grid; gap: 0.75rem; animation: ui-fade-up 0.55s cubic-bezier(0.2, 0.7, 0.2, 1) both; }
.pb-strip { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 0.75rem; align-items: stretch; }
.pb-stat { min-width: 0; }
.pb-stat .card-body { padding: 1rem 1.15rem; }
.pb-stat .stat-head { margin-bottom: 0.7rem; }
.pb-stat .stat-value { font-size: 2rem; }
.pb-case { min-width: 0; }
.pb-case .card-body { padding: 0.9rem 1.1rem; }
.pb-case-name { font-weight: 700; letter-spacing: -0.01em; }
.pb-case-meta { font-size: 0.8rem; color: var(--ui-text-3); margin-top: 0.1rem; }
.pb-case-num { font-size: 0.88rem; color: var(--ui-text-2); margin-top: 0.45rem; font-variant-numeric: tabular-nums; }
.pb-case-save { color: var(--ui-green-text); margin-left: 0.4rem; font-weight: 700; }
.pb-case-note { font-size: 0.85rem; color: var(--ui-text); margin-top: 0.35rem; }
.pb-cta { flex: 0 1 auto; display: flex; flex-direction: column; gap: 0.6rem; background: var(--ui-accent); color: #fff; border-radius: var(--ui-radius); padding: 0.9rem 1.25rem; max-width: 100%; }
.pb-cta-actions { display: flex; flex-wrap: wrap; align-items: center; gap: 0.5rem 1rem; }
.pb-cta-text { font-weight: 700; letter-spacing: -0.01em; }
.pb-cta-btn { background: #fff; color: var(--ui-accent); font-weight: 700; border-radius: 999px; padding: 0.45rem 1.1rem; text-decoration: none; white-space: nowrap; }
.pb-cta-btn:hover { background: #eef4ff; color: var(--ui-accent); }
.pb-cta-link { color: #fff; opacity: 0.9; text-decoration: none; font-size: 0.9rem; white-space: nowrap; }
.pb-cta-link:hover { opacity: 1; text-decoration: underline; color: #fff; }

/* ---- 月次アンケート: 枝番つきの条件つき設問 ---- */
.mq-no { border-radius: 999px; padding: 0 0.5rem; }
.mq-q.mq-sub { margin-left: 1.1rem; padding-left: 1rem; border-left: 3px solid var(--ui-line); }
.mq-q.mq-sub .mq-no { background: var(--ui-text-2); }

/* ---- 業務マップ: 同じアウトプット（統合候補） ---- */
.lg-dupout { background: #fff3cd; border-color: #fd7e14; border-style: dashed; }
.fe-lg-dupout { background: #fff3cd; border-color: #fd7e14; border-style: dashed; }
.fe-cls-dupout { background: #fff3cd; border-color: #fd7e14; border-style: dashed; color: #664d03; transform: skewX(-6deg); }
.fe-cls-dupout .fe-text { transform: skewX(6deg); }
''',
    'survey.js': r'''// 自動化依頼（1枚のページ）。入力欄はサーバーの定義（rq-options）から描く。
// 「あなたのこと」「関わり方」「全体について」は入力すると自動で保存、業務は「登録」ボタン、最後に「送信して完了」。
(function () {
  var O = JSON.parse(document.getElementById("rq-options").textContent);
  var S = JSON.parse(document.getElementById("rq-state").textContent);
  var CFG = window.RQ_CONFIG || { api: {} };
  var toastTimer = null;
  var timers = {};

  // ---------- 小さな道具 ----------
  function $(id) { return document.getElementById(id); }
  function el(tag, cls, text) {
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text !== undefined && text !== null) e.textContent = text;
    return e;
  }
  async function api(method, url, body) {
    var r = await fetch(url, { method: method, headers: { "Content-Type": "application/json", Accept: "application/json" }, body: body === undefined ? undefined : JSON.stringify(body) });
    var j = {};
    try { j = await r.json(); } catch (e) { j = {}; }
    j._ok = r.ok;
    return j;
  }
  function toast(text) {
    var t = $("rq-toast");
    t.textContent = text;
    t.classList.add("is-show");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { t.classList.remove("is-show"); }, 2400);
  }
  function debounce(key, fn, ms) { clearTimeout(timers[key]); timers[key] = setTimeout(fn, ms || 700); }
  function optionsOf(f) {
    if (f.options === "workplaces") return O.workplaces.map(function (w) { return { value: String(w.id), label: w.name }; });
    if (f.options === "domains") return O.domains.map(function (d) { return { value: String(d.id), label: d.name }; });
    return (f.options || []).map(function (o) { return { value: o, label: o }; });
  }
  function isEmpty(v) { return v === "" || v === undefined || v === null || (Array.isArray(v) && v.length === 0); }

  // ---------- 入力欄の描画 ----------
  function renderField(f, onChange) {
    var wrap = el("div", "rq-field");
    wrap.dataset.key = f.key;
    wrap.dataset.type = f.type;
    var label = el("div", "rq-label");
    label.appendChild(el("span", null, f.label));
    if (f.required) label.appendChild(el("span", "rq-req", "必須"));
    wrap.appendChild(label);

    if (f.type === "text" || f.type === "number") {
      var input = el("input", "form-control");
      input.type = f.type === "number" ? "number" : "text";
      input.name = f.key;
      if (f.placeholder) input.placeholder = f.placeholder;
      if (f.type === "number") { input.min = f.min || 0; input.value = f.default !== undefined ? f.default : ""; input.classList.add("rq-num"); }
      if (f.lookup) input.autocomplete = "off";
      var box = el("div", f.unit ? "rq-unit-wrap" : "");
      box.appendChild(input);
      if (f.unit) box.appendChild(el("span", "rq-unit", f.unit));
      wrap.appendChild(box);
    } else if (f.type === "textarea") {
      var ta = el("textarea", "form-control");
      ta.name = f.key; ta.rows = 2;
      if (f.placeholder) ta.placeholder = f.placeholder;
      wrap.appendChild(ta);
    } else if (f.type === "select") {
      var sel = el("select", "form-select");
      sel.name = f.key;
      var first = el("option", null, "選択");
      first.value = "";
      sel.appendChild(first);
      optionsOf(f).forEach(function (o) { var op = el("option", null, o.label); op.value = o.value; sel.appendChild(op); });
      wrap.appendChild(sel);
    } else if (f.type === "choice" || f.type === "multi") {
      var pills = el("div", "rq-pills");
      optionsOf(f).forEach(function (o) {
        var b = el("button", "rq-pill", o.label);
        b.type = "button";
        b.dataset.value = o.value;
        b.addEventListener("click", function () {
          if (f.type === "choice") pills.querySelectorAll(".rq-pill").forEach(function (p) { if (p !== b) p.classList.remove("is-on"); });
          b.classList.toggle("is-on");
          clearError(wrap);
          if (onChange) onChange();
        });
        pills.appendChild(b);
      });
      wrap.appendChild(pills);
      if (f.note) {
        var note = el("input", "form-control form-control-sm rq-note");
        note.type = "text"; note.name = f.note.key; note.placeholder = f.note.placeholder || "";
        wrap.appendChild(note);
      }
    }
    wrap.appendChild(el("div", "rq-err"));
    wrap.addEventListener("input", function () { clearError(wrap); if (onChange) onChange(); });
    wrap.addEventListener("change", function () { if (onChange) onChange(); });
    return wrap;
  }
  function clearError(wrap) { wrap.classList.remove("is-invalid"); wrap.querySelector(".rq-err").textContent = ""; }
  function renderFields(container, fields, grid, onChange) {
    var host = grid ? el("div", "rq-grid") : container;
    fields.forEach(function (f) { host.appendChild(renderField(f, onChange)); });
    if (grid) container.appendChild(host);
  }
  function renderSections(container, sections) {
    sections.forEach(function (sec) {
      var s = el("div", "rq-section" + (sec.collapsed ? " is-collapsed" : ""));
      var head = el("div", "rq-section-title");
      head.appendChild(el("span", null, sec.title));
      if (sec.optional) {
        var tg = el("button", "rq-toggle", sec.collapsed ? "＋ 開く（任意）" : "－ 閉じる");
        tg.type = "button";
        tg.addEventListener("click", function () {
          s.classList.toggle("is-collapsed");
          tg.textContent = s.classList.contains("is-collapsed") ? "＋ 開く（任意）" : "－ 閉じる";
        });
        head.appendChild(tg);
      }
      s.appendChild(head);
      var body = el("div", "rq-section-body");
      var grid = sec.fields.every(function (f) { return f.type === "text" || f.type === "number" || f.type === "choice"; }) && sec.fields.length <= 3;
      renderFields(body, sec.fields, grid);
      s.appendChild(body);
      container.appendChild(s);
    });
  }

  // ---------- 値の読み書き ----------
  function getValues(container) {
    var v = {};
    container.querySelectorAll(".rq-field").forEach(function (w) {
      var key = w.dataset.key, type = w.dataset.type;
      if (type === "choice") { var on = w.querySelector(".rq-pill.is-on"); v[key] = on ? on.dataset.value : ""; }
      else if (type === "multi") { v[key] = Array.prototype.map.call(w.querySelectorAll(".rq-pill.is-on"), function (p) { return p.dataset.value; }); }
      else { var inp = w.querySelector("input, textarea, select"); v[key] = inp ? inp.value.trim() : ""; }
      var note = w.querySelector(".rq-note");
      if (note) v[note.name] = note.value.trim();
    });
    return v;
  }
  function setValues(container, obj) {
    obj = obj || {};
    container.querySelectorAll(".rq-field").forEach(function (w) {
      var key = w.dataset.key, type = w.dataset.type, val = obj[key];
      if (type === "choice" || type === "multi") {
        var vals = Array.isArray(val) ? val.map(String) : (isEmpty(val) ? [] : [String(val)]);
        w.querySelectorAll(".rq-pill").forEach(function (p) { p.classList.toggle("is-on", vals.indexOf(p.dataset.value) >= 0); });
      } else {
        var inp = w.querySelector("input, textarea, select");
        if (inp) inp.value = isEmpty(val) ? (inp.classList.contains("rq-num") ? "1" : "") : String(val);
      }
      var note = w.querySelector(".rq-note");
      if (note) note.value = obj[note.name] || "";
      clearError(w);
    });
  }
  function clearValues(container) { setValues(container, {}); }
  function showErrors(container, errors) {
    var first = null;
    Object.keys(errors || {}).forEach(function (k) {
      var w = container.querySelector('.rq-field[data-key="' + k + '"]');
      if (!w) return;
      w.classList.add("is-invalid");
      w.querySelector(".rq-err").textContent = errors[k];
      var sec = w.closest(".rq-section");
      if (sec && sec.classList.contains("is-collapsed")) sec.classList.remove("is-collapsed");
      if (!first) first = w;
    });
    if (first) first.scrollIntoView({ behavior: "smooth", block: "center" });
    return !!first;
  }
  function requiredErrors(container, fields) {
    var errors = {}, vals = getValues(container);
    fields.forEach(function (f) {
      if (f.required && isEmpty(vals[f.key])) errors[f.key] = (f.type === "text" || f.type === "textarea" || f.type === "number") ? "入力してください" : "選んでください";
    });
    return errors;
  }
  function hasAnyValue(container) {
    var vals = getValues(container);
    return Object.keys(vals).some(function (k) { return !isEmpty(vals[k]); });
  }
  var allTaskFields = O.task_sections.reduce(function (a, s) { return a.concat(s.fields); }, []);

  // ---------- 1. あなたのこと（社員IDだけ。氏名・職場は照会して表示） ----------
  var personBox = $("rq-person");
  var idField = O.person.filter(function (f) { return f.key === "emp_id"; })[0];
  var fallbackFields = O.person.filter(function (f) { return f.fallback; });
  var wpField = O.person.filter(function (f) { return f.key === "workplace_id"; })[0];
  var idRow = el("div", "rq-id-row");
  var idWrap = renderField(idField, null);
  idRow.appendChild(idWrap);
  var wpWrap = renderField(wpField, function () { debounce("person", function () { savePerson(false); }, 300); });
  idRow.appendChild(wpWrap);
  var identity = el("div", "rq-identity");
  idRow.appendChild(identity);
  personBox.appendChild(idRow);
  var fallback = el("div", "rq-fallback d-none");
  renderFields(fallback, fallbackFields, true, function () { debounce("person", function () { savePerson(false); }, 800); });
  personBox.appendChild(fallback);
  var empInput = idWrap.querySelector("input");
  var found = null;
  function setIdentity(text, missing) { identity.textContent = text || ""; identity.classList.toggle("is-missing", !!missing); }
  function fillFallback(name, workplaceId) {
    fallback.querySelector('input[name="name"]').value = name || "";
    var ws = personBox.querySelector('select[name="workplace_id"]');
    if (ws && workplaceId) ws.value = String(workplaceId);
  }
  async function lookupEmp() {
    var id = empInput.value.trim();
    found = null;
    if (!id) { setIdentity(""); fallback.classList.add("d-none"); return; }
    var j = await api("GET", CFG.api.lookup + encodeURIComponent(id));
    if (empInput.value.trim() !== id) return;
    if (j.found) {
      found = j;
      fillFallback(j.name, j.workplace_id);
      fallback.classList.toggle("d-none", !!j.name);
      setIdentity(j.name ? "" : "初回だけ氏名を入れてください", !j.name);
      if (j.name) debounce("person", function () { savePerson(false); }, 200);
    } else {
      fillFallback("", "");
      setIdentity("見つかりません。氏名を入れてください", true);
      fallback.classList.remove("d-none");
    }
  }
  empInput.addEventListener("input", function () { debounce("lookup", lookupEmp, 500); });
  empInput.addEventListener("change", lookupEmp);
  var personSaving = null;
  function savePerson(showErr) {
    if (personSaving) return personSaving;
    var v = getValues(personBox);
    var errors = {};
    if (!v.emp_id) errors.emp_id = "入力してください";
    if (!v.workplace_id) errors.workplace_id = "選んでください";
    if (!v.name && !(found && found.name)) errors.name = "入力してください";
    if (Object.keys(errors).length) {
      if (showErr) { if (errors.name) fallback.classList.remove("d-none"); showErrors(personBox, errors); }
      else if (found && !errors.emp_id) { var partial = {}; if (errors.workplace_id) partial.workplace_id = errors.workplace_id; if (errors.name) partial.name = errors.name; showErrors(personBox, partial); }
      return Promise.resolve(false);
    }
    personSaving = api("POST", CFG.api.respondent, { emp_id: v.emp_id, name: v.name, workplace_id: v.workplace_id }).then(function (j) {
      personSaving = null;
      if (!j._ok) { fallback.classList.remove("d-none"); showErrors(personBox, j.errors || {}); return false; }
      var changed = !S.respondent || S.respondent.id !== j.respondent.id;
      S.respondent = j.respondent;
      if (changed) {
        S.tasks = j.tasks || [];
        setValues(cBox, S.respondent);
        setValues(dBox, S.respondent);
        if (j.resumed && S.tasks.length) toast("前回の続きから登録できます");
      }
      setIdentity("");
      return true;
    });
    return personSaving;
  }

  // ---------- 2. 依頼する業務 ----------
  var taskBox = $("rq-task");
  renderSections(taskBox, O.task_sections);
  $("rq-save-task").addEventListener("click", async function () {
    var btn = this;
    if (!S.respondent) {
      var ok = await savePerson(true);
      if (!ok) { toast("先に「あなたのこと」を入力してください"); return; }
    }
    var errors = requiredErrors(taskBox, allTaskFields);
    if (showErrors(taskBox, errors)) return;
    btn.disabled = true;
    var j = await api("POST", CFG.api.tasks, getValues(taskBox));
    btn.disabled = false;
    if (!j._ok) { showErrors(taskBox, j.errors || {}); if (j.message) toast(j.message); return; }
    S.tasks = j.tasks || S.tasks;
    clearValues(taskBox);
    toast("登録しました：" + j.task.name + "（月 " + j.task.hours + " h）");
    $("sec-task").scrollIntoView({ behavior: "smooth", block: "start" });
  });

  // ---------- 3. 改善への関わり方（自動保存） ----------
  var cBox = $("rq-c");
  renderFields(cBox, O.c_fields, false, function () { debounce("c", function () { saveC(false); }, 800); });
  async function saveC(showErr) {
    if (!S.respondent || !hasAnyValue(cBox)) return true;
    var errors = requiredErrors(cBox, O.c_fields);
    if (Object.keys(errors).length) { if (showErr) showErrors(cBox, errors); return false; }
    var j = await api("POST", CFG.api.partC, getValues(cBox));
    if (!j._ok) { showErrors(cBox, j.errors || {}); return false; }
    S.respondent = j.respondent || S.respondent;
    return true;
  }

  // ---------- 4. 部全体について（自動保存） ----------
  var dBox = $("rq-d");
  renderFields(dBox, O.d_fields, false, function () { debounce("d", function () { saveD(); }, 900); });
  async function saveD() {
    if (!S.respondent || !hasAnyValue(dBox)) return true;
    var j = await api("POST", CFG.api.partD, getValues(dBox));
    if (!j._ok) { return false; }
    S.respondent = j.respondent || S.respondent; S.tasks = j.tasks || S.tasks;
    return true;
  }

  // ---------- 送信して完了 ----------
  $("rq-finish").addEventListener("click", async function () {
    var btn = this; btn.disabled = true;
    try {
      if (!S.respondent || Object.keys(requiredErrors(personBox, O.person)).length) {
        var ok = await savePerson(true);
        if (!ok) { toast("「あなたのこと」を入力してください"); return; }
      }
      if (!S.tasks.length) {
        toast("業務を1件以上登録してください");
        $("sec-task").scrollIntoView({ behavior: "smooth", block: "start" });
        return;
      }
      if (!(await saveC(true))) return;
      var j = await api("POST", CFG.api.partD, getValues(dBox));
      if (!j._ok) { toast(j.message || "送信できませんでした"); return; }
      S.respondent = j.respondent || S.respondent; S.tasks = j.tasks || S.tasks;
      var total = S.tasks.reduce(function (a, t) { return a + (t.monthly_minutes || 0); }, 0);
      $("rq-done-sum").textContent = "ありがとうございました。登録した業務 " + S.tasks.length + " 件（月 約 " + (total / 60).toFixed(1) + " 時間）は改善案件として確認し、活動状況に載せます。";
      var done = $("rq-done");
      done.classList.remove("d-none");
      done.scrollIntoView({ behavior: "smooth", block: "center" });
      toast("送信しました");
    } finally {
      btn.disabled = false;
    }
  });

  // ---------- 初期表示 ----------
  if (S.respondent) {
    empInput.value = S.respondent.emp_id || "";
    found = { name: S.respondent.name, workplace: S.respondent.workplace, workplace_id: S.respondent.workplace_id };
    fillFallback(S.respondent.name, S.respondent.workplace_id);
    setIdentity("");
    setValues(cBox, S.respondent);
    setValues(dBox, S.respondent);
  } else if (CFG.emp && empInput) {
    empInput.value = CFG.emp;
    lookupEmp();
  }
})();
''',
}


# ==============================================================================
# アプリ生成
# ==============================================================================
def _persistent_secret(instance_path):
    """セッション用の秘密鍵を instance フォルダに保存して再起動後も使い回す。"""
    path = os.path.join(instance_path, "secret_key.txt")
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            value = f.read().strip()
            if value:
                return value
    value = secrets.token_hex(32)
    with open(path, "w", encoding="utf-8") as f:
        f.write(value)
    return value


def create_app(test_config=None):
    app = Flask(__name__, static_folder=None, instance_path=os.path.join(BASE_DIR, "instance"), instance_relative_config=True)
    os.makedirs(app.instance_path, exist_ok=True)
    app.jinja_loader = DictLoader(TEMPLATES)

    db_path = os.path.join(app.instance_path, "portal.sqlite3").replace("\\", "/")
    app.config.update(
        SECRET_KEY=str(conf.get("SECRET_KEY", "") or "") or _persistent_secret(app.instance_path),
        SQLALCHEMY_DATABASE_URI="sqlite:///" + db_path,
        SQLALCHEMY_TRACK_MODIFICATIONS=False,
        MAX_CONTENT_LENGTH=2 * 1024 * 1024,
    )
    app.json.ensure_ascii = False
    if test_config:
        app.config.update(test_config)

    db.init_app(app)
    for blueprint in (bp_main, bp_survey, bp_cases, bp_bizmap, bp_people, bp_admin, bp_monthly, bp_board, bp_auth):
        app.register_blueprint(blueprint)

    @app.route("/static/<path:filename>", endpoint="static")
    def static_file(filename):
        body = STATIC.get(filename)
        if body is None:
            abort(404)
        mimetype = mimetypes.guess_type(filename)[0] or "application/octet-stream"
        response = Response(body, content_type=f"{mimetype}; charset=utf-8")
        response.headers["Cache-Control"] = "no-cache"
        return response

    # ログイン不要で使えるもの: 活動状況、自動化依頼（棚卸し）、月次アンケートの回答、ログイン画面
    public_blueprints = {"board", "survey", "auth"}
    public_endpoints = {"static", "main.index", "monthly.form", "monthly.submit", "monthly.done", "monthly.api_employee"}

    @app.before_request
    def require_login():
        if request.endpoint is None:  # 存在しないURLはログインに回さず、そのまま 404 にする
            return None
        if request.endpoint in public_endpoints or request.blueprint in public_blueprints:
            return None
        if session.get("user"):
            return None
        nxt = request.full_path.rstrip("?") if request.method == "GET" else url_for("main.index")
        return redirect(url_for("auth.login", next=nxt))

    @app.errorhandler(404)
    def not_found(_error):
        return render_template("errors/404.html"), 404

    @app.errorhandler(500)
    def server_error(_error):
        db.session.rollback()
        try:
            return render_template("errors/500.html"), 500
        except Exception:  # noqa: BLE001  DB まで壊れていて画面が描けないとき
            return "<h1>エラーが発生しました</h1><p>しばらくしてからもう一度お試しください。</p>", 500

    @app.context_processor
    def inject_globals():
        respondent = None
        rid = session.get("respondent_id")
        if rid:
            respondent = db.session.get(Respondent, rid)
        return {
            "settings": Setting.as_dict(),
            "C": C,
            "current_user": session.get("user"),
            "current_respondent": respondent,
            "ai_enabled": ai_enabled(),
            "now": datetime.now(),
        }

    @app.template_filter("jpdate")
    def jpdate(value, fmt="%Y/%m/%d"):
        if not value:
            return ""
        return value.strftime(fmt)

    @app.template_filter("hours")
    def hours(minutes):
        """分 → 時間（小数1桁）"""
        try:
            return "{:,.1f}".format((minutes or 0) / 60.0)
        except (TypeError, ValueError):
            return "0.0"

    @app.template_filter("splitlist")
    def splitlist(value):
        return C.split_list(value)

    with app.app_context():
        db.create_all()
        ensure_schema()
        seed_masters()

    return app


# ==============================================================================
# 起動
# ==============================================================================
if __name__ == "__main__":
    port = int(conf.get("PORT", 5000) or 5000)
    if len(sys.argv) > 1 and sys.argv[1] == "serve":
        from waitress import serve

        threads = int(conf.get("THREADS", 8) or 8)
        print(f"PC定型業務ゼロ化プロジェクト: http://0.0.0.0:{port}  (waitress, threads={threads})")
        serve(create_app(), host="0.0.0.0", port=port, threads=threads)
    else:
        create_app().run(host="0.0.0.0", port=port, debug=bool(conf.get("DEBUG", False)))
