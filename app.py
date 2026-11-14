"""PC定型業務ゼロ化プロジェクト ポータル（Flask + SQLite）。アプリ本体はこの1ファイル。設定は config.py、社員ID照会は ldap_auth.py"""
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
import unicodedata
import urllib.error
import urllib.request
from collections import Counter, OrderedDict, defaultdict
from datetime import datetime, timedelta

from flask import Blueprint, Flask, Response, abort, flash, g, has_request_context, jsonify, redirect, render_template, request, session, url_for
from flask_sqlalchemy import SQLAlchemy
from jinja2 import DictLoader

import config as _config
import ldap_auth

log = logging.getLogger(__name__)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

C = sys.modules[__name__]  # 旧モジュール名の別名


# ===== 設定（config.py の読み込み） =====
class _Conf:
    def get(self, name, default=None):
        value = getattr(_config, name, default)
        return default if value is None else value

    def override(self, name, value):
        setattr(_config, name, value)  # ldap_auth も同じ config を見る

conf = _Conf()


# ===== 選択肢・係数・しきい値（アンケート設計書に準拠） =====
SEP = "|"


def split_list(value):
    if not value:
        return []
    return [v for v in str(value).split(SEP) if v]

def join_list(values):
    return SEP.join([v for v in values if v])

# ---- パートB 検討する業務（1業務ごと） ----
OUTPUT_USERS = ["自分", "職場内", "上司", "他部署", "顧客・社外", "分からない"]
INPUT_SOURCES = ["生産管理システム", "品質システム", "設備データ", "紙の帳票", "メール", "他のExcel", "人から聞く", "その他"]
FREQUENCIES = ["毎日", "週に数回", "毎週", "毎月", "四半期・年次", "不定期"]
DURATIONS = ["15分未満", "15〜30分", "30分〜1時間", "1〜2時間", "2時間以上"]
TOOLS = ["Excel", "Access", "メール", "社内システム", "紙", "VBA・マクロ", "RPA", "その他"]
NATURES = ["転記・コピー", "集計・グラフ作成", "資料作成", "確認・チェック", "データ取得", "メール・配布", "台帳更新", "判断を伴う"]
PAINS = ["時間がかかる", "ミスが起きやすい", "同じ情報を何度も入力する", "手順が人によって違う", "急ぎで割り込まれる", "何のためか分からない", "特になし"]
WISHES = ["なくしたい", "減らしたい", "まずは相談したい"]
LEGACY_WISHES = {"自動化したい": "減らしたい", "分からない": "まずは相談したい"}
SYSTEM_SOURCES = {"生産管理システム", "品質システム", "設備データ"}
SELF_IMPROVE_YES = "スキルを身につけながら、一緒にやってみたい"
SELF_IMPROVE_NO = "改善チームに任せたい"
SELF_IMPROVE_UNCHECKED = "チェックなし"
SELF_IMPROVE_LATER_ASK = "その後の月次で依頼したい"
ASSIGNMENT_REVIEW = "担当区分の見直しを"
DX_WISH_OPTIONS = [SELF_IMPROVE_YES, SELF_IMPROVE_UNCHECKED]
INTAKE_MAX_QUESTIONS = 6
INTAKE_STEP_LIMIT = 12
WORK_PATTERNS = ["普通勤務", "交代勤務"]
SHIFT_OPTIONS = ["昼勤", "夜勤", "昼勤・夜勤の両方"]
FREQUENCY_HINTS = {"毎日": "毎日（出勤のたび）", "不定期": "不定期（月1回として計算）"}

# ---- 使える技術・学びたい技術（「一緒にやってみたい」にチェックした人に出すチップ） ----
SKILLS = ["Excel関数", "ピボットテーブル", "Excel VBA", "Access", "Power Query", "Power BI", "Python", "RPA", "生成AI", "特になし"]
SKILL_NONE = "特になし"

# ---- 月間時間の算出（設計書 10. 集計・活用の流れ） ----
FREQ_COEF = {
    "毎日": 20.0,
    "週に数回": 10.0,
    "毎週": 4.0,
    "毎月": 1.0,
    "四半期・年次": 0.3,
    "不定期": 1.0,
}
DURATION_MINUTES = {
    "15分未満": 10.0,
    "15〜30分": 22.0,
    "30分〜1時間": 45.0,
    "1〜2時間": 90.0,
    "2時間以上": 150.0,
}

# ---- 案件管理 ----
STATUSES = ["受付", "対応中", "完了", "保留"]
DECISIONS = ["未定", "廃止", "統合", "簡素化", "自動化", "高度化", "現状維持"]
ASSIGNMENTS = ["未定", "改善チーム主導", "自主改善"]
REQUIRED_FIELDS = [("domain_id", "業務領域"), ("purpose", "目的"), ("output", "アウトプット"), ("output_users", "利用者"), ("inputs", "入力元"), ("frequency", "頻度"), ("duration", "時間")]
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
STAGES = [
    ("start", "開始", "検討の登録を受け付け、活動状況に載せる", ""),
    ("request", "要求定義", "何に困っていて、何を実現したいかを本人と整理する", "手順1・2（アウトプットの確認・目的の明確化）"),
    ("spec", "要件定義", "ECRSで業務を最適化し、残る作業の仕様を決める", "手順3（ECRSによる最適化）"),
    ("build", "作成", "ツール・仕組みを作る。自主改善の案件は改善チームが伴走する", "手順4・5（自動化・高度化）"),
    ("test", "テスト", "実際のデータで動作と効果を確かめる", ""),
    ("check", "確認", "本人・職場で結果を確認し、運用に乗せる", ""),
    ("done", "完了！", "運用開始。改善前後の結果を記録する", ""),
]
STAGE_KEYS = [s[0] for s in STAGES]
STAGE_INDEX = {k: i for i, k in enumerate(STAGE_KEYS)}
STAGE_NAMES = {s[0]: s[1] for s in STAGES}
STAGE_STEPS = {"request": (1, 2), "spec": (3,)}
BUILD_STAGE_SELF = "伴走"

# ---- 案件候補・改善意欲の判定 ----
CANDIDATE_WISHES = {"なくしたい", "減らしたい", "自動化したい"}
SELF_IMPROVE_POSITIVE = {SELF_IMPROVE_YES, "取り組みたい", "教えてもらえれば取り組みたい"}
STANCE_MONTHS = 6

# ---- 分析不足領域の判定しきい値 ----
ANALYSIS_LEVELS = [
    (3, "分析済み", "十分", 8, 0.8),
    (2, "追加調査", "一部", 3, 0.5),
    (1, "情報不足", "不十分", 1, 0.0),
    (0, "未分析", "不明", 0, 0.0),
]

# ---- AI分析のフラグ ----
AI_FLAGS = ["目的不明", "廃止候補", "統合候補", "重複入力候補", "自動化候補", "判断業務", "要ヒアリング"]
AI_PLAN_KEYS = ("tobe_steps", "saving_minutes_per_month", "hearing_questions")

# ---- 月次アンケート（voiceloop 方式: 月度ごとの設問セット。基本設問は毎月同じで KPI につなぐ） ----
SHIFTS_PER_MONTH = 20
SHIFTS_PER_MONTH_BY_PATTERN = {"普通勤務": 20, "交代勤務": 20}


def shifts_per_month(pattern=None):
    return SHIFTS_PER_MONTH_BY_PATTERN.get(pattern or "", SHIFTS_PER_MONTH)

MONTHLY_HOURS_OPTIONS = {
    "15分未満": 0.125, "15〜30分": 0.375, "30〜45分": 0.625, "45分〜1時間": 0.875, "1〜1.5時間": 1.25,
    "1.5〜2時間": 1.75, "2〜3時間": 2.5, "3〜4時間": 3.5, "4〜6時間": 5.0, "6時間以上": 7.0,
}
MONTHLY_HOURS_LEGACY = {"0時間": 0.0, "30分未満": 0.25, "30分〜1時間": 0.75, "1〜2時間": 1.5, "2〜4時間": 3.0, "4時間以上": 5.0}
MONTHLY_HOURS_VALUE = {**MONTHLY_HOURS_OPTIONS, **MONTHLY_HOURS_LEGACY}
MONTHLY_TREND_GROUPS = {
    "hours": {
        "30分未満": ["0時間", "30分未満", "15分未満", "15〜30分"],
        "30分〜1時間": ["30分〜1時間", "30〜45分", "45分〜1時間"],
        "1〜2時間": ["1〜2時間", "1〜1.5時間", "1.5〜2時間"],
        "2〜4時間": ["2〜4時間", "2〜3時間", "3〜4時間"],
        "4時間以上": ["4時間以上", "4〜6時間", "6時間以上"],
    },
}
_IF_HAS = {"key": "has", "show_if": ["ある"], "required": True}
STANCE_SELF = "自分で取り組みたい"
STANCE_TOGETHER = "学びながら改善チームと一緒に取り組みたい"
STANCE_ASK = "改善チームに依頼したい"
STANCE_OPTIONS = [STANCE_SELF, STANCE_TOGETHER, STANCE_ASK]
MONTHLY_CORE_QUESTIONS = [
    {"key": "has", "type": "choice", "text": "PCの定型・定期業務（繰り返しの手動作業）はありますか", "options": ["ある", "ない"], "required": True},
    {"key": "hours", "type": "choice", "text": "1シフト（1出勤）あたり、PCの定型・定期業務を行っている合計時間は", "options": list(MONTHLY_HOURS_OPTIONS), "required": False, "depends": dict(_IF_HAS)},
    {"key": "task_names", "type": "list", "text": "行っているPCの定型・定期業務名", "required": False, "placeholder": "例：日報の集計", "depends": dict(_IF_HAS)},
    {"key": "trend", "type": "choice", "text": "先月と比べて、PCの定型・定期業務は", "options": ["減った", "変わらない", "増えた"], "required": False, "depends": dict(_IF_HAS, main=True)},
    {"key": "decreased_reason", "type": "text", "text": "なぜ減ったのか", "required": False, "placeholder": "例：日報の転記を自動化した、やり方を見直した、担当が変わった",
     "depends": {"key": "trend", "show_if": ["減った"], "required": True}},
    {"key": "new_tasks", "type": "text", "text": "増えたPCの定型・定期業務名", "required": False, "placeholder": "例：新しい報告フォーマットへの転記",
     "depends": {"key": "trend", "show_if": ["増えた"], "required": True}},
    {"key": "wants", "type": "choice", "text": "なくしたい・自動化したいPCの定型・定期業務はありますか", "options": ["ある", "ない"], "required": False, "depends": dict(_IF_HAS, main=True)},
    {"key": "wants_tasks", "type": "list", "text": "あれば、その業務名", "required": False, "placeholder": "例：日報の実績をExcelに転記する",
     "depends": {"key": "wants", "show_if": ["ある"], "required": True}},
    {"key": "stance", "type": "choice", "text": "どのように進めたいですか", "options": STANCE_OPTIONS, "required": False,
     "depends": {"key": "wants", "show_if": ["ある"], "required": True}},
]
MONTHLY_CORE_KEYS = {q["key"] for q in MONTHLY_CORE_QUESTIONS}
MONTHLY_COPY_FROM = {"wants_tasks": "task_names"}
MONTHLY_KPI_LABELS = {
    "has": "PCの定型・定期業務あり率",
    "hours": "推計 月間時間",
    "trend": "減った率・増えた率",
    "wants": "自動化したい率",
    "wants_tasks": "挙げられた業務（自動化を検討へ誘導）",
    "stance": "依頼したい率・一緒に取り組みたい率・自分で取り組みたい率",
    "new_tasks": "新規業務の件数",
    "task_names": "挙げられた定型・定期業務",
    "decreased_reason": "減った理由",
}
MONTHLY_KPI = {
    "has_positive": {"ある"},
    "trend_decreased": {"減った"},
    "trend_increased": {"増えた"},
    "wants_positive": {"ある"},
    "ask_positive": {STANCE_ASK},
    "self_positive": {STANCE_SELF, STANCE_TOGETHER},
    "together_positive": {STANCE_TOGETHER},
    "self_only": {STANCE_SELF},
}


# ===== 自動化を検討（1画面）のフォーム定義 =====
PERSON_FIELDS = [
    {"key": "emp_id", "type": "text", "label": "社員ID", "required": True, "lookup": True, "placeholder": "例：E001"},
    {"key": "workplace_id", "type": "select", "label": "職場", "required": True, "options": "workplaces"},
    {"key": "work_pattern", "type": "select", "label": "勤務形態", "required": True, "options": "work_patterns"},
]

# ===== データモデル（SQLite） =====
db = SQLAlchemy()


def _json_column(text, default):
    try:
        return json.loads(text) if text else default
    except ValueError:
        return default

def _json_typed(text, default):
    data = _json_column(text, default)
    return data if isinstance(data, type(default)) else default


class Workplace(db.Model):
    __tablename__ = "workplaces"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), unique=True, nullable=False)
    sort = db.Column(db.Integer, default=0)

    respondents = db.relationship("Respondent", back_populates="workplace")

class Domain(db.Model):
    __tablename__ = "domains"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    description = db.Column(db.String(300), default="")
    sort = db.Column(db.Integer, default=0)
    parent_id = db.Column(db.Integer, db.ForeignKey("domains.id"), nullable=True)

    parent = db.relationship("Domain", remote_side=[id], backref="children")
    tasks = db.relationship("Task", back_populates="domain")

class Respondent(db.Model):
    __tablename__ = "respondents"
    id = db.Column(db.Integer, primary_key=True)
    emp_id = db.Column(db.String(50))
    name = db.Column(db.String(100), nullable=False)
    workplace_id = db.Column(db.Integer, db.ForeignKey("workplaces.id"))
    created_at = db.Column(db.DateTime, default=datetime.now)
    is_demo = db.Column(db.Boolean, default=False)
    work_pattern = db.Column(db.String(20), default="")

    c1_wish = db.Column(db.String(50))
    c2_skills = db.Column(db.Text, default="")
    c3_learn = db.Column(db.Text, default="")
    c4_hours = db.Column(db.String(50))
    c5_targets = db.Column(db.Text, default="")
    c6_share = db.Column(db.String(50))
    part_c_done = db.Column(db.Boolean, default=False)

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
        return any(t.self_improve in C.SELF_IMPROVE_POSITIVE for t in self.tasks)

class Task(db.Model):
    __tablename__ = "tasks"
    id = db.Column(db.Integer, primary_key=True)
    respondent_id = db.Column(db.Integer, db.ForeignKey("respondents.id"), nullable=False)
    workplace_id = db.Column(db.Integer, db.ForeignKey("workplaces.id"))
    is_demo = db.Column(db.Boolean, default=False)

    name = db.Column(db.String(200), nullable=False)
    domain_id = db.Column(db.Integer, db.ForeignKey("domains.id"))
    purpose = db.Column(db.Text, default="")
    output = db.Column(db.Text, default="")
    output_users = db.Column(db.Text, default="")
    output_users_note = db.Column(db.Text, default="")
    inputs = db.Column(db.Text, default="")
    inputs_note = db.Column(db.Text, default="")
    prev_task = db.Column(db.Text, default="")
    next_task = db.Column(db.Text, default="")
    frequency = db.Column(db.String(30))
    duration = db.Column(db.String(30))
    headcount = db.Column(db.Integer, default=1)
    tools = db.Column(db.Text, default="")
    nature = db.Column(db.Text, default="")
    pains = db.Column(db.Text, default="")
    pains_note = db.Column(db.Text, default="")
    wish = db.Column(db.String(30))
    self_improve = db.Column(db.String(50))
    note = db.Column(db.Text, default="")

    actors = db.Column(db.Text, default="")
    info_from = db.Column(db.Text, default="")
    steps_json = db.Column(db.Text, default="[]")
    intake_json = db.Column(db.Text, default="[]")
    ai_fields = db.Column(db.Text, default="")
    ai_extra = db.Column(db.Text, default="{}")
    shift = db.Column(db.String(20), default="")
    monthly_name = db.Column(db.String(200), default="")

    monthly_minutes = db.Column(db.Float, default=0.0)

    ai_domain = db.Column(db.String(100))
    ai_related = db.Column(db.Text)
    ai_summary = db.Column(db.Text)
    ai_ecrs = db.Column(db.Text)
    ai_flags = db.Column(db.Text)
    ai_priority = db.Column(db.String(10))
    ai_source = db.Column(db.String(20))
    ai_at = db.Column(db.DateTime)

    status = db.Column(db.String(20), default="受付")
    steps_done = db.Column(db.String(50), default="")
    decision = db.Column(db.String(20), default="未定")
    assignment = db.Column(db.String(20), default="未定")
    assignee = db.Column(db.String(100), default="")
    ecrs_eliminate = db.Column(db.Text, default="")
    ecrs_combine = db.Column(db.Text, default="")
    ecrs_rearrange = db.Column(db.Text, default="")
    ecrs_simplify = db.Column(db.Text, default="")
    after_minutes = db.Column(db.Float, nullable=True)
    result_note = db.Column(db.Text, default="")
    completed_at = db.Column(db.DateTime, nullable=True)
    verified = db.Column(db.Boolean, default=False)
    verified_by = db.Column(db.String(100), default="")

    stage = db.Column(db.String(20), default="start")
    stage_at = db.Column(db.DateTime, nullable=True)
    progress_note = db.Column(db.Text, default="")

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
        return [a.strip() for a in re.split(r"[、,，/／\n;；]+", self.actors or "") if a.strip()]

    @property
    def has_flow_info(self):
        return bool(self.actors or self.info_from)

    @property
    def missing_required(self):
        return [label for key, label in C.REQUIRED_FIELDS if not str(getattr(self, key) or "").strip()]

    @property
    def required_filled(self):
        return not self.missing_required

    @property
    def is_sufficient(self):
        return self.required_filled and self.verified

    @property
    def purpose_unknown(self):
        p = (self.purpose or "").strip()
        return (not p) or ("不明" in p) or ("分からない" in p) or ("わからない" in p)

    @property
    def is_candidate(self):
        return self.wish in C.CANDIDATE_WISHES or any(p != "特になし" for p in self.pains_list)

    @property
    def is_elimination_candidate(self):
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
    def steps_lagging(self):
        done = set(self.steps_done_list)
        return [n for key in C.STAGE_KEYS[:self.stage_index] for n in C.STAGE_STEPS.get(key, ()) if n not in done]

    @property
    def stage_dates(self):
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
        if key not in C.STAGE_INDEX:
            return False
        changed = key != self.stage_key
        if changed:
            now = datetime.now()
            self.stage = key
            self.stage_at = now
            self.stage_logs.append(StageLog(stage=key, at=now, by=by or "", note=note or ""))
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
        return _json_column(self.ai_related, [])

    @property
    def ai_ecrs_dict(self):
        return _json_column(self.ai_ecrs, {})

    @property
    def ai_flags_list(self):
        return _json_column(self.ai_flags, [])

    @property
    def steps(self):
        return [x for x in _json_typed(self.steps_json, []) if isinstance(x, dict) and x.get("action")]

    @property
    def intake(self):
        return [m for m in _json_typed(self.intake_json, []) if isinstance(m, dict)]

    @property
    def ai_fields_list(self):
        return C.split_list(self.ai_fields)

    @property
    def ai_extra_dict(self):
        return _json_typed(self.ai_extra, {})

    @property
    def ai_plan_note(self):
        extra = self.ai_extra_dict
        if not any(extra.get(k) for k in C.AI_PLAN_KEYS):
            return ""
        if extra.get("ai_plan_demo"):
            return "デモ"
        if self.ai_source not in ("openai", "claude"):
            return "前回の AI 分析（" + (extra.get("ai_plan_at") or "日付不明") + "）"
        return ""

    @property
    def ai_stale_note(self):
        at = self.ai_extra_dict.get("ai_before_edit")
        return "業務情報を直す前の AI 分析（" + at + "）" if at and self.ai_source in ("openai", "claude") else ""

    @property
    def domain_guess(self):
        if not self.ai_domain or (self.domain is not None and self.ai_domain == self.domain.name):
            return ""
        return self.ai_domain if self.domain is None or ("domain" in self.ai_fields_list and not self.verified) else ""

    def apply_ai_result(self, result):
        mask = lambda v: [mask_person_names(x) for x in v] if isinstance(v, list) else (mask_person_names(v) if isinstance(v, str) else v)
        self.ai_domain = result.get("domain")
        self.ai_related = json.dumps(result.get("related") or [], ensure_ascii=False)
        self.ai_summary = mask_person_names(result.get("summary") or "")
        self.ai_ecrs = json.dumps({k: mask(v) for k, v in (result.get("ecrs") or {}).items()}, ensure_ascii=False)
        self.ai_flags = json.dumps(result.get("flags") or [], ensure_ascii=False)
        self.ai_priority = result.get("priority") or ""
        self.ai_source = result.get("source") or ""
        self.ai_at = datetime.now()

class StageLog(db.Model):
    __tablename__ = "stage_logs"
    id = db.Column(db.Integer, primary_key=True)
    task_id = db.Column(db.Integer, db.ForeignKey("tasks.id"), nullable=False, index=True)
    stage = db.Column(db.String(20), nullable=False)
    at = db.Column(db.DateTime, default=datetime.now)
    by = db.Column(db.String(100), default="")
    note = db.Column(db.Text, default="")

    task = db.relationship("Task", back_populates="stage_logs")

class Setting(db.Model):
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
    __tablename__ = "employees"
    id = db.Column(db.Integer, primary_key=True)
    emp_id = db.Column(db.String(50), unique=True, nullable=False)
    name = db.Column(db.String(100), nullable=False)
    department = db.Column(db.String(100), default="")
    work_pattern = db.Column(db.String(20), default="")
    workplace_id = db.Column(db.Integer, db.ForeignKey("workplaces.id"), nullable=True)
    email = db.Column(db.String(200), default="")
    active = db.Column(db.Boolean, default=True)
    excluded_at = db.Column(db.DateTime, nullable=True)
    source = db.Column(db.String(20), default="csv")
    is_demo = db.Column(db.Boolean, default=False)
    updated_at = db.Column(db.DateTime, default=datetime.now, onupdate=datetime.now)
    created_at = db.Column(db.DateTime, default=datetime.now)

    workplace = db.relationship("Workplace")

class MonthlySurvey(db.Model):
    __tablename__ = "monthly_surveys"
    id = db.Column(db.Integer, primary_key=True)
    period = db.Column(db.String(7), unique=True, nullable=False)
    questions_json = db.Column(db.Text, default="[]")
    is_active = db.Column(db.Boolean, default=True)
    created_at = db.Column(db.DateTime, default=datetime.now)
    updated_at = db.Column(db.DateTime, default=datetime.now, onupdate=datetime.now)

    responses = db.relationship("MonthlyResponse", back_populates="survey")

    @property
    def questions(self):
        return _json_typed(self.questions_json, [])

    def set_questions(self, questions):
        self.questions_json = json.dumps(questions, ensure_ascii=False)

    @property
    def period_label(self):
        return period_label(self.period)

class MonthlyResponse(db.Model):
    __tablename__ = "monthly_responses"
    __table_args__ = (db.UniqueConstraint("period", "emp_id", name="uq_monthly_period_emp"),)
    id = db.Column(db.Integer, primary_key=True)
    period = db.Column(db.String(7), nullable=False)
    emp_id = db.Column(db.String(50), nullable=False)
    work_pattern = db.Column(db.String(20), default="")
    name = db.Column(db.String(100), default="")
    workplace_id = db.Column(db.Integer, db.ForeignKey("workplaces.id"), nullable=True)
    respondent_id = db.Column(db.Integer, db.ForeignKey("respondents.id"), nullable=True)
    survey_id = db.Column(db.Integer, db.ForeignKey("monthly_surveys.id"), nullable=True)
    answers_json = db.Column(db.Text, default="{}")
    comment = db.Column(db.Text, default="")
    created_at = db.Column(db.DateTime, default=datetime.now)
    updated_at = db.Column(db.DateTime, default=datetime.now, onupdate=datetime.now)
    is_demo = db.Column(db.Boolean, default=False)

    workplace = db.relationship("Workplace")
    survey = db.relationship("MonthlySurvey", back_populates="responses")

    @property
    def answers(self):
        return _json_typed(self.answers_json, {})

    def answer(self, key):
        return self.answers.get(key) or ""

    @property
    def hours_shift(self):
        return C.MONTHLY_HOURS_VALUE.get(self.answer("hours"), 0.0)

def period_label(period):
    try:
        y, m = period.split("-")
        return f"{int(y)}年{int(m)}月度"
    except (ValueError, AttributeError):
        return period or ""

class FlowMap(db.Model):
    __tablename__ = "flow_maps"
    __table_args__ = (db.UniqueConstraint("kind", "scope", name="uq_flow_map_kind_scope"),)
    id = db.Column(db.Integer, primary_key=True)
    kind = db.Column(db.String(20), nullable=False)
    scope = db.Column(db.String(60), nullable=False)
    data = db.Column(db.Text, default="{}")
    updated_by = db.Column(db.String(100), default="")
    updated_at = db.Column(db.DateTime, default=datetime.now, onupdate=datetime.now)

_ADDED_COLUMNS = {
    "respondents": {
        "emp_id": "VARCHAR(50)",
        "work_pattern": "VARCHAR(20) DEFAULT ''",
    },
    "employees": {
        "work_pattern": "VARCHAR(20) DEFAULT ''",
        "created_at": "DATETIME",
        "excluded_at": "DATETIME",
    },
    "monthly_responses": {
        "survey_id": "INTEGER",
        "answers_json": "TEXT DEFAULT '{}'",
        "comment": "TEXT DEFAULT ''",
        "work_pattern": "VARCHAR(20) DEFAULT ''",
    },
    "tasks": {
        "actors": "TEXT DEFAULT ''",
        "info_from": "TEXT DEFAULT ''",
        "stage": "VARCHAR(20)",
        "stage_at": "DATETIME",
        "progress_note": "TEXT DEFAULT ''",
        "steps_json": "TEXT DEFAULT '[]'",
        "intake_json": "TEXT DEFAULT '[]'",
        "ai_fields": "TEXT DEFAULT ''",
        "ai_extra": "TEXT DEFAULT '{}'",
        "shift": "VARCHAR(20) DEFAULT ''",
        "monthly_name": "VARCHAR(200) DEFAULT ''",
    },
}

_BACKFILL = {
    ("tasks", "stage"): [
        "UPDATE tasks SET stage = CASE status WHEN '完了' THEN 'done' WHEN '対応中' THEN 'request' ELSE 'start' END WHERE stage IS NULL",
        "UPDATE tasks SET stage_at = COALESCE(completed_at, updated_at, created_at) WHERE stage_at IS NULL",
        "INSERT INTO stage_logs (task_id, stage, at, by, note) SELECT id, stage, stage_at, '', '' FROM tasks WHERE stage <> 'start'",
    ],
}


_DROPPED_TABLES = ("resources", "allowed_users", "mail_logs")


def ensure_schema():
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

# ----- 番号の払い出し（業務・職場・領域の番号は使い回さない。消した行の手修正が新しい行に付かないように） -----
ID_SEQ_TABLES = (("tasks", Task), ("workplaces", Workplace), ("domains", Domain))


def _id_seq_key(table):
    return f"id_seq:{table}"

def _assign_id(mapper, connection, target):
    if target.id is not None:
        return
    from sqlalchemy import text

    table, key = mapper.local_table.name, _id_seq_key(mapper.local_table.name)
    connection.execute(text("INSERT OR IGNORE INTO settings (key, value) VALUES (:k, '0')"), {"k": key})
    connection.execute(text(f"UPDATE settings SET value = CAST(MAX(CAST(value AS INTEGER), (SELECT COALESCE(MAX(id), 0) FROM {table})) + 1 AS TEXT) WHERE key = :k"), {"k": key})
    target.id = int(connection.execute(text("SELECT value FROM settings WHERE key = :k"), {"k": key}).scalar())

for _seq_table, _seq_model in ID_SEQ_TABLES:
    db.event.listen(_seq_model, "before_insert", _assign_id)

def _referenced_ids():
    top = {table: 0 for table, _ in ID_SEQ_TABLES}
    for scope, data in db.session.query(FlowMap.scope, FlowMap.data).all():
        m = re.match(r"wp:(\d+)\|dom:(\d+)", scope or "")
        if m:
            top["workplaces"], top["domains"] = max(top["workplaces"], int(m.group(1))), max(top["domains"], int(m.group(2)))
        graph = _json_typed(data, {})
        refs = [str(x) for x in (graph.get("suppressed_nodes") or [])]
        refs += [part for x in (graph.get("suppressed_edges") or []) for part in str(x).split("->")]
        for n in graph.get("nodes") or []:
            if isinstance(n, dict):
                refs.append(str(n.get("id") or ""))
                wp_ref = re.fullmatch(r"wp(\d+)", str(n.get("group") or ""))
                if wp_ref:
                    top["workplaces"] = max(top["workplaces"], int(wp_ref.group(1)))
        for e in graph.get("edges") or []:
            if isinstance(e, dict):
                refs += [str(e.get("from") or ""), str(e.get("to") or "")]
        for ref in refs:
            if re.fullmatch(r"[to]\d+", ref):
                top["tasks"] = max(top["tasks"], int(ref[1:]))
    for ids in _map_written().values():
        for tid in ids if isinstance(ids, dict) else {}:
            if str(tid).isdigit():
                top["tasks"] = max(top["tasks"], int(tid))
    return top

def init_id_seq():
    refs = _referenced_ids()
    for table, model in ID_SEQ_TABLES:
        stored = Setting.get(_id_seq_key(table), "0")
        top = max(int(stored) if str(stored).isdigit() else 0, db.session.query(db.func.max(model.id)).scalar() or 0, refs[table])
        if str(top) != stored:
            Setting.set(_id_seq_key(table), str(top))
    db.session.commit()

# ===== 社員ID照会の控え（ldap_auth の結果を employees に残し、足りない氏名・職場を補う） =====
ldap_status_text = ldap_auth.status_text
EMP_NOT_FOUND = "この社員IDは見つかりません。IDを確かめてください"
EMP_EXCLUDED = "この社員IDは対象外になっています。管理者に連絡してください"


def workplace_for(department):
    department = (department or "").strip()
    if not department:
        return None
    for w in Workplace.query.all():
        if w.name == department or w.name in department or department in w.name:
            return w
    return None

def norm_emp_id(value):
    return unicodedata.normalize("NFKC", str(value or "")).strip()

def find_employee(emp_id):
    if not norm_emp_id(emp_id):
        return None
    return Employee.query.filter(same_emp_id(Employee.emp_id, emp_id)).first()

def lookup(emp_id):
    emp_id = norm_emp_id(emp_id)
    if not emp_id:
        return None
    info = ldap_auth.lookup_directory(emp_id)
    emp = (find_employee(info["emp_id"]) if info else None) or find_employee(emp_id)
    if info is None and emp is None:
        return None
    info = info or {}
    name = str(info.get("name") or "").strip()
    if emp is not None and name and emp.name != name:
        emp.name = name
        db.session.commit()
    return {
        "emp_id": emp.emp_id if emp else (info.get("emp_id") or emp_id),
        "name": name or (emp.name if emp else "") or "",
        "active": emp is None or emp.active is not False,
    }

def emp_key(emp_id):
    return norm_emp_id(emp_id).lower()

def same_emp_id(column, emp_id):
    return db.func.lower(column) == emp_key(emp_id)

def remember_person(emp_id, name=None, workplace_id=None, work_pattern=None):
    emp = find_employee(emp_id)
    if emp is None:
        return False
    if name and emp.name != name:
        emp.name = name
    if workplace_id:
        emp.workplace_id = workplace_id
    if work_pattern in C.WORK_PATTERNS:
        emp.work_pattern = work_pattern
    return True

def register_member(emp_id, name="", workplace_id=None, work_pattern=""):
    if not norm_emp_id(emp_id):
        return False
    if find_employee(emp_id) is None:
        db.session.add(Employee(emp_id=norm_emp_id(emp_id), name=name or "", source="answer"))
        db.session.flush()
    return remember_person(emp_id, name, workplace_id, work_pattern)

def newest_period(emp_id):
    if not emp_key(emp_id):
        return ""
    monthly = db.session.query(db.func.max(MonthlyResponse.period)).filter(same_emp_id(MonthlyResponse.emp_id, emp_id)).scalar()
    task_at = (db.session.query(db.func.max(Task.created_at)).join(Respondent, Task.respondent_id == Respondent.id)
               .filter(same_emp_id(Respondent.emp_id, emp_id)).scalar())
    return max(monthly or "", _period_of(task_at) or "")

def current_place(emp_id, fallback=None):
    r = Respondent.query.filter(same_emp_id(Respondent.emp_id, emp_id)).order_by(Respondent.id).first()
    for x in (r, find_employee(emp_id), fallback):
        if x is not None and x.workplace_id:
            return x.workplace_id, x.work_pattern or ""
    return None, ""

# ===== 集計・KPI・分析不足領域の判定・次アクション提案 =====
# ----- KPI・ダッシュボードのグラフ -----
def kpis():
    tasks = Task.query.all()
    people = dx_people()
    m_periods = member_periods()[-1:]
    latest = member_stats(member_rows(), m_periods)[m_periods[0]] if m_periods else None

    total_minutes = sum(t.monthly_minutes or 0 for t in tasks)
    saved_minutes = sum(t.saved_minutes for t in tasks)
    completed = [t for t in tasks if t.status == "完了"]
    status_counts = Counter(t.status for t in tasks)
    sufficient = [t for t in tasks if t.is_sufficient]
    filled = [t for t in tasks if t.required_filled]
    verified = [t for t in tasks if t.verified]
    self_improvers = [p for p in people if p["positive"]]
    automated_or_eliminated = [t for t in completed if t.decision in ("廃止", "自動化")]
    candidates = [t for t in tasks if t.is_candidate]
    domain_rows = domain_analysis()

    return {
        "task_count": len(tasks),
        "respondent_count": len(people),
        "response_rate": latest["rate"] if latest else None,
        "response_period": period_label(m_periods[0]) if m_periods else "",
        "response_answered": latest["answered"] if latest else 0,
        "response_members": latest["members"] if latest else 0,
        "total_hours": total_minutes / 60.0,
        "annual_hours": total_minutes * 12 / 60.0,
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
    }

def dashboard_chart_data():
    wp = workplace_stats()
    tasks = Task.query.all()
    wish_counter = Counter(t.wish or "未回答" for t in tasks)
    pain_counter = Counter(p for t in tasks for p in t.pains_list if p != "特になし")
    wish_labels = [w for w in C.WISHES + ["未回答"] if wish_counter.get(w)]
    pain_counts = pain_counter.most_common()
    return {
        "workplace": {
            "labels": [w["workplace"].name for w in wp],
            "tasks": [w["tasks"] for w in wp],
            "hours": [round(w["minutes"] / 60.0, 1) for w in wp],
        },
        "wish": {
            "labels": wish_labels,
            "values": [wish_counter[w] for w in wish_labels],
        },
        "pains": {
            "labels": [p for p, _ in pain_counts],
            "values": [n for _, n in pain_counts],
        },
    }

# ----- 職場別・領域別 -----
def workplace_stats():
    rows = []
    people = dx_people()
    for wp in Workplace.query.order_by(Workplace.sort, Workplace.id).all():
        tasks = Task.query.filter_by(workplace_id=wp.id).all()
        here = [p for p in people if p["workplace_id"] == wp.id]
        rows.append(
            {
                "workplace": wp,
                "respondents": len(here),
                "ever_answered": any(wp.id in p["workplace_ids"] for p in people),
                "tasks": len(tasks),
                "minutes": sum(t.monthly_minutes or 0 for t in tasks),
                "candidates": len([t for t in tasks if t.is_candidate]),
                "self_improvers": len([p for p in here if p["positive"]]),
            }
        )
    return rows

def judge_level(count, fill_rate):
    for level, state, grasp, min_count, min_fill in C.ANALYSIS_LEVELS:
        if count >= min_count and fill_rate >= min_fill:
            return level, state, grasp
    return 0, "未分析", "不明"

def domain_analysis():
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

# ----- 案件候補・重複候補・つながり -----
def candidate_tasks(limit=None):
    tasks = [t for t in Task.query.all() if t.is_candidate and t.status != "完了"]
    tasks.sort(key=lambda t: t.monthly_minutes or 0, reverse=True)
    return tasks[:limit] if limit else tasks

_NORMALIZE_RE = re.compile(r"[\s　（）()・,、。.／/\-−―「」『』]")


def normalize_text(value):
    value = (value or "").strip().lower()
    return _NORMALIZE_RE.sub("", value)

def duplicate_output_groups():
    groups = defaultdict(list)
    for t in Task.query.all():
        key = normalize_text(t.output)
        if len(key) >= 2:
            groups[key].append(t)
    result = [
        {"output": tasks[0].output, "tasks": tasks, "minutes": sum(x.monthly_minutes or 0 for x in tasks)}
        for tasks in groups.values()
        if len({x.respondent_id for x in tasks}) >= 2
    ]
    result.sort(key=lambda g: g["minutes"], reverse=True)
    return result

def duplicate_input_tasks():
    rows = defaultdict(list)
    for t in Task.query.all():
        if "同じ情報を何度も入力する" in t.pains_list:
            key = "・".join(t.inputs_list) or "入力元不明"
            rows[key].append(t)
    return sorted(rows.items(), key=lambda kv: -len(kv[1]))

def flow_rows():
    tasks = Task.query.order_by(Task.domain_id, Task.name).all()
    return [t for t in tasks if (t.prev_task or "").strip() or (t.next_task or "").strip()]

RELATED_MIN_OVERLAP = 3


def _related_text(t):
    return t.name + " " + (t.output or "") + " " + (t.purpose or "")

def _link_key(text):
    return normalize_text(text).replace("の", "")

def _links_to(a, b):
    name = _link_key(b.name)
    keys = [_link_key(p) for text in (a.next_task, a.prev_task) for p in task_names(text)]
    return len(name) >= 3 and any(len(k) >= 3 and (k in name or name in k) for k in keys)

def related_tasks_for(task, limit=8):
    words = _word_kinds(_related_text(task))
    output = normalize_text(task.output)
    scored = []
    for other in Task.query.filter(Task.id != task.id).all():
        other_words = _word_kinds(_related_text(other))
        shared = [_pair_weight(w, words, other_words) for w in words.keys() & other_words.keys()]
        link = 0
        if output and normalize_text(other.output) == output:
            link += 5
        if _links_to(task, other) or _links_to(other, task):
            link += 4
        if not link and (sum(shared) < C.RELATED_MIN_OVERLAP or max(shared, default=0) < 2):
            continue
        score = sum(shared) + link
        if other.domain_id and other.domain_id == task.domain_id:
            score += 1
        if set(task.inputs_list) & set(other.inputs_list):
            score += 1
        scored.append((score, other))
    scored.sort(key=lambda x: -x[0])
    return [o for _, o in scored[:limit]]

_WORD_RE = re.compile(r"[一-龥々〆]{2,}|[ァ-ヶー]{2,}|[A-Za-z0-9]{2,}")
_STOPWORDS = {"作成", "業務", "確認", "管理", "対応", "実施", "内容", "情報", "毎日", "毎週", "毎月", "前日", "翌日"}
# 関連業務の判定で重みを1にする語（ツール・媒体・どの業務にも出る動作や相手）
_WEAK_WORDS = {"excel", "エクセル", "csv", "pdf", "word", "ワード", "powerpoint", "パワポ", "access", "vba", "マクロ", "rpa", "メール", "システム", "データ", "ファイル", "シート", "フォルダ",
               "グラフ", "チェック", "報告", "状況", "共有", "提出", "連絡", "把握", "集計", "転記", "入力", "記録", "更新", "一覧", "資料", "会議",
               "集計表", "一覧表", "報告書", "管理表", "日次", "週次", "月次", "年次", "課長", "部長", "上司", "職場", "各職場", "各部署", "関係部署", "担当者", "ライン"}


def _word_kinds(text):
    kinds = {}
    for w in _WORD_RE.findall(unicodedata.normalize("NFKC", text or "")):
        if w in _STOPWORDS or re.fullmatch(r"ー+", w):
            continue
        kinds[w.lower()] = "whole"
        if len(w) >= 4 and re.fullmatch(r"[一-龥々〆]+", w):
            for i in range(0, len(w) - 1):
                if w[i : i + 2] not in _STOPWORDS:
                    kinds.setdefault(w[i : i + 2], "part")
    return kinds

def _content_words(text):
    return set(_word_kinds(text))

def _pair_weight(w, a, b):
    if w in _WEAK_WORDS or (a[w] == "part" and b[w] == "part"):
        return 1
    return 3 if len(w) >= 3 else 2

# ----- 回答した人（業務を1件以上登録したか、月次アンケートに回答した人。対象外は除く。対象者一覧と同じ決まり） -----
def person_key(emp_id, name=""):
    key = emp_key(emp_id)
    return ("id", key) if key else ("name", (name or "").strip())

def answered_people():
    excluded = {emp_key(e) for (e,) in db.session.query(Employee.emp_id).filter_by(active=False).all()}
    people = OrderedDict()

    def person(emp_id, name):
        k = person_key(emp_id, name)
        if k[0] == "id" and k[1] in excluded:
            return None
        if k not in people:
            people[k] = {"key": k, "emp_id": norm_emp_id(emp_id), "name": name or "", "respondent": None, "tasks": [], "monthly": [], "workplace_ids": set(),
                         "workplace_id": None, "placed_by": None}
        return people[k]

    def place(p, workplace_id, period, at):
        if not workplace_id:
            return
        p["workplace_ids"].add(workplace_id)
        key = (period or "", at or datetime.min)
        if p["placed_by"] is None or key >= p["placed_by"]:
            p["workplace_id"], p["placed_by"] = workplace_id, key

    for t in Task.query.order_by(Task.created_at, Task.id).all():
        p = person(t.respondent.emp_id, t.respondent.name) if t.respondent is not None else None
        if p is None:
            continue
        p["respondent"] = p["respondent"] or t.respondent
        p["tasks"].append(t)
        place(p, t.workplace_id, _period_of(t.created_at), t.created_at)
    for m in MonthlyResponse.query.order_by(MonthlyResponse.period.desc(), MonthlyResponse.id.desc()).all():
        p = person(m.emp_id, m.name)
        if p is None:
            continue
        p["monthly"].append(m)
        place(p, m.workplace_id, m.period, m.created_at)
    for r in Respondent.query.order_by(Respondent.id).all():
        p = people.get(person_key(r.emp_id, r.name))
        if p is not None and p["respondent"] is None:
            p["respondent"] = r
    return people

def latest_stance(responses):
    for m in responses:
        if m.answer("stance"):
            return m.answer("stance"), m
    return "", None

def stance_since():
    return shift_period(current_period(), 1 - C.STANCE_MONTHS)

def person_stance(emp_id):
    if not emp_key(emp_id):
        return "", None
    rows = (MonthlyResponse.query.filter(same_emp_id(MonthlyResponse.emp_id, emp_id), MonthlyResponse.period >= stance_since())
            .order_by(MonthlyResponse.period.desc(), MonthlyResponse.id.desc()).all())
    return latest_stance(rows)

# ----- DX人財情報 -----
def skills_summary(people=None):
    have = Counter()
    learn = Counter()
    for p in dx_people() if people is None else people:
        have.update(p["skills"])
        learn.update(p["learn"])
    return have, learn

def _positive(last_check, stance, stance_row):
    if last_check is not None and (stance_row is None or last_check >= (stance_row.created_at or datetime.min)):
        return True
    return stance in C.MONTHLY_KPI["self_positive"]

_TEAM_TEXT_FIELDS = ("assignee", "progress_note", "result_note") + tuple("ecrs_" + key for key, _label, _q in ECRS)


def _team_touched(t):
    assignment = t.assignment or "未定"
    by_team = assignment not in ("未定", "自主改善") or (assignment == "自主改善" and t.self_improve not in C.SELF_IMPROVE_POSITIVE)
    return bool(by_team or t.verified or (t.decision or "未定") != "未定" or t.after_minutes is not None or t.steps_done_list
                or any((getattr(t, f) or "").strip() for f in _TEAM_TEXT_FIELDS) or t.stage_logs or t.ai_extra_dict.get("info_edited_at"))

def _accompanied(t):
    return t.assignment == "自主改善" and (t.stage_index > 0 or _team_touched(t))

def dx_people(workplace_id=None, wish=None):
    names = {emp_key(e.emp_id): e.name for e in Employee.query.all() if e.name}
    places = {w.id: w for w in Workplace.query.all()}
    wish_rank = {opt: i for i, opt in enumerate(C.DX_WISH_OPTIONS)}
    stance_rank = {opt: i for i, opt in enumerate(C.STANCE_OPTIONS)}
    since = stance_since()
    rows = []
    for p in answered_people().values():
        r, tasks, monthly = p["respondent"], p["tasks"], p["monthly"]
        latest = monthly[0] if monthly else None
        checked = [t for t in tasks if t.self_improve in C.SELF_IMPROVE_POSITIVE]
        accompanied = [t for t in tasks if _accompanied(t)]
        stance, stance_row = latest_stance(monthly)
        recent, recent_row = latest_stance([m for m in monthly if m.period >= since])
        last_check = max((t.created_at or datetime.min for t in checked), default=None)
        asked_at = (recent_row.created_at or datetime.min) if recent in C.MONTHLY_KPI["ask_positive"] else None
        newest_ask = asked_at is not None and (last_check is None or asked_at > last_check)
        later_ask = newest_ask and last_check is not None
        review = newest_ask and any(asked_at > (t.created_at or datetime.min) for t in accompanied)
        raw_skills, raw_learn = (r.skills, r.learn) if r is not None else ([], [])
        want_row = next((m for m in monthly if m.answer("wants") == "ある" and task_names(m.answer("wants_tasks"))), None) if not tasks else None
        p.update(
            name=(r.name if r is not None else "") or (latest.name if latest else "") or p["name"] or names.get(p["key"][1], ""),
            workplace=places.get(p["workplace_id"]) or (r.workplace if r is not None else None),
            work_pattern=(r.work_pattern if r is not None else "") or next((m.work_pattern for m in monthly if m.work_pattern), ""),
            wish=C.SELF_IMPROVE_YES if checked else (C.SELF_IMPROVE_UNCHECKED if tasks else ""),
            wish_note=C.SELF_IMPROVE_LATER_ASK if later_ask or review else "",
            wish_hint=C.ASSIGNMENT_REVIEW if review else "",
            skills=[s for s in raw_skills if s != C.SKILL_NONE], learn=[s for s in raw_learn if s != C.SKILL_NONE],
            skills_none=C.SKILL_NONE in raw_skills, learn_none=C.SKILL_NONE in raw_learn,
            task_count=len(tasks), checked_count=len(checked), minutes=sum(t.monthly_minutes or 0 for t in tasks),
            target_tasks=[{"task": t, "label": C.BUILD_STAGE_SELF if t in accompanied else "希望"} for t in tasks if t in checked or t.assignment == "自主改善"],
            latest_period=latest.period if latest else "",
            stance_latest=stance, stance_label=period_label(stance_row.period) if stance_row else "",
            stance_old=stance_row is not None and stance_row.period < since,
            self_positive_count=sum(1 for m in monthly if m.answer("stance") in C.MONTHLY_KPI["self_positive"]),
            positive=_positive(last_check, recent, recent_row) or bool(accompanied),
            monthly_wants=monthly_task_names(want_row.answer("wants_tasks")) if want_row else [],
            monthly_wants_id=want_row.id if want_row else None, monthly_wants_label=period_label(want_row.period) if want_row else "",
        )
        p["workplace_id"] = p["workplace"].id if p["workplace"] is not None else None
        p["wish_positive"] = p["wish"] in C.SELF_IMPROVE_POSITIVE and not later_ask
        if workplace_id and p["workplace_id"] != workplace_id:
            continue
        if wish == "positive" and not p["positive"]:
            continue
        if wish and wish != "positive" and p["wish"] != wish:
            continue
        rows.append(p)
    rows.sort(key=lambda p: (not p["positive"], wish_rank.get(p["wish"], 9), 9 if p["stance_old"] else stance_rank.get(p["stance_latest"], 9), p["workplace_id"] or 0, p["name"]))
    return rows

def wish_text(p):
    note = "・".join(x for x in (p["wish_note"], p["wish_hint"]) if x)
    return p["wish"] + (f"（{note}）" if p["wish"] and note else "")

def stance_period_text(p):
    return p["stance_label"] + (f"（{C.STANCE_MONTHS}か月より前）" if p["stance_old"] else "")

# ----- 次のアクション提案（ルールベース） -----
def suggest_actions(limit=10):
    suggestions = []
    tasks = Task.query.all()
    domain_rows = domain_analysis()

    for row in domain_rows:
        if row["level"] <= 1:
            suggestions.append(
                {
                    "kind": "調査",
                    "title": f"「{row['domain'].name}」領域の業務情報が不足しています",
                    "detail": f"登録 {row['count']} 件・必須項目の充足率 {row['fill_rate']:.0f}%。担当者へのヒアリングで業務の実態を把握してください。",
                    "url": ("bizmap.domain", {"domain_id": row["domain"].id}),
                }
            )

    for t in candidate_tasks(limit=3):
        suggestions.append(
            {
                "kind": "案件",
                "title": f"月間 {t.monthly_minutes / 60:.1f} 時間の「{t.name}」を優先案件に",
                "detail": f"改善希望: {t.wish or '－'}／困りごと: {'・'.join(t.pains_list) or '－'}。標準手順（アウトプット確認 → 目的 → ECRS → 自動化）で検討してください。",
                "url": ("cases.detail", {"task_id": t.id}),
            }
        )

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

    for group in duplicate_output_groups()[:2]:
        suggestions.append(
            {
                "kind": "ECRS 結合",
                "title": f"同じアウトプット「{group['output']}」を {len(group['tasks'])} 人が作成しています",
                "detail": f"合計 月間 {group['minutes'] / 60:.1f} 時間。1つにまとめる、または上流で自動生成する統合候補です。",
                "url": ("bizmap.index", {}),
            }
        )

    for row in domain_rows:
        if row["candidates"] >= 5:
            suggestions.append(
                {
                    "kind": "業務設計",
                    "title": f"「{row['domain'].name}」領域に改善案件候補が {row['candidates']} 件集中しています",
                    "detail": "個別の自動化ではなく、業務プロセス全体（To-Be）の再設計を検討してください。",
                    "url": ("bizmap.domain", {"domain_id": row["domain"].id}),
                }
            )

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

    quiet = [w["workplace"].name for w in workplace_stats() if not w["ever_answered"]]
    if quiet:
        suggestions.append(
            {
                "kind": "回答促進",
                "title": f"まだ回答のない職場: {'、'.join(quiet)}",
                "detail": "月次アンケートの回答も「自動化を検討」の登録もまだありません。推進メンバーから職場に声をかけ、回答をお願いしてください。",
                "url": ("admin.index", {}),
            }
        )

    return suggestions[:limit]

# ----- 月次アンケートの集計（〇月度ごと・時系列） -----
def _period_of(dt):
    return dt.strftime("%Y-%m") if dt else None

def member_rows():
    answered = defaultdict(set)
    for emp_id, period in db.session.query(MonthlyResponse.emp_id, MonthlyResponse.period).all():
        answered[emp_key(emp_id)].add(period)
    considered, first_task = Counter(), {}
    for emp_id, created in db.session.query(Respondent.emp_id, Task.created_at).join(Task, Task.respondent_id == Respondent.id).all():
        key = emp_key(emp_id)
        considered[key] += 1
        p = _period_of(created)
        if p and (key not in first_task or p < first_task[key]):
            first_task[key] = p
    saved = {emp_key(x) for (x,) in db.session.query(Respondent.emp_id).all()}
    rows = []
    for e in Employee.query.order_by(Employee.emp_id).all():
        key = emp_key(e.emp_id)
        periods = answered.get(key, set())
        starts = [p for p in (_period_of(e.created_at), min(periods) if periods else None, first_task.get(key)) if p]
        rows.append({"emp": e, "answered": periods, "since": min(starts) if starts else None, "considered": considered.get(key, 0),
                     "deletable": not periods and key not in saved})
    return rows

def member_periods():
    periods = {x.period for x in MonthlySurvey.query.all()} | {p for (p,) in db.session.query(MonthlyResponse.period).distinct().all()}
    return sorted(p for p in periods if p and p <= current_period())

def member_in(emp, period):
    if emp.active:
        return True
    return emp.excluded_at is not None and period < _period_of(emp.excluded_at)

def member_stats(rows, periods):
    stats = {}
    for p in periods:
        members = [r for r in rows if member_in(r["emp"], p) and (r["since"] is None or r["since"] <= p)]
        answered = len([r for r in members if p in r["answered"]])
        stats[p] = {"members": len(members), "answered": answered, "rate": (answered / len(members) * 100.0) if members else None}
    return stats

def monthly_summary():
    periods, by_period = _responses_by_period()
    member = member_stats(member_rows(), periods)
    kpi = C.MONTHLY_KPI
    rows = []
    for p in periods:
        rs = by_period[p]
        n = len(rs)

        def rate(key, positive):
            answered = [r for r in rs if r.answer(key)]
            return (len([r for r in answered if r.answer(key) in positive]) / len(answered) * 100.0) if answered else 0.0

        pattern_counts = Counter((r.work_pattern or "未設定") for r in rs)
        rows.append(
            {
                "period": p,
                "label": period_label(p),
                "n": n,
                "response_rate": member[p]["rate"],
                "member_answered": member[p]["answered"],
                "members": member[p]["members"],
                "has_rate": rate("has", kpi["has_positive"]),
                "hours_month": _hours_month(rs),
                "pattern_counts": {k: pattern_counts[k] for k in list(C.WORK_PATTERNS) + ["未設定"] if pattern_counts.get(k)},
                "decreased_rate": rate("trend", kpi["trend_decreased"]),
                "increased_rate": rate("trend", kpi["trend_increased"]),
                "wants_rate": rate("wants", kpi["wants_positive"]),
                "ask_rate": rate("stance", kpi["ask_positive"]),
                "self_rate": rate("stance", kpi["self_positive"]),
                "together_rate": rate("stance", kpi["together_positive"]),
                "self_only_rate": rate("stance", kpi["self_only"]),
                "new_count": len([r for r in rs if r.answer("new_tasks").strip()]),
                "comment_count": len([r for r in rs if (r.comment or "").strip()]),
            }
        )
    return rows

def _responses_by_period():
    responses = MonthlyResponse.query.all()
    periods = sorted({r.period for r in responses})
    return periods, {p: [r for r in responses if r.period == p] for p in periods}

def _hours_month(responses):
    return sum(r.hours_shift * C.shifts_per_month(r.work_pattern) for r in responses)

def monthly_question_trends():
    periods, by_period = _responses_by_period()
    trends = []
    for q in C.MONTHLY_CORE_QUESTIONS:
        if q["type"] != "choice":
            continue
        group = C.MONTHLY_TREND_GROUPS.get(q["key"])
        to_band = {label: band for band, labels in group.items() for label in labels} if group else None
        if to_band:
            counters = [Counter(to_band[r.answer(q["key"])] for r in by_period[p] if r.answer(q["key"]) in to_band) for p in periods]
        else:
            counters = [Counter(r.answer(q["key"]) for r in by_period[p] if r.answer(q["key"])) for p in periods]
        n_list = [sum(c.values()) for c in counters]
        series = []
        for opt in (list(group) if group else q["options"]):
            counts = [c.get(opt, 0) for c in counters]
            series.append({"option": opt, "percents": [round(c / n * 100) if n else 0 for c, n in zip(counts, n_list)]})
        note = "以前の区切り（30分・1時間・2時間・4時間）にまとめて表示。細かい区分は月度の詳細で見られます" if group else ""
        trends.append({"q": q, "label": C.MONTHLY_KPI_LABELS.get(q["key"], ""), "note": note, "periods": periods, "labels": [period_label(p) for p in periods], "n": n_list, "series": series})
    return trends

def _trend_cells(periods, by_period, match):
    kpi = C.MONTHLY_KPI
    cells = []
    for p in periods:
        sub = [r for r in by_period[p] if match(r)]
        n = len(sub)
        cells.append(
            {
                "n": n,
                "has_rate": (len([r for r in sub if r.answer("has") in kpi["has_positive"]]) / n * 100.0) if n else None,
                "hours_month": _hours_month(sub) if n else None,
            }
        )
    return cells

def _trend_table(name, items, match):
    periods, by_period = _responses_by_period()
    rows = []
    for item in items:
        cells = _trend_cells(periods, by_period, lambda r: match(r, item))
        if any(c["n"] for c in cells):
            rows.append({name: item, "cells": cells})
    return {"labels": [period_label(p) for p in periods], "rows": rows}

def monthly_workplace_trends():
    return _trend_table("workplace", Workplace.query.order_by(Workplace.sort, Workplace.id).all(), lambda r, wp: r.workplace_id == wp.id)

def monthly_pattern_trends():
    return _trend_table("pattern", list(C.WORK_PATTERNS) + ["未設定"], lambda r, pattern: (r.work_pattern or "未設定") == pattern)

def monthly_question_stats(survey, responses):
    stats = []
    for q in survey.questions:
        key = q["key"]
        if q["type"] == "choice":
            counts = Counter(r.answer(key) for r in responses if r.answer(key))
            n = sum(counts.values())
            options = list(q.get("options", []))
            options += [o for o in counts if o not in options]
            stats.append(
                {
                    "q": q,
                    "n": n,
                    "options": options,
                    "counts": [counts.get(o, 0) for o in options],
                    "percents": [round(counts.get(o, 0) / n * 100) if n else 0 for o in options],
                }
            )
        else:
            items = [(r.name, r.workplace.name if r.workplace else "", r.answer(key)) for r in responses if r.answer(key).strip()]
            stats.append({"q": q, "n": len(items), "items": items})
    return stats

# ----- 活動状況（開始 → 要求定義 → 要件定義 → 作成/伴走 → テスト → 確認 → 完了） -----
def board_data(tasks, done_limit=8):
    floor = datetime.min

    def since(t):
        return t.stage_since or floor

    active = sorted([t for t in tasks if t.status == "対応中"], key=lambda t: (t.stage_index, since(t)), reverse=True)
    hold = sorted([t for t in tasks if t.status == "保留"], key=since, reverse=True)
    done = sorted([t for t in tasks if t.status == "完了"], key=lambda t: t.completed_at or since(t), reverse=True)
    waiting = sorted([t for t in tasks if t.status == "受付"], key=lambda t: t.created_at or floor, reverse=True)
    return {
        "active": active,
        "hold": hold,
        "done": done,
        "done_shown": done[:done_limit] if done_limit else done,
        "waiting": waiting,
        "waiting_total": len(waiting),
    }

# ===== 業務マップの区分（作業のだれ・情報の媒体・情報の流れの方法・仕事の流れの引き継ぎ方。管理画面で変えられる） =====
MAP_VOCAB_KEY = "map_vocab"
VOCAB_GROUPS = ("actors", "media", "methods")
SHAPE_CHOICES = OrderedDict([("doc", "書類（下が波）"), ("cyl", "円柱"), ("stadium", "丸い端")])
DIRECTION_CHOICES = OrderedDict([("push", "前が伝えに行く"), ("pull", "後ろが取りに行く"), ("", "決めない")])
METHOD_KIND_CHOICES = OrderedDict([("channel", "渡し方"), ("place", "置き場")])
USER_CHOICES = OrderedDict([("職場内", "職場内"), ("上司", "上司"), ("他部署", "他部署"), ("顧客・社外", "顧客・社外"), ("分からない", "分からない"), ("", "なし")])
ICON_CHOICES = [
    "bi-file-earmark-excel-fill", "bi-file-earmark-word-fill", "bi-file-earmark-ppt-fill", "bi-file-earmark-pdf-fill", "bi-filetype-csv",
    "bi-file-earmark-text-fill", "bi-file-earmark", "bi-file-earmark-image-fill", "bi-database-fill", "bi-globe2", "bi-envelope-fill",
    "bi-chat-dots-fill", "bi-chat-quote-fill", "bi-easel-fill", "bi-telephone-fill", "bi-person-fill", "bi-person-badge-fill",
    "bi-people-fill", "bi-building", "bi-briefcase-fill", "bi-truck", "bi-cpu-fill", "bi-question-circle-fill", "bi-question-square",
    "bi-folder-fill", "bi-printer-fill", "bi-clipboard-check-fill", "bi-camera-fill", "bi-usb-drive-fill", "bi-shop",
]
DEFAULT_MAP_VOCAB = {
    "actors": [
        {"key": "external", "label": "社外・顧客", "words": "顧客、社外、取引先、お客、仕入先、協力会社、外注、業者、メーカー、外部", "exclude": "作業者", "icon": "bi-briefcase-fill", "user": "顧客・社外"},
        {"key": "dept", "label": "他部署", "words": "他部署、部署、倉庫、購買、営業、総務、経理、人事、品証、品質保証、～課、～係、～センター、～グループ", "exclude": "", "icon": "bi-building", "user": "他部署"},
        {"key": "boss", "label": "上司・役職", "words": "上司、課長、係長、部長、班長、職長、組長、工場長、リーダー、主任、次長、副課長", "exclude": "", "icon": "bi-person-badge-fill", "user": "上司"},
        {"key": "team", "label": "職場内", "words": "職場、チーム、メンバー、全員、みんな、本人、自分、担当、作業者、検査員、オペレーター、オペレータ、課内、班内、同僚、先輩、後輩、ライン、前の勤務、次の勤務", "exclude": "オンライン", "icon": "bi-people-fill", "user": "職場内"},
        {"key": "system", "label": "システム（自動）", "words": "自動、RPA、マクロ、バッチ", "exclude": "", "icon": "bi-cpu-fill", "user": ""},
        {"key": "unknown", "label": "不明", "words": "不明、分からない、わからない", "exclude": "", "icon": "bi-question-circle-fill", "user": "分からない"},
    ],
    "media": [
        {"key": "excel", "label": "Excel", "words": "Excel、エクセル、xlsx、xlsm、xls、スプレッドシート、表計算", "exclude": "", "icon": "bi-file-earmark-excel-fill", "shape": "doc"},
        {"key": "word", "label": "Word", "words": "Word、ワード、docx", "exclude": "パスワード、キーワード", "icon": "bi-file-earmark-word-fill", "shape": "doc"},
        {"key": "ppt", "label": "PowerPoint", "words": "PowerPoint、パワーポイント、パワポ、ppt、pptx、スライド", "exclude": "", "icon": "bi-file-earmark-ppt-fill", "shape": "doc"},
        {"key": "pdf", "label": "PDF", "words": "PDF", "exclude": "", "icon": "bi-file-earmark-pdf-fill", "shape": "doc"},
        {"key": "csv", "label": "CSV", "words": "CSV", "exclude": "", "icon": "bi-filetype-csv", "shape": "doc"},
        {"key": "paper", "label": "紙・帳票", "words": "紙、帳票、伝票、用紙、手書き、印刷、プリント、チェックシート、チェック表、記録表、メモ", "exclude": "", "icon": "bi-file-earmark-text-fill", "shape": "doc"},
        {"key": "photo", "label": "写真・画像", "words": "写真、画像、スクリーンショット、スクショ", "exclude": "", "icon": "bi-file-earmark-image-fill", "shape": "doc"},
        {"key": "system", "label": "システムのデータ", "words": "システム、DB、データベース、SAP、基幹、MES、ERP、Access、サーバ、データ、モニタ、モニター、画面、端末、ディスプレイ、タブレット、表示器、表示盤、パネル、HMI", "exclude": "ファイルサーバ", "icon": "bi-database-fill", "shape": "cyl"},
        {"key": "web", "label": "Web", "words": "Web、ウェブ、ホームページ、サイト、ブラウザ、ポータル、イントラ、URL", "exclude": "", "icon": "bi-globe2", "shape": "stadium"},
        {"key": "mail", "label": "メール本文", "words": "メール本文、～メール", "exclude": "", "icon": "bi-envelope-fill", "shape": "stadium"},
        {"key": "chat", "label": "チャット", "words": "チャット、Teams、Slack", "exclude": "", "icon": "bi-chat-dots-fill", "shape": "stadium"},
        {"key": "verbal", "label": "口頭", "words": "口頭、伝言、聞いた内容、言われたこと、指示", "exclude": "指示書", "icon": "bi-chat-quote-fill", "shape": "stadium"},
        {"key": "board", "label": "掲示・ホワイトボード", "words": "ホワイトボード、掲示物、掲示板、アンドン", "exclude": "", "icon": "bi-easel-fill", "shape": "stadium"},
        {"key": "doc", "label": "その他の書類", "words": "日報、週報、月報、年報、報告書、集計表、一覧表、管理表、台帳、記録、帳簿、計画書、指示書、手順書、資料、リスト、～表、～書、～票、～簿", "exclude": "", "icon": "bi-file-earmark", "shape": "doc"},
        {"key": "unknown", "label": "不明", "words": "", "exclude": "", "icon": "bi-question-square", "shape": "doc"},
    ],
    "methods": [
        {"key": "phs", "label": "PHS", "words": "PHS、ピッチ", "exclude": "", "kind": "channel", "direction": "", "medium": "verbal"},
        {"key": "phone", "label": "電話", "words": "電話、TEL、内線、携帯", "exclude": "", "kind": "channel", "direction": "", "medium": "verbal"},
        {"key": "mail", "label": "メール", "words": "メール、mail、Outlook", "exclude": "", "kind": "channel", "direction": "push", "medium": "mail"},
        {"key": "chat", "label": "チャット（Teams）", "words": "Teams、チャット、Slack", "exclude": "", "kind": "channel", "direction": "push", "medium": "chat"},
        {"key": "fax", "label": "FAX", "words": "FAX、ファックス、ファクス", "exclude": "", "kind": "channel", "direction": "push", "medium": "paper"},
        {"key": "meeting", "label": "会議・朝礼", "words": "会議、朝礼、朝会、ミーティング、打ち合わせ、打合せ、MTG", "exclude": "", "kind": "channel", "direction": "", "medium": "verbal"},
        {"key": "board", "label": "掲示", "words": "掲示、ホワイトボード、貼り出、アンドン", "exclude": "", "kind": "channel", "direction": "pull", "medium": "board"},
        {"key": "hand", "label": "手渡し", "words": "手渡し、手で渡、持っていく、持参、回覧、社内便、郵送、紙で届、紙で渡、紙で受け取", "exclude": "", "kind": "channel", "direction": "push", "medium": "paper"},
        {"key": "folder", "label": "共有フォルダ", "words": "共有フォルダ、フォルダ、ファイルサーバ、共有ドライブ、SharePoint、OneDrive、NAS", "exclude": "", "kind": "place", "direction": "pull", "medium": "doc"},
        {"key": "system", "label": "システム（登録・参照）", "words": "システム、DB、データベース、SAP、基幹、MES、ERP", "exclude": "", "kind": "place", "direction": "pull", "medium": "system"},
        {"key": "verbal", "label": "口頭", "words": "口頭、対面、声をかけ、直接伝え、聞く、人に聞、聞いて", "exclude": "", "kind": "channel", "direction": "", "medium": "verbal"},
        {"key": "other", "label": "その他", "words": "その他", "exclude": "", "kind": "channel", "direction": "", "medium": "unknown"},
    ],
    "handoff": {
        "push": {"label": "前が伝えに行く", "words": "渡す、渡し、送る、送付、届ける、配る、配布、提出、報告、連絡、知らせる、伝える、共有する、掲示する、回覧する、届く、来る、送られて、もらう、受ける、受け取る、連絡が入る"},
        "pull": {"label": "後ろが取りに行く", "words": "集める、回収、取りに、取ってくる、見に行く、見に、見る、聞く、聞き、問い合わせ、確認しに、参照、開く、ダウンロード、抽出、引き出"},
    },
}
BUILTIN_VOCAB_KEYS = {g: {r["key"] for r in DEFAULT_MAP_VOCAB[g]} for g in VOCAB_GROUPS}
_VOCAB_CACHE = {"raw": None, "value": None}
_TILDES = str.maketrans({"〜": "~", "～": "~", "∼": "~"})
_SUFFIX_AFTER = r"(?=$|[\s　（()）「」『』【】、。・,.:：/／!?！？]|[ぁ-ゟ])"


def vnorm(text):
    return unicodedata.normalize("NFKC", str(text or "")).translate(_TILDES)

def _split_words(value):
    return [w.strip() for w in re.split(r"[、,，\n]+", vnorm(value)) if w.strip() and w.strip() != "~"]

def _word_pattern(word):
    suffix = word.startswith("~")
    w = word[1:].strip() if suffix else word
    if not w:
        return None
    p = re.escape(w)
    if re.fullmatch(r"[A-Za-z0-9 ._-]+", w):
        p = r"(?<![A-Za-z0-9])" + p + r"(?![A-Za-z])"
    if suffix:
        p = r"(?<=.)" + p + _SUFFIX_AFTER
    return p

def _row_pattern(row):
    return "|".join(p for p in (_word_pattern(w) for w in _split_words(row.get("words"))) if p)

def _default_vocab():
    return json.loads(json.dumps(DEFAULT_MAP_VOCAB))

def _clean_vocab(data):
    base = _default_vocab()
    if not isinstance(data, dict):
        return base
    out = {}
    try:
        seq = max(0, int(data.get("_seq") or 0))
    except (TypeError, ValueError):
        seq = 0
    for group in VOCAB_GROUPS:
        rows = data.get(group) if isinstance(data.get(group), list) else []
        defaults = {r["key"]: r for r in base[group]}
        # 1回目: 行に付いてきたキーを先に確保する（並べ替えや追加でキーが変わらないように）
        incoming, taken = [], set()
        for r in rows:
            if not isinstance(r, dict):
                continue
            key = re.sub(r"[^a-z0-9_]", "", str(r.get("key") or "").lower())[:20]
            if key and key not in taken and (key in defaults or re.fullmatch(r"u\d{1,6}", key)):
                taken.add(key)
                if key not in defaults:
                    seq = max(seq, int(key[1:]))
            else:
                key = ""
            incoming.append((r, key))
        # 2回目: キーのない行（追加した行）に、まだ一度も使っていないキーを付ける
        seen, clean = set(), []
        for r, key in incoming:
            if not key:
                seq += 1
                key = f"u{seq}"
            label = str(r.get("label") or "").strip()[:30] or (defaults[key]["label"] if key in defaults else "")
            if not label:
                continue
            seen.add(key)
            row = {"key": key, "label": label, "words": "、".join(_split_words(r.get("words")))[:600], "exclude": "、".join(_split_words(r.get("exclude")))[:300]}
            if group != "methods":
                row["icon"] = r.get("icon") if r.get("icon") in ICON_CHOICES else (defaults[key]["icon"] if key in defaults else "bi-question-square")
            if group == "actors":
                row["user"] = r.get("user") if r.get("user") in USER_CHOICES else (defaults[key]["user"] if key in defaults else "")
            if group == "media":
                row["shape"] = r.get("shape") if r.get("shape") in SHAPE_CHOICES else "doc"
            if group == "methods":
                row["kind"] = r.get("kind") if r.get("kind") in METHOD_KIND_CHOICES else "channel"
                row["direction"] = r.get("direction") if r.get("direction") in DIRECTION_CHOICES else ""
                row["medium"] = str(r.get("medium") or "unknown")[:20]
            clean.append(row)
        for key, d in defaults.items():
            if key not in seen:
                clean.append(d)
        out[group] = clean
    media_keys = {r["key"] for r in out["media"]}
    for r in out["methods"]:
        if r["medium"] not in media_keys:
            r["medium"] = "unknown"
    out["_seq"] = seq
    hand = data.get("handoff") if isinstance(data.get("handoff"), dict) else {}
    out["handoff"] = {}
    for k, d in base["handoff"].items():
        h = hand.get(k) if isinstance(hand.get(k), dict) else {}
        out["handoff"][k] = {"label": str(h.get("label") or "").strip()[:20] or d["label"], "words": "、".join(_split_words(h.get("words") if "words" in h else d["words"]))[:600]}
    return out

def _compile_vocab(vocab):
    for group in VOCAB_GROUPS:
        for row in vocab[group]:
            pat = _row_pattern(row)
            row["_rx"] = re.compile(pat, re.I) if pat else None
            row["_ex"] = _split_words(row.get("exclude"))
    for row in vocab["handoff"].values():
        pat = _row_pattern(row)
        row["_rx"] = re.compile(pat, re.I) if pat else None
    return vocab

def map_vocab():
    try:
        raw = Setting.get(MAP_VOCAB_KEY, "")
    except Exception:
        raw = ""
    if _VOCAB_CACHE["raw"] != raw or _VOCAB_CACHE["value"] is None:
        try:
            data = json.loads(raw) if raw else None
        except ValueError:
            data = None
        _VOCAB_CACHE.update(raw=raw, value=_compile_vocab(_clean_vocab(data)))
    return _VOCAB_CACHE["value"]

def vocab_public(vocab=None):
    vocab = vocab or map_vocab()
    out = {"handoff": {k: {"label": v["label"], "words": v["words"]} for k, v in vocab["handoff"].items()}, "_seq": vocab.get("_seq", 0)}
    for group in VOCAB_GROUPS:
        out[group] = [dict({k: v for k, v in r.items() if not k.startswith("_")}, builtin=r["key"] in BUILTIN_VOCAB_KEYS[group]) for r in vocab[group]]
    return out

def _masked(text, row):
    t = vnorm(text)
    for ex in row["_ex"]:
        t = t.replace(ex, " " * len(ex))
    return t

def vocab_match(group, text, vocab=None):
    if not str(text or "").strip():
        return None
    for row in (vocab or map_vocab())[group]:
        if row["_rx"] and row["_rx"].search(_masked(text, row)):
            return row["key"]
    return None

def vocab_row(group, key, vocab=None):
    return next((r for r in (vocab or map_vocab())[group] if r["key"] == key), None)

def vocab_label(group, key):
    row = map_vocab()["handoff"].get(key) if group == "handoff" else vocab_row(group, key)
    return row["label"] if row else ""

def medium_of(name, hint=None, vocab=None):
    return vocab_match("media", name, vocab) or (hint if hint and vocab_row("media", hint, vocab) else None) or "unknown"

def method_of(text, vocab=None):
    return vocab_match("methods", text, vocab) or ""

def field_role(text, vocab=None):
    """手順の system・渡し方の欄を、線の方法（line）・箱と方法（both）・箱だけ（box）のどれにするか"""
    m = method_of(text, vocab)
    if not m:
        return "box", ""
    return ("both" if vocab_row("methods", m, vocab)["kind"] == "place" else "line"), m

def handoff_of(*texts, side="", method="", vocab=None):
    vocab = vocab or map_vocab()
    hits = set()
    for t in texts:
        t = vnorm(t)
        for key, row in vocab["handoff"].items():
            if t and row["_rx"] and row["_rx"].search(t):
                hits.add(key)
    if len(hits) == 1:
        return hits.pop()
    if side in ("push", "pull"):
        return side
    row = vocab_row("methods", method, vocab) if method else None
    return row["direction"] if row else ""

# ----- 氏名の置き換えと作業のだれ（社員名簿・職場マスタ・区分の語を使う） -----
_HONORIFIC_RX = re.compile(r"(さん|様|さま|氏|くん|君|ちゃん)$")
_NAME_HONORIFIC_RX = re.compile(r"((?:お|ご)?(?:[一-龥々]{1,4}|[ァ-ヶー]{2,8}|[A-Za-z]{2,12}))[ 　]?(さん|さま|くん|ちゃん|(?:様|氏|君)(?![一-龥々]))")
_NAME_SEP_RX = re.compile(r"([、,，／/\n・]|[ 　]+)")
_NOT_PERSON_WORDS = {"皆", "お客", "客", "先方", "相手", "先様", "お得意", "各", "副", "前", "元", "新", "旧", "当", "同", "全", "直", "現場", "職場"}
# 「佐藤課長」「佐藤」のように役職の前や単独で出てきたときだけ氏名とみなす名字（社員名簿の名字も加える）
COMMON_SURNAMES = frozenset((
    "佐藤、鈴木、高橋、田中、伊藤、渡辺、山本、中村、小林、加藤、吉田、山田、佐々木、山口、松本、井上、木村、斎藤、清水、山崎、池田、橋本、阿部、石川、"
    "山下、中島、石井、小川、前田、岡田、長谷川、藤田、後藤、近藤、村上、遠藤、青木、坂本、斉藤、福田、太田、西村、藤井、金子、岡本、藤原、中野、三浦、"
    "原田、中川、松田、竹内、小野、田村、中山、和田、石田、森田、上田、内田、柴田、酒井、宮崎、横山、高木、安藤、宮本、大野、小島、谷口、今井、工藤、"
    "高田、増田、丸山、杉山、村田、大塚、新井、小山、平野、藤本、河野、上野、野口、武田、千葉、岩崎、佐野、菅原、木下、久保、野村、松尾、菊地、杉本、"
    "市川、大西、古川、島田、水野、桜井、高野、渡部、吉川、山内、西田、飯田、菊池、西川、小松、北村、安田、五十嵐、川口、平田、中田、久保田、服部、岩田、"
    "土屋、川崎、福島、本田、樋口、秋山、田口、永井、山中、中西、吉村、川上、石原、大橋、松井、馬場、浅野、荒木、大久保、野田、小池、松下、熊谷、菅野、"
    "内藤、黒田、尾崎、永田、川村、望月、田辺、松岡、本間、荒井").split("、"))


WORKPLACE_ALIASES_KEY = "workplace_aliases"


def workplace_aliases():
    try:
        data = _json_column(Setting.get(WORKPLACE_ALIASES_KEY, ""), {})
    except Exception:
        data = {}
    return data if isinstance(data, dict) else {}

def remember_workplace_name(workplace_id, old, new):
    data = workplace_aliases()
    names = [n for n in data.get(str(workplace_id)) or [] if isinstance(n, str) and vnorm(n).strip() != vnorm(new).strip()]
    if (old or "").strip() and vnorm(old).strip() != vnorm(new).strip() and old not in names:
        names.append(old)
    data[str(workplace_id)] = names
    Setting.set(WORKPLACE_ALIASES_KEY, json.dumps(data, ensure_ascii=False))

def _workplace_names():
    """[(見つける名前, 職場名)] 長い順。「検査・品質保証課」は「品質保証課」でも見つける。1回のリクエストの中では使い回す"""
    if has_request_context() and "wp_names" in g:
        return g.wp_names
    pairs = []
    try:
        rows = [w for w in Workplace.query.all() if w.name]
        current, aliases = {vnorm(w.name).strip() for w in rows}, workplace_aliases()
        for w in rows:
            olds = [n for n in aliases.get(str(w.id)) or [] if isinstance(n, str) and n.strip() and vnorm(n).strip() not in current]
            for name in [w.name] + olds:
                pairs.append((name, w.name))
                alias = name.split("・")[-1].strip() if "・" in name else ""
                if len(alias) >= 3:
                    pairs.append((alias, w.name))
    except Exception:
        pairs = []
    pairs.sort(key=lambda p: len(p[0]), reverse=True)
    if has_request_context():
        g.wp_names = pairs
    return pairs

_WP_OTHER_BEFORE_RX = re.compile(r"(?:第?[一二三四五六七八九十]+|[A-Za-z0-9])$")
_WP_OTHER_WORDS = frozenset(("課題", "課程", "課金", "課税", "部品", "部分", "部位", "部材", "室温"))
_ASCII_PAIR_RX = re.compile(r"[A-Za-z0-9]{2}")


def _name_spans(text, name):
    spans, i = [], text.find(name) if name else -1
    while i >= 0:
        j = i + len(name)
        pair = name[-1] + text[j:j + 1]
        if not _WP_OTHER_BEFORE_RX.search(text[:i]) and pair not in _WP_OTHER_WORDS and not _ASCII_PAIR_RX.fullmatch(pair):
            spans.append((i, j))
        i = text.find(name, i + 1)
    return spans

def wp_name_in(found, text):
    nf = vnorm(found)
    return bool(nf) and bool(_name_spans(vnorm(text), nf))

def current_wp_text(text):
    nt = vnorm(text)
    taken, swaps = [], []
    for found, w in _workplace_names():
        nf, nw = vnorm(found), vnorm(w)
        current = nf in nw
        for i, j in _name_spans(nt, nf):
            if any(i < b and a < j for a, b in taken):
                continue
            taken.append((i, j))
            if not current:
                swaps.append((i, j, nw))
    for i, j, nw in sorted(swaps, reverse=True):
        nt = nt[:i] + nw + nt[j:]
    return nt if swaps else text

def _person_names():
    """社員名簿の氏名 {"full": 3文字以上の氏名（長い順）, "surnames": 名字}。1回のリクエストの中では使い回す"""
    if has_request_context() and "person_names" in g:
        return g.person_names
    full, surnames = set(), set()
    try:
        for (name,) in db.session.query(Employee.name).all():
            parts = [p for p in re.split(r"[\s　]+", (name or "").strip()) if p]
            if not parts:
                continue
            full.update({" ".join(parts), "".join(parts), "　".join(parts)})
            if len(parts) >= 2 and len(parts[0]) >= 2:
                surnames.add(parts[0])
    except Exception:
        pass
    out = {"full": sorted((n for n in full if len(n) >= 3), key=len, reverse=True), "surnames": surnames}
    if has_request_context():
        g.person_names = out
    return out

def _role_only(text, vocab=None):
    row = vocab_row("actors", "boss", vocab)
    return bool(row) and vnorm(text).strip() in set(_split_words(row.get("words")))

def _not_person(word, vocab=None):
    """敬称がついても人の名前ではない語（営業さん、お客様、検査員さん、品質保証課さん など）"""
    w = vnorm(word).strip()
    if not w or w in _NOT_PERSON_WORDS:
        return True
    if _role_only(w, vocab) or vocab_match("actors", w, vocab) or vocab_match("media", w, vocab):
        return True
    return any(found in w or (len(w) >= 2 and w in found) for found, _ in _workplace_names())

def mask_person_names(text, vocab=None):
    """氏名を置き換える。社員名簿の氏名・「〇〇さん」・単独の名字 → 担当者、「佐藤課長」→ 課長。「課長さん」は課長、「営業さん」「お客様」はそのまま"""
    text = str(text or "")
    if not text.strip():
        return text
    vocab = vocab or map_vocab()
    people = _person_names()
    for name in people["full"]:
        if name in text:
            text = re.sub(re.escape(name) + r"[ 　]?(さん|様|さま|氏|くん|君|ちゃん)?", "担当者", text)
    surnames = COMMON_SURNAMES | people["surnames"]
    roles = [r for r in _split_words((vocab_row("actors", "boss", vocab) or {}).get("words")) if not r.startswith("~")]
    if roles:
        rx = re.compile(r"([一-龥々]+?)[ 　]?(" + "|".join(re.escape(r) for r in sorted(roles, key=len, reverse=True)) + ")")

        def strip_surname(m):
            run, role = m.group(1), m.group(2)
            for n in range(min(len(run), 4), 1, -1):
                if run[-n:] in surnames:
                    return run[:-n] + role
            return m.group(0)
        text = rx.sub(strip_surname, text)

    def honorific(m):
        word = m.group(1)
        if m.group(2) == "様" and word[-1:] in "仕同多模異一両各今":
            return m.group(0)
        core = re.sub(r"^[おご]", "", word) or word
        if _role_only(word, vocab) or _role_only(core, vocab):
            return core
        if _not_person(word, vocab) or _not_person(core, vocab):
            return m.group(0)
        return "担当者"
    text = _NAME_HONORIFIC_RX.sub(honorific, text)
    return "".join("担当者" if p.strip() in surnames else p for p in _NAME_SEP_RX.split(text))

def _mask_destination(text):
    """「課長へ提出する日報作成」の「課長へ」のように、後ろに続きがある行き先の語は、だれの判定に使わない"""
    return re.sub(r"[^\s　、。をがはでとも]+(への|へ|宛て|宛に)(?=\S)", lambda m: " " * len(m.group(0)), text)

def who_of(name, own_wp="", vocab=None):
    """作業のだれ {actor, label, state}。職場マスタ → 区分の語 → 未確認 の順。氏名は担当者に置き換え、だれは未確認にする"""
    vocab = vocab or map_vocab()
    raw = mask_person_names(str(name or "").strip(), vocab).strip()
    if not raw:
        return {"actor": "", "label": "", "state": "unconfirmed"}
    if raw == "担当者":
        return {"actor": "", "label": "担当者", "state": "unconfirmed"}
    raw = current_wp_text(raw)
    nraw, nown = vnorm(raw), vnorm(own_wp)
    for found, w in _workplace_names():
        if wp_name_in(found, nraw):
            rest = nraw.replace(vnorm(found), " ").strip(" 　の・（）()")
            role = _HONORIFIC_RX.sub("", rest)
            mine = bool(nown) and vnorm(w) == nown
            if rest and _role_only(role, vocab):
                return {"actor": "boss" if mine else "dept", "label": f"{w} {role}", "state": "ok"}
            if mine:
                return {"actor": "team", "label": f"{w}の人" if not rest or vocab_match("actors", rest, vocab) == "team" else raw, "state": "ok"}
            return {"actor": "dept", "label": raw, "state": "ok"}
    bare = _HONORIFIC_RX.sub("", raw).strip()
    if _role_only(bare, vocab):
        return {"actor": "boss", "label": f"{own_wp} {bare}".strip() if own_wp else bare, "state": "ok"}
    key = vocab_match("actors", _mask_destination(bare), vocab)
    if key == "unknown":
        return {"actor": "unknown", "label": "利用者不明", "state": "unknown"}
    if key == "team" and own_wp and normalize_text(bare) in ("職場内", "職場", "本人", "自分", "全員", "みんな", "課内", "班内"):
        return {"actor": "team", "label": f"{own_wp}の人", "state": "ok"}
    if key:
        return {"actor": key, "label": raw, "state": "ok"}
    return {"actor": "", "label": raw, "state": "unconfirmed"}

def actor_user(actor, vocab=None):
    row = vocab_row("actors", actor, vocab) if actor else None
    return row.get("user", "") if row else ""


# ===== 業務マップ（1枚の流れ図: グラフ生成、Mermaid 変換、手修正のマージ、つながりの表） =====
_ESCAPES = [
    ('"', "”"), ("#", "＃"), ("<", "＜"), (">", "＞"), ("&", "＆"), ("`", "'"), ("|", "｜"),
    ("[", "［"), ("]", "］"), ("{", "｛"), ("}", "｝"), ("(", "（"), (")", "）"), ("\\", "＼"),
    ("\r", " "), ("\n", " "), (";", "；"),
]

SHAPES = {
    "rect": ('["', '"]'),
    "stadium": ('(["', '"])'),
    "cyl": ('[("', '")]'),
}
CLS_LABELS = OrderedDict([
    ("task", "登録済みの業務"), ("done", "完了した業務"), ("hold", "保留中の業務"), ("abolished", "廃止した業務"), ("work", "他の人・部署の作業"), ("pending", "未登録の業務（月次）"), ("warn", "使う人が分からない"),
    ("source", "取り出し元の情報"), ("output", "できるもの"), ("dupout", "同じできるもの（統合候補）"),
])
CLS_TYPE = {"task": "work", "done": "work", "hold": "work", "abolished": "work", "work": "work", "pending": "work", "warn": "work", "source": "info", "output": "info", "dupout": "info"}
TASK_CLASSES = ("task", "done", "hold", "abolished")
STATUS_MARKS = {"hold": "（保留）", "abolished": "（廃止済み）"}
CLASS_DEFS = [
    "classDef task fill:#e7f1ff,stroke:#0d6efd,color:#052c65",
    "classDef done fill:#d1e7dd,stroke:#198754,color:#0f5132",
    "classDef hold fill:#e7f1ff,stroke:#0d6efd,color:#052c65,stroke-dasharray:5 4",
    "classDef abolished fill:#f1f3f5,stroke:#ced4da,color:#adb5bd",
    "classDef work fill:#ffffff,stroke:#6c757d,color:#343a40",
    "classDef pending fill:#f8f9fa,stroke:#adb5bd,color:#6c757d,stroke-dasharray:4 3",
    "classDef warn fill:#f8d7da,stroke:#dc3545,color:#842029",
    "classDef source fill:#fff3cd,stroke:#e0a800,color:#664d03",
    "classDef output fill:#ffffff,stroke:#0d6efd,color:#052c65",
    "classDef dupout fill:#ffe5d0,stroke:#fd7e14,color:#7a3a00,stroke-width:2px",
]
INPUT_SOURCE_MAP = {
    "紙の帳票": ("paper", "", "紙の帳票"), "他のExcel": ("excel", "", "他のExcel"), "メール": ("mail", "mail", "メールで届く内容"),
    "人から聞く": ("verbal", "verbal", "人から聞く内容"), "その他": ("unknown", "", "その他の情報"),
}
LEGACY_CLS = {
    "task": ("work", "task"), "done": ("work", "done"), "pending": ("work", "pending"), "warn": ("work", "warn"), "work": ("work", "work"),
    "wp": ("work", "work"), "ext": ("work", "work"), "person": ("work", "work"), "note": ("work", "work"), "place": ("work", "work"), "goods": ("work", "work"),
    "source": ("info", "source"), "system": ("info", "source"), "output": ("info", "output"), "dupout": ("info", "dupout"),
}
_SHAPE_GLYPHS = {
    "rect": '<rect x="1" y="1.5" width="16" height="11" rx="0.5"/>',
    "stadium": '<rect x="1" y="1.5" width="16" height="11" rx="5.5"/>',
    "cyl": '<path d="M1 4C1 1.4 17 1.4 17 4V10C17 12.6 1 12.6 1 10Z M1 4C1 6.6 17 6.6 17 4"/>',
    "doc": '<path d="M1 1.5H17V10.6C13.5 8.4 10.5 13.6 6.2 11.8C4.2 11 2.6 10.8 1 11.4Z"/>',
}

def esc(text):
    s = str(text or "").strip()
    for a, b in _ESCAPES:
        s = s.replace(a, b)
    return s
_BREAK_AFTER = set("、。，．・：；）」』】を ")
_BREAK_PARTICLES = set("にでがはのともへやら")
_BREAK_BEFORE = set("（「『【")

def _can_break(prev, nxt):
    if nxt in _BREAK_BEFORE or prev in _BREAK_AFTER:
        return True
    return prev in _BREAK_PARTICLES and not ("\u3041" <= nxt <= "\u309f")

def wrap(text, width=14, max_lines=3):
    raw = str(text or "").strip()
    if len(raw) <= width:
        return esc(raw)
    lines, rest = [], raw
    while rest:
        if len(rest) <= width or len(lines) == max_lines - 1:
            lines.append(rest)
            break
        cut = width
        for i in range(width, int(width * 0.55), -1):
            if _can_break(rest[i - 1], rest[i]):
                cut = i
                break
        if 0 < len(rest) - cut <= 2 and len(rest) <= width + 2:
            cut = len(rest)
        lines.append(rest[:cut].rstrip())
        rest = rest[cut:].lstrip()
    if len(lines[-1]) > width + 2:
        lines[-1] = lines[-1][: width - 1] + "…"
    return "<br/>".join(esc(l) for l in lines)

def nid(prefix, text):
    return prefix + "_" + hashlib.md5(normalize_text(text).encode("utf-8")).hexdigest()[:8]

def scope_key(workplace_id=None, domain_id=None):
    return f"wp:{workplace_id or 0}|dom:{domain_id or 0}"

def shape_glyph(shape):
    return f'<svg class="ks" viewBox="0 0 18 14" aria-hidden="true">{_SHAPE_GLYPHS.get(shape, _SHAPE_GLYPHS["rect"])}</svg>'

def info_shape(medium):
    row = vocab_row("media", medium or "unknown")
    shape = row["shape"] if row else "doc"
    return shape if shape in ("doc", "cyl", "stadium") else "doc"

def node_icon(n):
    """箱のアイコン（クラス, 色の印）。情報は媒体、作業はだれ、登録済みの業務は付けない（完了は ✓）"""
    if n.get("type") == "info":
        medium = n.get("medium") or "unknown"
        row = vocab_row("media", medium) or vocab_row("media", "unknown")
        return (row["icon"] if row else "bi-question-square"), f"m-{medium}"
    cls = n.get("cls")
    if cls == "done":
        return "bi-check-circle-fill", "done"
    if cls in ("task", "hold", "abolished"):
        return "", ""
    if cls == "pending":
        return "bi-hourglass-split", "pending"
    if n.get("who_state") == "unconfirmed" or not n.get("actor"):
        return "bi-question-circle", "who-unconfirmed"
    row = vocab_row("actors", n["actor"])
    return (row["icon"] if row else "bi-person-fill"), f"a-{n['actor']}"

def show_who(n):
    who = (n.get("who") or "").strip()
    return n.get("type") == "work" and n.get("cls") in ("work", "warn") and n.get("who_state") in ("ok", "unknown") and who and normalize_text(who) != normalize_text(n.get("text"))

def node_label(n):
    icon, mark = node_icon(n)
    head = f"<i class='bi {icon} fi {mark}'></i> " if icon else ""
    body = wrap(n["text"] + STATUS_MARKS.get(n.get("cls"), ""), n.get("width") or 14)
    if show_who(n):
        return head + "<b>" + esc(n["who"]) + "</b><br/>" + body
    return head + body

def edge_label(e):
    if (e.get("label") or "").strip():
        return e["label"].strip()
    if e.get("type") == "work":
        return vocab_label("handoff", e.get("handoff")) if e.get("handoff") in ("push", "pull") else ""
    if e.get("show_method") and e.get("method"):
        return vocab_label("methods", e["method"])
    return ""

def show_method_for(method, medium):
    """媒体から分かる方法（システムのデータをシステムで見る、メール本文をメールで受け取る）はラベルに書かない"""
    row = vocab_row("methods", method) if method else None
    return bool(row) and row.get("medium") != (medium or "unknown")


class Diagram:
    def __init__(self, direction="LR"):
        self.direction = direction
        self.nodes = OrderedDict()
        self.subgraphs = OrderedDict()
        self.edges = OrderedDict()
        self.clicks = OrderedDict()

    def node(self, node_id, text, cls="task", subgraph=None, width=14, hint=None, medium=None, who=None):
        if node_id not in self.nodes:
            ntype = CLS_TYPE.get(cls, "work")
            n = {"text": str(text or "").strip(), "cls": cls, "type": ntype, "hint": hint, "width": width, "group": None, "group_title": None,
                 "medium": "", "actor": "", "who": "", "who_state": ""}
            if ntype == "info":
                n["medium"] = medium or medium_of(n["text"], hint)
            else:
                w = who or {"actor": "", "label": "", "state": "unconfirmed"}
                n.update(actor=w.get("actor") or "", who=w.get("label") or "", who_state=w.get("state") or "unconfirmed")
            self.nodes[node_id] = n
        if subgraph is not None:
            key, title = subgraph
            self.subgraphs.setdefault(key, title)
            self.nodes[node_id]["group"] = key
            self.nodes[node_id]["group_title"] = title
        return node_id

    def edge(self, a, b, kind="info", label="", method="", handoff="", show_method=False, inferred=False):
        if not a or not b or a == b:
            return
        new = {"type": kind, "label": label, "method": method or "", "handoff": handoff or "", "show_method": bool(show_method), "inferred": bool(inferred)}
        old = self.edges.get((a, b))
        if old is None:
            self.edges[(a, b)] = new
        elif old["inferred"] and not new["inferred"]:
            self.edges[(a, b)] = dict(new, method=new["method"] or old["method"], show_method=new["show_method"] or old["show_method"])
        elif not old["method"] and new["method"]:
            old.update(method=new["method"], show_method=new["show_method"])

    def info_edge(self, a, b, method, info_node=None, inferred=False):
        """情報の流れの線。方法が情報の箱（info_node。省くと a）の媒体から分かるときはラベルに書かない"""
        self.edge(a, b, "info", method=method, show_method=show_method_for(method, self.nodes[info_node or a]["medium"]), inferred=inferred)

    def click(self, node_id, url, tooltip=""):
        self.clicks[node_id] = (url, tooltip)

    def derive_roles(self):
        """情報の箱の色を線の向きから決める（登録済みの業務が作る＝できるもの、それ以外＝取り出し元）"""
        made = {b for (a, b), e in self.edges.items() if e["type"] == "info" and self.nodes.get(a, {}).get("cls") in TASK_CLASSES}
        for i, n in self.nodes.items():
            if n["type"] == "info" and n["cls"] in ("source", "output"):
                n["cls"] = "output" if i in made else "source"
        return self

    @property
    def empty(self):
        return not self.nodes

    # ---- Mermaid ----
    def _node_line(self, node_id):
        n = self.nodes[node_id]
        label = node_label(n)
        shape = info_shape(n["medium"]) if n["type"] == "info" else "rect"
        if shape == "doc":
            return f'{node_id}:::{n["cls"]}@{{ shape: doc, label: "{label}" }}'
        open_, close = SHAPES[shape]
        return f"{node_id}{open_}{label}{close}:::{n['cls']}"

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
        work_links = []
        for idx, ((a, b), e) in enumerate(self.edges.items()):
            lab = esc(edge_label(e))
            if e.get("inferred"):
                lines.append(f'  {a} -. "{lab}" .-> {b}' if lab else f"  {a} -.-> {b}")
            elif e["type"] == "work":
                lines.append(f'  {a} == "{lab}" ==> {b}' if lab else f"  {a} ==> {b}")
            else:
                lines.append(f'  {a} -- "{lab}" --> {b}' if lab else f"  {a} --> {b}")
            if e["type"] == "work":
                work_links.append(str(idx))
        if self.edges:
            lines.append("  linkStyle default stroke:#5b8def,stroke-width:1.4px")
        if work_links:
            lines.append(f"  linkStyle {','.join(work_links)} stroke:#495057,stroke-width:2.6px")
        for i, (url, tip) in self.clicks.items():
            lines.append(f'  click {i} "{url}" "{esc(tip)}"')
        return "\n".join(lines)

    # ---- グラフ JSON（エディタ・保存用） ----
    def to_graph(self):
        nodes = []
        for i, n in self.nodes.items():
            nodes.append(
                {
                    "id": i, "text": n["text"], "cls": n["cls"], "type": n["type"], "medium": n["medium"], "actor": n["actor"], "who": n["who"],
                    "who_state": n["who_state"], "hint": n.get("hint"), "group": n.get("group"), "group_title": n.get("group_title"),
                    "url": self.clicks[i][0] if i in self.clicks else None,
                    "source": "auto", "x": None, "y": None, "color": None, "text_edited": False,
                }
            )
        edges = []
        for (a, b), e in self.edges.items():
            edges.append(
                {
                    "id": f"{a}->{b}", "from": a, "to": b, "type": e["type"], "method": e["method"], "handoff": e["handoff"],
                    "show_method": e["show_method"], "inferred": e["inferred"], "label": e["label"], "source": "auto", "reversed": False, "label_edited": False,
                }
            )
        return {"schema": 2, "nodes": nodes, "edges": edges, "suppressed_nodes": [], "suppressed_edges": [], "view": None}

def _node_type(n):
    """箱の種類（work / info）。type が無ければ色の区分から"""
    return n.get("type") if n.get("type") in ("work", "info") else CLS_TYPE.get(n.get("cls"), "work")

def _cls_for(cls, ntype):
    """箱の種類に合う色の区分。合わなければ、情報は取り出し元、作業は他の人・部署の作業"""
    return cls if CLS_TYPE.get(cls) == ntype else ("source" if ntype == "info" else "work")

def graph_to_diagram(graph, direction="LR"):
    d = Diagram(direction)
    for n in graph.get("nodes", []):
        ntype = _node_type(n)
        cls = _cls_for(n.get("cls"), ntype)
        sub = (n["group"], n.get("group_title") or n["group"]) if n.get("group") else None
        who = None if ntype == "info" else {"actor": n.get("actor") or "", "label": n.get("who") or "", "state": n.get("who_state") or ("ok" if n.get("actor") else "unconfirmed")}
        d.node(n["id"], n.get("text") or "", cls, subgraph=sub, hint=n.get("hint"), medium=n.get("medium") or None, who=who)
        if n.get("url"):
            d.click(n["id"], n["url"], "案件詳細を開く")
    types = {i: x["type"] for i, x in d.nodes.items()}
    for e in graph.get("edges", []):
        a, b = (e["to"], e["from"]) if e.get("reversed") else (e["from"], e["to"])
        if a not in types or b not in types:
            continue
        kind = e.get("type") if e.get("type") in ("work", "info") else ("work" if types[a] == "work" and types[b] == "work" else "info")
        d.edge(a, b, kind, e.get("label") or "", e.get("method") or "", e.get("handoff") or "", bool(e.get("show_method")), bool(e.get("inferred")) or (types[a] == "info" and types[b] == "info"))
    return d.derive_roles()

def upgrade_graph(saved):
    """保存済みの手修正を新しい形（作業・情報・2種類の線）に読み替える。ID は変えない"""
    if not isinstance(saved, dict) or saved.get("schema") == 2:
        return saved
    out = dict(saved)
    nodes, types = [], {}
    for n in saved.get("nodes") or []:
        if not isinstance(n, dict) or not n.get("id"):
            continue
        m = dict(n)
        old = m.get("cls") or "manual"
        if old in LEGACY_CLS:
            m["type"], m["cls"] = LEGACY_CLS[old]
        else:
            m["type"] = "info" if medium_of(m.get("text")) != "unknown" else "work"
            m["cls"] = "source" if m["type"] == "info" else "work"
        if old in ("wp", "ext", "person", "warn"):
            # 昔の人・部署の箱は、名前をだれにした作業の箱にする
            m["who"] = m.get("who") or m.get("text") or ""
            w = who_of(m["who"], m.get("group_title") or "") if old != "warn" else {"actor": "unknown", "state": "unknown"}
            m["actor"], m["who_state"] = w["actor"], w["state"]
            if m.get("source") == "manual" or m.get("text_edited"):
                m["who_edited"] = True
            elif old != "warn":
                m["text"] = "使いみち未確認"
        if old == "system":
            m["medium"] = "system"
            if m.get("source") == "manual":
                m["medium_edited"] = True
        types[m["id"]] = m["type"]
        nodes.append(m)
    edges = []
    for e in saved.get("edges") or []:
        if not isinstance(e, dict) or not e.get("from") or not e.get("to"):
            continue
        m = dict(e)
        kind = m.get("kind")
        a, b = m["from"], m["to"]
        if kind in ("seq", "relation", "goods"):
            m["type"] = "work"
        elif kind == "info":
            m["type"] = "info"
        else:
            m["type"] = "work" if types.get(a) == "work" and types.get(b) == "work" else "info"
        label = (m.get("label") or "").strip()
        if label in ("次の業務", "前の業務", "関与"):
            m["label"] = ""
        elif label and m["type"] == "info" and (m.get("source") == "manual" or m.get("label_edited")) and _exact_method(label):
            m.update(method=_exact_method(label), show_method=True, method_edited=True, label="", label_edited=False)
        if m["type"] == "info" and types.get(a) == "info" and types.get(b) == "info":
            m["inferred"] = True
        if m["type"] == "info" and types.get(a) == "work" and types.get(b) == "work":
            m["type"] = "work"
        edges.append(m)
    out.update(nodes=nodes, edges=edges, schema=2)
    return out

def _exact_method(label):
    """ラベル全体が方法の名前か語の1つと同じなら、その方法のキー（昔の線のラベルを方法に読み替えるとき）"""
    v = vnorm(label).strip().lower()
    for r in map_vocab()["methods"]:
        if v == vnorm(r["label"]).strip().lower() or v in {w.lower() for w in _split_words(r.get("words")) if not w.startswith("~")}:
            return r["key"]
    return ""

PLACEHOLDER_TEXTS = {"使いみち未確認", "使う人が分からない", "（渡すもの未確認）", "新しい作業", "新しい情報"}

def _same_box(a, b):
    if a.get("type") != b.get("type"):
        return False
    if a.get("type") == "info":
        return normalize_text(a.get("text")) == normalize_text(b.get("text"))
    return normalize_text(a.get("text")) == normalize_text(b.get("text")) and normalize_text(a.get("who")) == normalize_text(b.get("who"))

def _ghost_twin(sn, nodes, saved_edges, auto_edges):
    """手で足した箱・昔の箱が、再生成した自動の箱と同じものか。名前とだれが同じ、またはだれが同じ（「課長」は渡す業務の職場名つきでも比べる）で同じ業務から渡されている。
    未登録の業務は、同じ職場で登録された同じ名前の業務"""
    if sn.get("cls") == "pending":
        vocab = map_vocab()
        key = pending_key(_pending_name(sn.get("text")), vocab)
        return next((x for x in nodes if x.get("source") == "auto" and x.get("cls") in TASK_CLASSES and x.get("group") == sn.get("group")
                     and key in {pending_key(v, vocab) for v in [x.get("text")] + task_aliases(_task_id_of(x["id"]))}), None) if key else None
    twin = next((x for x in nodes if x.get("source") == "auto" and _same_box(x, sn)), None)
    if twin is not None or sn.get("type") != "work" or not normalize_text(sn.get("who")):
        return twin
    feeders = {se.get("from") for se in saved_edges if se.get("to") == sn.get("id")}
    tids = {_task_id_of(f) for f in feeders if _task_id_of(f)}
    feeders |= {f"t{i}" for i in tids} | {f"o{i}" for i in tids}
    fed = {e["to"] for e in auto_edges if e["from"] in feeders}
    own = {x.get("who") for x in nodes if x["id"] in {f"t{i}" for i in tids} and x.get("who")}
    whos = {normalize_text(sn.get("who"))} | {normalize_text(recipient_who(sn.get("who"), wp)["label"]) for wp in own}
    return next((x for x in nodes if x.get("source") == "auto" and x.get("type") == "work" and x["id"] in fed
                 and normalize_text(x.get("who")) in whos), None)

def merge_graph(auto, saved):
    result = {
        "schema": 2, "nodes": [], "edges": [], "suppressed_nodes": [], "suppressed_edges": [], "view": None,
        "new_node_ids": [], "needs_layout": True,
    }
    auto_values = lambda e: {"auto_method": e.get("method") or "", "auto_show_method": bool(e.get("show_method")), "auto_handoff": e.get("handoff") or ""}
    if not saved:
        result["nodes"] = [dict(n) for n in auto["nodes"]]
        result["edges"] = [dict(e, **auto_values(e)) for e in auto["edges"]]
        return result

    saved = upgrade_graph(saved)
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
            if sn.get("text_edited") and sn.get("text") and sn["text"].strip() != (m.get("text") or "").strip():
                m["text"] = sn["text"]
                m["text_edited"] = True
                if m["type"] == "info" and not sn.get("medium_edited"):
                    m["medium"] = medium_of(m["text"], m.get("hint"))
            if sn.get("medium_edited") and m["type"] == "info" and vocab_row("media", sn.get("medium")):
                m["medium"], m["medium_edited"] = sn["medium"], True
            if sn.get("who_edited") and m["type"] == "work":
                auto_who, auto_actor = m.get("who"), m.get("actor")
                m["who"], m["who_edited"] = sn.get("who") or "", True
                if sn.get("actor_locked") and vocab_row("actors", sn.get("actor")):
                    m["actor"], m["who_state"], m["actor_locked"] = sn["actor"], "unknown" if sn["actor"] == "unknown" else "ok", True
                elif normalize_text(m["who"]) != normalize_text(auto_who):
                    w = who_of(m["who"], (m.get("group_title") or "") if auto_actor in ("team", "boss") else "")
                    m["actor"], m["who_state"] = w["actor"], w["state"]
        else:
            result["new_node_ids"].append(n["id"])
        nodes.append(m)
    by_id = {n["id"]: n for n in nodes}
    redirect = {}
    saved_edges = saved.get("edges", [])

    def take_twin(sn):
        twin = _ghost_twin(sn, nodes, saved_edges, auto["edges"])
        if twin is None:
            return False
        redirect[sn["id"]] = twin["id"]
        text = (sn.get("text") or "").strip()
        if twin["cls"] == "work" and text and text not in PLACEHOLDER_TEXTS and normalize_text(text) != normalize_text(sn.get("who")) and not twin.get("text_edited"):
            twin["text"], twin["text_edited"] = text, True
        return True

    def keep_ghost(sn):
        m = dict(sn)
        m.setdefault("type", "work")
        m["cls"] = _cls_for(m.get("cls"), m["type"])
        if m["type"] == "info":
            m["medium"] = m.get("medium") if m.get("medium_edited") and vocab_row("media", m.get("medium")) else medium_of(m.get("text"), m.get("hint"))
        elif not m.get("who_edited") and m.get("source") == "manual":
            m["who"], m["actor"], m["who_state"] = "", "", "unconfirmed"
        elif not (m.get("actor_locked") and vocab_row("actors", m.get("actor"))):
            w = who_of(m.get("who"), m.get("group_title") or "")
            m["actor"], m["who_state"], m["actor_locked"] = w["actor"], w["state"], False
        nodes.append(m)
        by_id[m["id"]] = m

    for sn in saved.get("nodes", []):
        if sn.get("source") == "manual" and sn.get("id") and sn["id"] not in by_id and not take_twin(sn):
            keep_ghost(sn)
    manual_ends = {x for se in saved_edges if se.get("source") == "manual" for x in (se.get("from"), se.get("to"))}
    for sn in saved.get("nodes", []):
        sid = sn.get("id")
        if sid in manual_ends and sid not in by_id and sid not in redirect and sid not in sup_n and not re.match(r"^[to]\d+$", sid) and not take_twin(sn):
            keep_ghost(sn)
    node_ids = set(by_id)

    edges = []
    seen = set()
    for e in auto["edges"]:
        if e["id"] in sup_e or e["from"] not in node_ids or e["to"] not in node_ids:
            continue
        m = dict(e, **auto_values(e))
        se = s_edges.get(e["id"])
        if se:
            m["reversed"] = bool(se.get("reversed"))
            if se.get("label_edited"):
                m["label"] = se.get("label") or ""
                m["label_edited"] = True
            # 手で選んだ方法。「なし」はそのまま、消された区分のキーなら自動に戻す
            if se.get("method_edited") and (not se.get("method") or vocab_row("methods", se.get("method"))):
                m["method"] = se.get("method") or ""
                m["show_method"], m["method_edited"] = bool(m["method"]), True
            if se.get("handoff_edited") and se.get("handoff") in ("push", "pull", "none", ""):
                m["handoff"], m["handoff_edited"] = se.get("handoff"), True
        edges.append(m)
        seen.add(e["id"])
    pairs = {(e["from"], e["to"]) for e in edges}
    fed = defaultdict(set)
    for e in edges:
        if _task_id_of(e["from"]):
            fed[_task_id_of(e["from"])].add(e["to"])
    for se in saved.get("edges", []):
        if se.get("source") != "manual" or se.get("id") in seen:
            continue
        a, b = redirect.get(se.get("from"), se.get("from")), redirect.get(se.get("to"), se.get("to"))
        tid = _task_id_of(a)
        if tid and b in by_id and by_id[b].get("source") == "auto" and by_id[b].get("cls") == "work" and b not in fed[tid] and normalize_text(by_id[b].get("who")):
            b = next((x for x in fed[tid] if by_id.get(x, {}).get("cls") == "work" and normalize_text(by_id[x].get("who")) == normalize_text(by_id[b].get("who"))), b)
        if a not in node_ids or b not in node_ids or a == b or (a, b) in pairs:
            continue
        m = dict(se, id=f"{a}->{b}", auto_method="", auto_show_method=False, auto_handoff="", **{"from": a, "to": b})
        if m.get("method") and not vocab_row("methods", m["method"]):
            m["method"], m["show_method"], m["method_edited"] = "", False, False
        ta, tb = by_id[a].get("type"), by_id[b].get("type")
        m["type"] = m.get("type") if m.get("type") in ("work", "info") else ("work" if ta == "work" and tb == "work" else "info")
        if m["type"] == "info" and ta == "work" and tb == "work":
            m["type"] = "work"
        edges.append(m)
        seen.add(m["id"])
        pairs.add((a, b))

    result.update(
        nodes=nodes, edges=edges, suppressed_nodes=sorted(sup_n), suppressed_edges=sorted(sup_e),
        view=saved.get("view"), needs_layout=False,
    )
    return result

_KEY_RX = re.compile(r"^[a-z0-9_]{0,20}$")


def _key(v):
    v = str(v or "")
    return v if _KEY_RX.match(v) else ""

def sanitize_graph(data):
    if not isinstance(data, dict):
        raise ValueError("graph must be an object")
    vocab = map_vocab()
    mask = lambda v: mask_person_names(str(v or ""), vocab)
    nodes, edges = [], []
    for n in (data.get("nodes") or [])[:2000]:
        if not isinstance(n, dict) or not n.get("id"):
            continue
        ntype = _node_type(n)
        nodes.append(
            {
                "id": str(n["id"])[:80], "text": mask(n.get("text"))[:200], "type": ntype,
                "cls": _cls_for(n.get("cls"), ntype),
                "medium": _key(n.get("medium")), "actor": _key(n.get("actor")), "who": mask(n.get("who"))[:100],
                "who_state": n.get("who_state") if n.get("who_state") in ("ok", "unconfirmed", "unknown") else "",
                "hint": _key(n.get("hint")) or None, "group": n.get("group"), "group_title": n.get("group_title"),
                "url": n.get("url"), "source": "manual" if n.get("source") == "manual" else "auto",
                "x": _num(n.get("x")), "y": _num(n.get("y")), "color": (str(n.get("color"))[:20] if n.get("color") else None),
                "text_edited": bool(n.get("text_edited")), "medium_edited": bool(n.get("medium_edited")), "who_edited": bool(n.get("who_edited")),
                "actor_locked": bool(n.get("actor_locked")),
            }
        )
    for e in (data.get("edges") or [])[:4000]:
        if not isinstance(e, dict) or not e.get("from") or not e.get("to"):
            continue
        edges.append(
            {
                "id": str(e.get("id") or f"{e['from']}->{e['to']}")[:170], "from": str(e["from"])[:80], "to": str(e["to"])[:80],
                "type": e.get("type") if e.get("type") in ("work", "info") else "info", "method": _key(e.get("method")),
                "handoff": e.get("handoff") if e.get("handoff") in ("push", "pull", "none") else "", "show_method": bool(e.get("show_method")),
                "inferred": bool(e.get("inferred")), "label": mask(e.get("label"))[:120], "source": "manual" if e.get("source") == "manual" else "auto",
                "reversed": bool(e.get("reversed")), "label_edited": bool(e.get("label_edited")),
                "method_edited": bool(e.get("method_edited")), "handoff_edited": bool(e.get("handoff_edited")),
            }
        )
    view = data.get("view") if isinstance(data.get("view"), dict) else None
    return {
        "schema": 2, "nodes": nodes, "edges": edges,
        "suppressed_nodes": [str(x)[:80] for x in (data.get("suppressed_nodes") or [])][:2000],
        "suppressed_edges": [str(x)[:170] for x in (data.get("suppressed_edges") or [])][:4000],
        "view": {"x": _num(view.get("x")) or 0, "y": _num(view.get("y")) or 0, "zoom": _num(view.get("zoom")) or 1} if view else None,
    }

def _num(v):
    try:
        return float(v) if v is not None else None
    except (TypeError, ValueError):
        return None

# ----- 手修正の業務データへの書き戻し -----
def _task_id_of(node_id):
    if not node_id or len(node_id) < 2 or node_id[0] not in ("t", "o") or not node_id[1:].isdigit():
        return None
    return int(node_id[1:])

def _pending_name(text):
    """未登録の業務の箱の名前から「（2人）」を外す"""
    return re.sub(r"[（(]\d+人[）)]\s*$", "", (text or "").strip()).strip()

def append_text(existing, value):
    existing = (existing or "").strip()
    value = (value or "").strip()
    if not value:
        return existing
    if not existing:
        return value
    if any(normalize_text(p) == normalize_text(value) for p in task_names(existing) + _names(existing)):
        return existing
    return existing + "、" + value

def remove_text(existing, value, exact=False):
    existing = (existing or "").strip()
    value = (value or "").strip()
    if not existing or not value:
        return existing
    parts = task_names(existing)
    kept = [p for p in parts if not (pending_key(p) == pending_key(value) if exact else _mentions(p, value))]
    return existing if len(kept) == len(parts) else "、".join(kept)

def rename_text(existing, old, new):
    parts, out, key = task_names(existing), [], pending_key(old)
    for p in parts:
        p = new if pending_key(p) == key else p
        if p not in out:
            out.append(p)
    return (existing or "").strip() if out == parts else "、".join(out)

MAP_WRITTEN_KEY = "map_written"


def _map_written():
    if has_request_context() and "map_written" in g:
        return g.map_written
    try:
        data = _json_column(Setting.get(MAP_WRITTEN_KEY, ""), {})
    except Exception:
        data = {}
    data = data if isinstance(data, dict) else {}
    if has_request_context():
        g.map_written = data
    return data

def remember_written(kind, task_id, value):
    data = _map_written()
    names = data.setdefault(kind, {}).setdefault(str(task_id), [])
    if value and value not in names:
        names.append(value)
        Setting.set(MAP_WRITTEN_KEY, json.dumps(data, ensure_ascii=False))

def forget_written(kind, value):
    data = _map_written()
    lists = [names for names in (data.get(kind) or {}).values() if isinstance(names, list) and value in names]
    for names in lists:
        names.remove(value)
    if lists:
        Setting.set(MAP_WRITTEN_KEY, json.dumps(data, ensure_ascii=False))

def written_links():
    return {v for ids in (_map_written().get("links") or {}).values() if isinstance(ids, list) for v in ids}

def task_aliases(task_id):
    return list((_map_written().get("names") or {}).get(str(task_id)) or []) if task_id else []

def written_sources(t, info_from=None):
    names = (_map_written().get("sources") or {}).get(str(t.id)) if t is not None and t.id else None
    if not names:
        return set()
    text = t.info_from if info_from is None else info_from
    return {normalize_text(v) for v in names} & {normalize_text(v) for v in task_names(text) + _names(text, dots=True)}

def writeback_candidates(saved, tasks_by_id):
    cands = []
    if not saved:
        return cands
    saved = upgrade_graph(saved)
    nodes = {n["id"]: n for n in saved.get("nodes", []) if n.get("id")}

    def task_of(node_id):
        tid = _task_id_of(node_id)
        return tasks_by_id.get(tid) if tid else None

    def node_of(node_id):
        return nodes.get(node_id) or {}

    def text_of(node_id):
        return (node_of(node_id).get("text") or "").strip()

    for n in saved.get("nodes", []):
        if n.get("text_edited") and str(n.get("id", "")).startswith("t"):
            t = task_of(n["id"])
            new = (n.get("text") or "").strip()
            if t and new and new != t.name:
                key, refs, partial = pending_key(t.name), [], []
                for o in tasks_by_id.values():
                    if o.id == t.id:
                        continue
                    for f in ("prev_task", "next_task"):
                        for p in task_names(getattr(o, f)):
                            if pending_key(p) == key:
                                refs.append((o, f))
                            elif _mentions(p, t.name):
                                partial.append((o, f, p))
                refs = list(dict.fromkeys(refs))
                text = f"業務名「{t.name}」を「{new}」に変更"
                if refs:
                    text += "（" + "、".join(dict.fromkeys(f"「{o.name}」" for o, _ in refs)) + f"の前後の業務に書いた「{t.name}」も直す）"
                cands.append(
                    {
                        "id": f"rename:{t.id}", "task_id": t.id, "text": text,
                        "ops": [{"task_id": t.id, "field": "name", "action": "set", "value": new, "old": t.name}]
                               + [{"task_id": o.id, "field": f, "action": "rename", "old": t.name, "value": new} for o, f in refs],
                    }
                )
                for o, f, p in partial:
                    to = p.replace(t.name, new) if t.name in p else new
                    side = "前の業務" if f == "prev_task" else "後の業務"
                    cands.append(
                        {
                            "id": f"rename:{t.id}:{o.id}:{f}:{nid('r', p)}", "task_id": o.id, "checked": False,
                            "text": f"「{o.name}」の{side}に書いた「{p}」を「{to}」に直す（業務名「{t.name}」と名前の一部だけが同じなので、選んだときだけ直す）",
                            "ops": [{"task_id": o.id, "field": f, "action": "rename", "old": p, "value": to}],
                        }
                    )

    for e in saved.get("edges", []):
        if e.get("source") != "manual":
            continue
        a, b = (e["to"], e["from"]) if e.get("reversed") else (e["from"], e["to"])
        ta, tb = task_of(a), task_of(b)
        na, nb = node_of(a), node_of(b)
        if ta and tb and ta.id != tb.id and a.startswith("t") and b.startswith("t"):
            cands.append(
                {
                    "id": f"seq:{ta.id}:{tb.id}", "task_id": ta.id, "link": f"{a}->{b}",
                    "text": f"「{ta.name}」の後の業務に「{tb.name}」、「{tb.name}」の前の業務に「{ta.name}」を追加",
                    "ops": [
                        {"task_id": ta.id, "field": "next_task", "action": "append", "value": tb.name, "name_of": tb.id},
                        {"task_id": tb.id, "field": "prev_task", "action": "append", "value": ta.name, "name_of": ta.id},
                    ],
                }
            )
        elif tb and not ta and b.startswith("t"):
            src = text_of(a)
            if na.get("type") == "info":
                if not src or src in PLACEHOLDER_TEXTS or src.endswith("（名前未確認）"):
                    continue
                if src in C.SYSTEM_SOURCES:
                    cands.append({"id": f"input:{tb.id}:{a}", "task_id": tb.id, "text": f"「{tb.name}」の入力情報に「{src}」を追加",
                                  "ops": [{"task_id": tb.id, "field": "inputs", "action": "add_list", "value": src}]})
                else:
                    cands.append({"id": f"info_from:{tb.id}:{a}", "task_id": tb.id, "text": f"「{tb.name}」の情報の受け取り元に「{src}」を追加",
                                  "ops": [{"task_id": tb.id, "field": "info_from", "action": "append", "value": src}]})
            elif na.get("cls") == "pending":
                name = _pending_name(src)
                if name:
                    cands.append({"id": f"prev:{tb.id}:{a}", "task_id": tb.id, "link": f"{a}->{b}", "text": f"「{tb.name}」の前の業務に「{name}」を追加",
                                  "ops": [{"task_id": tb.id, "field": "prev_task", "action": "append", "value": name}]})
            elif na.get("cls") != "warn":
                who = (na.get("who") or "").strip()
                content = "" if src in PLACEHOLDER_TEXTS or normalize_text(src) == normalize_text(who) else src
                ops = [{"task_id": tb.id, "field": "prev_task", "action": "append", "value": content}] if content else []
                if who:
                    ops.append({"task_id": tb.id, "field": "actors", "action": "append", "value": who})
                if ops:
                    text = "、".join(([f"前の業務に「{content}」"] if content else []) + ([f"関わる人・部署に「{who}」"] if who else []))
                    cands.append({"id": f"prev:{tb.id}:{a}", "task_id": tb.id, "text": f"「{tb.name}」の{text}を追加", "ops": ops})
        elif ta and not tb and nb.get("type") == "work" and nb.get("cls") != "warn":
            who = (nb.get("who") or "").strip()
            content = text_of(b)
            if nb.get("cls") == "pending":
                name = _pending_name(content)
                if name:
                    cands.append({"id": f"next:{ta.id}:{b}", "task_id": ta.id, "link": f"{a}->{b}", "text": f"「{ta.name}」の後の業務に「{name}」を追加",
                                  "ops": [{"task_id": ta.id, "field": "next_task", "action": "append", "value": name}]})
            elif who:
                actor = nb.get("actor") or ""
                user = actor_user(actor) if nb.get("who_state") == "ok" else ""
                if nb.get("actor_locked") and actor != "unknown" and vocab_row("actors", actor) and recipient_who(who, _own_wp(ta))["actor"] != actor:
                    who = f"{who}（{actor_user(actor) or vocab_label('actors', actor)}）"
                ops = [{"task_id": ta.id, "field": "output_users_note", "action": "append", "value": who}]
                if user and user != "分からない":
                    ops.append({"task_id": ta.id, "field": "output_users", "action": "add_list", "value": user})
                cands.append({"id": f"user:{ta.id}:{b}", "task_id": ta.id, "text": f"「{ta.name}」のアウトプットの利用者に「{who}」を追加", "ops": ops})
            elif content and content not in PLACEHOLDER_TEXTS:
                cands.append({"id": f"next:{ta.id}:{b}", "task_id": ta.id, "text": f"「{ta.name}」の後の業務に「{content}」を追加",
                              "ops": [{"task_id": ta.id, "field": "next_task", "action": "append", "value": content}]})

    for eid in saved.get("suppressed_edges", []):
        if "->" not in str(eid):
            continue
        a, b = str(eid).split("->", 1)
        ta, tb = task_of(a), task_of(b)
        if ta and tb and ta.id != tb.id and a.startswith("t") and b.startswith("t"):
            if mentions_task(ta.next_task, tb.name) or mentions_task(tb.prev_task, ta.name):
                cands.append(
                    {
                        "id": f"unseq:{ta.id}:{tb.id}", "task_id": ta.id, "unlink": f"{a}->{b}",
                        "text": f"「{ta.name}」の後の業務から「{tb.name}」、「{tb.name}」の前の業務から「{ta.name}」を外す",
                        "ops": [
                            {"task_id": ta.id, "field": "next_task", "action": "remove", "value": tb.name, "name_of": tb.id},
                            {"task_id": tb.id, "field": "prev_task", "action": "remove", "value": ta.name, "name_of": ta.id},
                        ],
                    }
                )
        elif ta and a.startswith("t") and node_of(b).get("cls") == "pending" and _pending_name(text_of(b)):
            name = _pending_name(text_of(b))
            cands.append({"id": f"unnext:{ta.id}:{b}", "task_id": ta.id, "unlink": f"{a}->{b}", "text": f"「{ta.name}」の後の業務から「{name}」を外す",
                          "ops": [{"task_id": ta.id, "field": "next_task", "action": "remove", "value": name, "exact": True}]})
        elif tb and b.startswith("t") and node_of(a).get("cls") == "pending" and _pending_name(text_of(a)):
            name = _pending_name(text_of(a))
            cands.append({"id": f"unprev:{tb.id}:{a}", "task_id": tb.id, "unlink": f"{a}->{b}", "text": f"「{tb.name}」の前の業務から「{name}」を外す",
                          "ops": [{"task_id": tb.id, "field": "prev_task", "action": "remove", "value": name, "exact": True}]})

    return [c for c in cands if any(_op_change(op, tasks_by_id) for op in c["ops"])]

WRITEBACK_FIELDS = {"name", "next_task", "prev_task", "inputs", "info_from", "actors", "output_users_note", "output_users"}


def _op_change(op, tasks_by_id):
    """書き戻しの1操作で変わる (業務, 新しい値)。業務がない・対象外の項目・値が変わらないなら None（受け取り元は、書き戻した印がまだなければ変わるものとする）"""
    t = tasks_by_id.get(op["task_id"])
    if t is None or op["field"] not in WRITEBACK_FIELDS:
        return None
    if op.get("name_of") in tasks_by_id:
        op = dict(op, value=tasks_by_id[op["name_of"]].name)
    current = getattr(t, op["field"]) or ""
    new = _new_value(current, op)
    if new != current or (op["field"] == "info_from" and normalize_text(mask_person_names(op["value"])) not in written_sources(t, new)):
        return t, new
    return None

def _new_value(current, op):
    action = op["action"]
    value = op["value"] if action in ("remove", "add_list") else mask_person_names(op["value"])
    if action == "set":
        return value
    if action == "append":
        return append_text(current, value)
    if action == "remove":
        return remove_text(current, value, op.get("exact"))
    if action == "rename":
        return rename_text(current, op["old"], value)
    if action == "add_list":
        items = C.split_list(current)
        if value not in items:
            items.append(value)
        return C.join_list(items)
    return current

def apply_writeback(candidates, selected_ids, tasks_by_id):
    applied, edited = 0, OrderedDict()
    for c in candidates:
        if c["id"] not in selected_ids:
            continue
        for op in c["ops"]:
            change = _op_change(op, tasks_by_id)
            if change:
                t, new = change
                if new != (getattr(t, op["field"]) or ""):
                    setattr(t, op["field"], new)
                    unmark_estimate(t, op["field"])
                    edited[t.id] = t
                applied += 1
                if op["field"] == "info_from":
                    remember_written("sources", op["task_id"], mask_person_names(op["value"]))
                elif op["field"] == "name" and op.get("old"):
                    remember_written("names", op["task_id"], op["old"])
        if c.get("link"):
            remember_written("links", c["task_id"], c["link"])
        if c.get("unlink"):
            forget_written("links", c["unlink"])
    refreshed = [_after_info_edit(t) for t in edited.values()]
    return applied, refreshed.count("ai")

# ----- 共通 -----
def _wp_name(t):
    return t.workplace.name if t.workplace else "職場未設定"

def _wp_key(t):
    return f"wp{t.workplace_id or 0}"

def _own_wp(t):
    return t.workplace.name if t is not None and t.workplace else ""

def task_cls(t):
    if t.status == "完了":
        return "abolished" if t.decision == "廃止" else "done"
    return "hold" if t.status == "保留" else "task"

def _task_node(d, t):
    tid = f"t{t.id}"
    d.node(tid, t.name, task_cls(t), subgraph=(_wp_key(t), _wp_name(t)),
           who={"actor": "team", "label": _own_wp(t), "state": "ok"})
    d.click(tid, url_for("cases.detail", task_id=t.id), "案件詳細を開く")
    return tid

_TASK_NAMES_RX = re.compile(r"[\n、,，]+")


def task_names(text):
    out = []
    for part in _TASK_NAMES_RX.split(text or ""):
        part = part.strip()
        if part and part not in out:
            out.append(part)
    return out

def name_key(name):
    return normalize_text(unicodedata.normalize("NFKC", str(name or "")))

def pending_key(name, vocab=None):
    return name_key(mask_person_names(name, vocab))

def _mentions(part, name):
    np_, nn = normalize_text(part), normalize_text(name)
    return len(np_) >= 2 and len(nn) >= 3 and (nn in np_ or (len(np_) >= 4 and np_ in nn))

def mentions_task(text, name):
    return any(_mentions(p, name) for p in task_names(text))

def _match_tasks(text, tasks, exclude=None, limit=3):
    hits = []
    for t in tasks:
        if t is not exclude and mentions_task(text, t.name):
            hits.append((len(normalize_text(t.name)), t))
    hits.sort(key=lambda x: -x[0])
    return [t for _, t in hits[:limit]]

_SPLIT_NAMES_RX = re.compile(r"[、,，／/\n]+")
MAX_PENDING = 40


def _names(value, dots=False):
    """「、」「／」などで区切った名前。dots=True なら「A・B」も分ける（かっこの中と職場名の「・」は分けない）"""
    names = [v.strip() for v in _SPLIT_NAMES_RX.split(value or "") if v.strip()]
    if not dots:
        return names
    out = []
    for v in names:
        parts = [p.strip() for p in _split_dots(v)]
        if len(parts) > 1 and all(len(p) >= 2 for p in parts) and not any(vnorm(found) in vnorm(v) and "・" in found for found, _ in _workplace_names()):
            out.extend(parts)
        else:
            out.append(v)
    return out

def _split_dots(text):
    parts, cur, depth = [], "", 0
    for ch in text:
        if ch in "（(「『【［[":
            depth += 1
        elif ch in "）)」』】］]":
            depth = max(depth - 1, 0)
        if ch == "・" and depth == 0:
            parts.append(cur)
            cur = ""
        else:
            cur += ch
    return parts + [cur]

def same_name(a, b, min_len=3):
    na, nb = normalize_text(a), normalize_text(b)
    if not na or not nb:
        return False
    if na == nb:
        return True
    short, long_ = (na, nb) if len(na) <= len(nb) else (nb, na)
    return len(short) >= min_len and short in long_

def _parts(name):
    return {p for p in (normalize_text(x) for x in re.split(r"[（()）・]", name or "")) if len(p) >= 2}

def same_thing(a, b):
    pa, pb = _parts(a), _parts(b)
    return same_name(a, b) or bool(pa and pb and (pa <= pb or pb <= pa))

def _bare_medium(part):
    return any(r["key"] not in ("doc", "unknown") and part in {normalize_text(w) for w in [r["label"]] + _split_words(r.get("words")) if not w.startswith("~")}
               for r in map_vocab()["media"])

def same_info_box(a, b):
    na, pa, pb = normalize_text(a), _parts(a), _parts(b)
    if na and na == normalize_text(b):
        return True
    sub = pa if pa <= pb else pb if pb <= pa else None
    return bool(sub) and not all(_bare_medium(p) for p in sub)

def _is_destination(system, action):
    action = action or ""
    for part in re.split(r"[（()）]", system or ""):
        part = part.strip()
        i = action.find(part) if len(part) >= 2 else -1
        if i >= 0 and action[i + len(part):i + len(part) + 1] in ("に", "へ"):
            return True
    return False

def _purpose(t):
    p = (t.purpose or "").strip() if t is not None else ""
    return "" if p in ("", "不明") else p

def _task_outputs(t, steps):
    """業務のできるもの（手順の output と業務の output。空は除く）"""
    return [m for m in [x.get("output") or "" for x in steps] + [t.output or ""] if m.strip()]

def _raw_task_sources(t):
    """業務の取り出し元 [{name, hint, method, shared, fixed}]。手順 → 具体名 → 受け取り元 → 区分の順に見て、同じものは1つにまとめる。
    業務マップから書き戻した受け取り元（fixed）は、言い換えとしてまとめたり省いたりせず必ず描く（同じもの＝same_thing のときだけ1つにする）"""
    items = []
    written = written_sources(t)

    def add(name, hint=None, method="", shared=True, exact=False, fixed=False):
        name = (name or "").strip()
        if not name:
            return
        for it in items:
            same = normalize_text(name) == normalize_text(it["name"])
            if not same and not exact and it["name"] not in C.SYSTEM_SOURCES:
                same = same_thing(name, it["name"]) if fixed or it["fixed"] else same_name(name, it["name"], 2)
            if same:
                it["method"] = it["method"] or method
                return
        items.append({"name": name, "hint": hint, "method": method or "", "shared": shared, "fixed": fixed})

    steps = t.steps
    made = _task_outputs(t, steps)

    def internal(v):
        return any(same_thing(v, m) for m in made)

    for x in steps:
        system, inp, action, to = [(x.get(k) or "").strip() for k in ("system", "input", "action", "to")]
        recv = method_of(x.get("how_from") or "")
        role, m = field_role(system) if system else ("box", "")
        if system and role == "line":
            if not to and inp and not internal(inp):
                add(inp, hint=vocab_row("methods", m)["medium"], method=m)
            continue
        if system and not internal(system) and not _is_destination(system, action):
            add(system, method=m or recv)
        elif inp and not internal(inp) and not to:
            add(inp, method=recv)
    if not items:
        for v in _names(t.inputs_note, dots=True):
            add(v, method=method_of(v))
    named = bool(items)
    cats = t.inputs_list
    specific = (t.info_from or "").strip()
    rows = [INPUT_SOURCE_MAP[c] for c in cats if c in INPUT_SOURCE_MAP and c != "その他"]

    def category_of(v):
        """名前に合う回答の区分（媒体が同じもの。1つしかなければそれ）"""
        medium = medium_of(v)
        hit = next((r for r in rows if r[0] == medium), None)
        return hit or (rows[0] if len(rows) == 1 and medium == "unknown" else None)

    def kinds():
        return {medium_of(it["name"], it["hint"]) for it in items}

    def channel_of(v):
        m = method_of(v)
        return m if m and vocab_row("methods", m)["kind"] == "channel" else ""

    for part in task_names(specific):
        for v in ([part] if normalize_text(part) in written else _names(part, dots=True)):
            row = category_of(v)
            hint = row[0] if row else None
            fixed = normalize_text(v) in written
            if fixed or not named or medium_of(v, hint) not in kinds():
                add(v, hint=hint, method=(row[1] if row and row[1] else "") or channel_of(v), fixed=fixed)
    derived = {_source_category(it["name"], it["method"]) for it in items}
    for c in cats:
        if c in derived:
            continue
        if c in C.SYSTEM_SOURCES:
            add(c, method="system", exact=True)
        elif c in INPUT_SOURCE_MAP and not specific:
            medium, method, label = INPUT_SOURCE_MAP[c]
            if c == "その他" and items:
                continue
            if medium not in kinds():
                add(f"{label}（名前未確認）", hint=medium, method=method, shared=False)
    return items

def task_sources(t):
    items = _raw_task_sources(t)
    steps = t.steps
    if not steps or all(s["shared"] for s in items):
        return items
    tied = {id(s): method for s, _into, method in _tie_case_sources(build_steps_flow(steps, _own_wp(t), t.purpose or ""), steps, items) if not s["shared"]}
    return [s if s["shared"] else dict(s, method=tied[id(s)]) for s in items if s["shared"] or id(s) in tied]

def task_destinations(t):
    """手順で入力・登録している置き場（業務 → 情報）"""
    made = _task_outputs(t, t.steps)
    out = []
    for x in t.steps:
        system = (x.get("system") or "").strip()
        if not system or not _is_destination(system, x.get("action")):
            continue
        role, m = field_role(system)
        if role == "line" or any(same_thing(system, v) for v in made) or any(same_name(system, v["name"], 2) for v in out):
            continue
        out.append({"name": system, "method": m})
    return out

def task_recipients(t):
    """渡す相手 [{name, how, use, who, step}]。手順の to と利用者の具体名を、1人・1部署ずつ。
    だれ（職場名つきの名前）が同じものだけ1つにする（「課長」と「製造1課 課長」は同じ、「検査・品質保証課 課長」は別）"""
    own, out, vocab = _own_wp(t), [], map_vocab()

    def add(name, how="", use="", step=False):
        name = (name or "").strip()
        if not name:
            return
        w = recipient_who(name, own, vocab)
        key = normalize_text(w["label"] or name)
        for r in out:
            if normalize_text(r["who"]["label"] or r["name"]) == key:
                r["how"], r["use"] = r["how"] or how, r["use"] or use
                if r["who"]["state"] == "unconfirmed" and w["state"] != "unconfirmed":
                    r["who"] = w
                return
        out.append({"name": name, "how": how or "", "use": use or "", "who": w, "step": step})

    for x in t.steps:
        to, system = (x.get("to") or "").strip(), (x.get("system") or "").strip()
        if not to:
            continue
        line = field_role(system) if system else ("box", "")
        how = method_of(x.get("how") or "") or (line[1] if line[0] == "line" else "")
        for n in _names(to):
            add(n, how, (x.get("use") or "").strip(), step=True)
    for n in _names(t.output_users_note):
        add(n)
    return out

def split_who_hint(name, vocab=None):
    text = (name or "").strip()
    for row in (vocab or map_vocab())["actors"]:
        for word in dict.fromkeys(w for w in (row.get("label"), row.get("user")) if w):
            for tail in (f"（{word}）", f"({word})"):
                if text.endswith(tail) and len(text) > len(tail):
                    return text[: -len(tail)].strip(), row["key"]
    return text, ""

def recipient_who(name, own_wp="", vocab=None):
    vocab = vocab or map_vocab()
    base, actor = split_who_hint(name, vocab)
    w = who_of(base, own_wp, vocab)
    if actor and actor != w["actor"]:
        w = {"actor": actor, "label": w["label"] or base, "state": "unknown" if actor == "unknown" else "ok"}
    return w

def dept_in_text(text, own_wp="", actor="dept"):
    """業務の目的などの文から、相手の部署・社外の名前を探す（「翌日の残業人員を総務に伝えるため」→ 総務）。なければ空"""
    text = str(text or "")
    if not text.strip():
        return ""
    if actor == "dept":
        for found, w in _workplace_names():
            if wp_name_in(found, text) and w != own_wp:
                return w
    row = vocab_row("actors", actor)
    m = row["_rx"].search(_masked(text, row)) if row and row.get("_rx") else None
    if not m:
        return ""
    word = m.group(0)
    if word in {w[1:] for w in _split_words(row.get("words")) if w.startswith("~")}:
        word = re.search(r"[一-龥々ァ-ヶーA-Za-z0-9]*$", vnorm(text)[:m.start()]).group(0) + word
    return word[:20] if len(word) >= 2 else ""

def category_recipient(u, t):
    """利用者の区分（アンケートの選択肢）しか分からない相手のだれ。他部署・社外は業務の目的の文から名前を探す"""
    own = _own_wp(t)
    w = category_who(u, own)
    if w and w["actor"] in ("dept", "external"):
        found = dept_in_text(_purpose(t), own, w["actor"])
        if found:
            return dict(w, label=found)
    return w

def task_users(t):
    named = task_recipients(t)
    out, covered = [], set()
    for r in named:
        system = r["who"]["state"] == "unconfirmed" and medium_of(r["name"]) == "system"
        out.append(dict(r, system=system, category=False))
        if not system:
            covered.add(actor_user(r["who"]["actor"]) if r["who"]["actor"] else "職場内")
    if named:
        covered.add("分からない")
    for u in t.output_users_list:
        w = category_recipient(u, t) if u != "自分" and u not in covered else None
        if w:
            out.append({"name": u, "how": "", "use": "", "who": w, "step": False, "system": False, "category": True})
    return out

def users_from_names(names, own_wp=""):
    """渡す相手の名前から、利用者の区分（職場内・上司・他部署・顧客・社外）を決める"""
    cats = []
    for n in names:
        w = who_of(n, own_wp)
        if w["state"] == "unconfirmed" and medium_of(n) == "system":
            continue
        c = actor_user(w["actor"]) if w["actor"] else "職場内"
        if c and c not in cats:
            cats.append(c)
    return cats

def category_who(u, own_wp=""):
    if u == "職場内":
        return {"actor": "team", "label": f"{own_wp}の人" if own_wp else "職場内", "state": "ok"}
    if u == "上司":
        return {"actor": "boss", "label": f"{own_wp} 上司".strip(), "state": "ok"}
    if u == "他部署":
        return {"actor": "dept", "label": "他部署", "state": "ok"}
    if u == "顧客・社外":
        return {"actor": "external", "label": "顧客・社外", "state": "ok"}
    if u == "分からない":
        return {"actor": "unknown", "label": "", "state": "unknown"}
    return None

def same_info(a, b):
    """業務をまたいで同じ情報とみなすか。名前の完全一致か、片方が「名前（媒体など）」でもう片方が「名前」"""
    na, nb = normalize_text(a), normalize_text(b)
    if not na or not nb:
        return False
    if na == nb:
        return True
    base = lambda v: normalize_text(re.sub(r"[（(][^（()）]*[）)]\s*$", "", str(v or "")))
    ba, bb = base(a), base(b)
    return (bool(ba) and ba == nb) or (bool(bb) and bb == na) or (len(ba) >= 3 and ba == bb)

def _wp_frame(name):
    w = Workplace.query.filter_by(name=name).first() if name else None
    return (f"wp{w.id}", w.name) if w else None

def recipient_node(d, name, own_wp="", use="", frame=None, who=None, task_id=None):
    """渡す相手の箱。システムなら書き込み先の情報、それ以外は相手の作業（だれ＋使いみち）。(node_id, 'work'|'info')
    task_id を渡すと、業務と答えた名前（区分だけの相手は区分）で箱の ID を決める（区分の語・目的・相手の数を変えても手修正が残るように）"""
    w = who or recipient_who(name, own_wp)
    if w["state"] == "unconfirmed" and medium_of(name) == "system":
        return d.node(nid("s", name), name, "source", medium="system"), "info"
    if w["state"] == "unknown":
        return d.node(nid("p", f"{task_id}:unknown") if task_id is not None else "p_unknown", "使う人が分からない", "warn", who=dict(w, label="")), "work"
    content = use or "使いみち未確認"
    inside = None
    if frame and own_wp and w["actor"] in ("team", "boss") and vnorm(w["label"]).startswith(vnorm(own_wp)):
        inside = frame
    elif frame and w["actor"] == "dept":
        inside = next((_wp_frame(wp) for found, wp in _workplace_names() if wp_name_in(found, w["label"])), None)
    key = f"{w['label']}|{content}" if task_id is None else f"{task_id}:{name}"
    return d.node(nid("p", key), content, "work", subgraph=inside, who=w), "work"

def task_user_nodes(d, t, frame=None):
    """利用者の箱 [{id, type, how}]。具体名を優先し、区分（職場内・上司・他部署・社外・分からない）は具体名で表せないものだけ。箱は業務ごと"""
    own, purpose = _own_wp(t), _purpose(t)
    users = task_users(t)
    named = [u for u in users if not u["category"]]
    cats = [u for u in t.output_users_list if u != "自分"]
    out = []
    for u in users:
        use = (purpose if not named and len(cats) == 1 else "") if u["category"] else (u["use"] or (purpose if len(named) == 1 else ""))
        node, typ = recipient_node(d, u["name"], own, use, frame, who=u["who"], task_id=t.id)
        out.append({"id": node, "type": typ, "how": u["how"]})
    return out

def _add_task_info(d, t, tid, outs=None, dups=None, frame=None):
    """1つの業務の取り出し元・書き込む置き場・できるもの・渡す相手を描く。outs があれば、別の業務のできるものとの一致をつなぐ"""
    for s in task_sources(t):
        prev = None
        if outs and s["shared"] and medium_of(s["name"], s["hint"]) != "system":
            prev = next((oid for o, oid in outs if o is not t and same_info(s["name"], o.output)), None)
        if prev:
            d.info_edge(prev, tid, s["method"], inferred=True)
            continue
        sid = nid("s", s["name"]) if s["shared"] else nid("s", f"{t.id}:{s['name']}")
        node = d.node(sid, s["name"], "source", hint=s["hint"])
        d.info_edge(node, tid, s["method"])
    for dst in task_destinations(t):
        node = d.node(nid("s", dst["name"]), dst["name"], "source")
        d.info_edge(tid, node, dst["method"], info_node=node)
    target = None
    if (t.output or "").strip():
        target = f"o{t.id}"
        d.node(target, t.output, "dupout" if dups and normalize_text(t.output) in dups else "output")
        d.edge(tid, target, "info")
    for u in task_user_nodes(d, t, frame):
        if u["type"] == "info":
            d.edge(tid, u["id"], "info", method="system")
            continue
        if target is None:
            target = d.node(f"o{t.id}", "（渡すもの未確認）", "output", medium="unknown")
            d.edge(tid, target, "info")
        d.info_edge(target, u["id"], u["how"])
    return target or tid

def _task_handoff(a, b, d):
    if a.respondent_id and a.respondent_id == b.respondent_id:
        return "none"
    e = d.edges.get((f"o{a.id}", f"t{b.id}"))
    last = (a.steps[-1].get("action") if a.steps else "") or ""
    first = (b.steps[0].get("action") if b.steps else "") or ""
    return handoff_of(last, first, method=e["method"] if e else "")

# ----- 業務マップ全体（登録済みの業務と、月次アンケートで挙がった未登録の業務） -----
def build_info_flow(tasks):
    d = Diagram("LR")
    dups = {normalize_text(grp["output"]) for grp in duplicate_output_groups()}
    outs = []
    for t in tasks:
        tid = _task_node(d, t)
        if (t.output or "").strip():
            oid = f"o{t.id}"
            d.node(oid, t.output, "dupout" if normalize_text(t.output) in dups else "output")
            d.edge(tid, oid, "info")
            outs.append((t, oid))
    for t in tasks:
        _add_task_info(d, t, f"t{t.id}", outs, dups, frame=(_wp_key(t), _wp_name(t)))
    written = written_links()
    for t in tasks:
        tid = f"t{t.id}"
        for other in _match_tasks(t.next_task, tasks, exclude=t):
            d.edge(tid, f"t{other.id}", "work", handoff=_task_handoff(t, other, d), inferred=f"{tid}->t{other.id}" not in written)
        for other in _match_tasks(t.prev_task, tasks, exclude=t):
            d.edge(f"t{other.id}", tid, "work", handoff=_task_handoff(other, t, d), inferred=f"t{other.id}->{tid}" not in written)
        if (t.output or "").strip():
            for other in tasks:
                if other is not t and same_name(t.output, other.prev_task):
                    d.edge(f"o{t.id}", f"t{other.id}", "info", inferred=True)
    return d

PENDING_MONTHS = 6


def pending_task_names(workplace_id=None):
    """月次アンケート（1-3・3-2）で挙がったが、まだ流れを登録していない業務名。各人の最新の回答（6か月以内・対象外の人を除く）から、回答の職場ごとに。
    名前は1行・「、」ごとに分け（「・」では分けない）、全角・半角と記号をそろえて完全に同じものだけを登録済みとみなす（「月報の作成」と「週報の作成」は別）"""
    vocab = map_vocab()
    since = shift_period(current_period(), 1 - PENDING_MONTHS)
    excluded = {norm_emp_id(e.emp_id).lower() for e in Employee.query.filter(Employee.active.is_(False)).all()}
    latest = {}
    for m in MonthlyResponse.query.filter(MonthlyResponse.period >= since).order_by(MonthlyResponse.period, MonthlyResponse.id).all():
        latest[norm_emp_id(m.emp_id).lower()] = m
    registered = defaultdict(set)
    for t in Task.query.all():
        registered[t.workplace_id] |= {pending_key(v, vocab) for v in [t.name, t.monthly_name] + task_aliases(t.id) if v}
    found = OrderedDict()
    for key, m in latest.items():
        if key in excluded or (workplace_id and m.workplace_id != workplace_id):
            continue
        for raw in task_names(m.answer("task_names")) + task_names(m.answer("wants_tasks")):
            name = mask_person_names(raw, vocab).strip()
            k = name_key(name)
            if len(k) < 2 or k in registered[m.workplace_id]:
                continue
            row = found.setdefault((m.workplace_id, k), {"name": name, "workplace_id": m.workplace_id, "people": set()})
            row["people"].add(key)
    return list(found.values())[:MAX_PENDING]

def add_pending_tasks(d, workplace_id=None, tasks=()):
    names = {w.id: w.name for w in Workplace.query.all()}
    pending, vocab = pending_task_names(workplace_id), map_vocab()
    links = {t.id: ({pending_key(v, vocab) for v in task_names(t.next_task)}, {pending_key(v, vocab) for v in task_names(t.prev_task)}) for t in tasks} if pending else {}
    written = written_links() if pending else set()
    for p in pending:
        n = len(p["people"])
        wp = names.get(p["workplace_id"], "職場未設定")
        node = d.node(nid("u", f"{p['workplace_id']}:{p['name']}"), p["name"] + (f"（{n}人）" if n >= 2 else ""), "pending",
                      subgraph=(f"wp{p['workplace_id'] or 0}", wp), who={"actor": "team", "label": wp, "state": "ok"})
        key = name_key(p["name"])
        for t in tasks:
            if t.workplace_id != p["workplace_id"]:
                continue
            if key in links[t.id][0]:
                d.edge(f"t{t.id}", node, "work", inferred=f"t{t.id}->{node}" not in written)
            if key in links[t.id][1]:
                d.edge(node, f"t{t.id}", "work", inferred=f"{node}->t{t.id}" not in written)
    return d

# ----- 図に添える表: 職場・部署の間の受け渡し、システム・情報源ごとの業務 -----
def handoff_rows(tasks):
    agg = defaultdict(lambda: {"info": 0, "involve": 0})
    for t in tasks:
        wp, own = _wp_name(t), _own_wp(t)
        inputs = t.inputs_list
        if (t.info_from or "").strip():
            for v in _names(t.info_from, dots=True):
                agg[(who_of(v, own)["label"] or v, wp)]["info"] += 1
        elif "人から聞く" in inputs:
            agg[("口頭での聞き取り", wp)]["info"] += 1
        for u in task_users(t):
            w = u["who"]
            if u["system"] or w["actor"] in ("team", "system") or w["state"] == "unconfirmed":
                continue
            agg[(wp, w["label"] or "利用者不明")]["info"] += 1
        for a_ in t.actors_list:
            label = recipient_who(a_, own)["label"] or a_
            if (label, wp) not in agg and (wp, label) not in agg:
                agg[(wp, label)]["involve"] += 1
    rows = [
        {"src": a_, "dst": b_, "info": c["info"], "involve": c["involve"]}
        for (a_, b_), c in agg.items()
    ]
    rows.sort(key=lambda r: (-r["info"], -r["involve"], r["src"], r["dst"]))
    return rows

def source_rows(tasks):
    by = OrderedDict()
    for t in tasks:
        for s in task_sources(t):
            medium = medium_of(s["name"], s["hint"])
            row = by.setdefault(normalize_text(s["name"]), {"name": s["name"], "cls": "system" if medium == "system" else "source", "hint": s["hint"], "tasks": [], "minutes": 0.0})
            if t not in row["tasks"]:
                row["tasks"].append(t)
                row["minutes"] += t.monthly_minutes or 0
    rows = list(by.values())
    rows.sort(key=lambda r: (r["cls"] != "system", -len(r["tasks"]), r["name"]))
    return rows

# ----- 1業務の流れ（案件詳細・聞き取り中） -----
_CIRCLED = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"

def _circled(n):
    return _CIRCLED[n - 1] if 1 <= n <= len(_CIRCLED) else f"({n})"

def build_steps_flow(steps, own_wp="", purpose=""):
    d = Diagram("LR")
    me = {"actor": "team", "label": own_wp or "本人", "state": "ok"}
    purpose = "" if (purpose or "").strip() in ("", "不明") else purpose.strip()
    total = sum(len(_names(x.get("to"))) for x in steps)
    seen, latest, prev = [], None, None

    def find(v):
        return next((nd for text, nd in reversed(seen) if same_name(v, text, 2)), None) if v else None

    def same_box(v, outputs=False):
        return next((nd for text, nd in seen if same_info_box(v, text) or (outputs and d.nodes[nd]["cls"] == "output" and same_thing(v, text))), None) if v else None

    def info(name, hint=None, key=None, cls="source"):
        node = d.node(key or nid("s", name), name, cls, hint=hint)
        if all(nd != node for _, nd in seen):
            seen.append((name, node))
        return node

    def method_medium(method):
        row = vocab_row("methods", method) if method else None
        return row["medium"] if row else None

    for x in steps:
        sid = f"st{x['no']}"
        action = mask_person_names(x.get("action") or "")
        d.node(sid, f"{_circled(x['no'])} {action}", "task", width=14, who=me)
        system, inp, to, how, out = [(x.get(k) or "").strip() for k in ("system", "input", "to", "how", "output")]
        recv = method_of(x.get("how_from") or "")
        role, m = field_role(system) if system else ("box", "")
        src = find(inp)
        if system and role == "line" and not to:
            node = src or (info(inp, hint=method_medium(m)) if inp else info(f"{vocab_label('methods', m)}で受け取る内容", hint=method_medium(m), key=nid("s", f"{x['no']}:{m}")))
            d.info_edge(node, sid, m)
        else:
            sys_node, dest = None, False
            if system and role != "line":
                dest = _is_destination(system, action) or same_info_box(out, system)
                if dest:
                    sys_node = same_box(system) or info(system)
                    d.info_edge(sid, sys_node, m, info_node=sys_node)
                else:
                    sys_node = find(system) or info(system)
                    d.info_edge(sys_node, sid, m or recv)
            if src and src != sys_node:
                d.info_edge(src, sid, recv)
            elif inp and not src and not to and (not system or dest) and not same_name(inp, system, 2):
                node = info(inp, hint=method_medium(recv))
                d.info_edge(node, sid, recv)
        if prev:
            d.edge(prev, sid, "work", handoff="none")
        oid = None
        if out:
            oid = same_box(out, outputs=True)
            if oid is None:
                oid = d.node(f"o{x['no']}", out, "output")
                seen.append((out, oid))
            d.nodes[oid]["cls"] = "output"
            d.edge(sid, oid, "info")
            latest = oid
        if to:
            meth = method_of(how) or (m if role == "line" else "")
            deliver = None
            for name in _names(to):
                target, typ = recipient_node(d, name, own_wp, (x.get("use") or "").strip() or (purpose if total == 1 else ""))
                if typ == "info":
                    d.edge(sid, target, "info", method=meth or "system")
                    continue
                if deliver is None:
                    # 渡すもの: この手順のできるもの → この手順の input → 直前までのできるもの → 「渡すもの未確認」
                    deliver = oid or src
                    if deliver is None and inp:
                        deliver = info(inp, hint=method_medium(recv))
                        d.edge(deliver, sid, "info")
                    deliver = deliver or latest
                    if deliver is None:
                        deliver = d.node(f"o{x['no']}", "（渡すもの未確認）", "output", medium="unknown")
                        d.edge(sid, deliver, "info")
                d.info_edge(deliver, target, meth)
                d.edge(sid, target, "work", handoff=handoff_of(action, how, side="push", method=meth))
        prev = sid
    return d

def _task_ref(name, t, hit=None):
    """前後の業務の名前が登録済みの業務（hit）なら、その業務の職場をだれにする"""
    wp = _own_wp(hit) if hit is not None else ""
    if wp:
        return {"actor": "team" if wp == _own_wp(t) else "dept", "label": wp, "state": "ok"}
    return who_of(name, _own_wp(t))

def _category_step(d, medium, steps):
    for x in steps:
        sid, system, to = f"st{x['no']}", (x.get("system") or "").strip(), (x.get("to") or "").strip()
        role, m = field_role(system) if system else ("box", "")
        recv = method_of(x.get("how_from") or "") or (m if role == "line" and not to else "")
        row = vocab_row("methods", recv) if recv else None
        if row and row["medium"] == medium:
            covered = any(b == sid and e["type"] == "info" and (e["method"] == recv or d.nodes[a]["medium"] == medium) for (a, b), e in d.edges.items())
            return (None, "") if covered else (sid, recv)
    return None, ""

def _tie_case_sources(d, steps, sources):
    texts = [v for x in steps for v in (x.get("system"), x.get("input"), x.get("output")) if v]
    tied = []
    for s_ in sources:
        if any(same_thing(s_["name"], v) for v in texts):
            continue
        into, method = ("st1", s_["method"]) if s_["shared"] else _category_step(d, s_["hint"], steps)
        if not into:
            continue
        node = d.node(nid("s", s_["name"]), s_["name"], "source", hint=s_["hint"])
        d.info_edge(node, into, method)
        tied.append((s_, into, method))
    return tied

def _extra_case_links(d, t, steps):
    """手順に出てこない業務レベルの取り出し元・渡す相手（アンケートや登録で答えたもの・業務マップから書き戻したもの）を、案件の流れ図に足す。手順の箱と同じもの（same_thing）はその箱にまとめる"""
    own, purpose = _own_wp(t), _purpose(t)
    last = f"st{steps[-1]['no']}"
    _tie_case_sources(d, steps, _raw_task_sources(t))
    users = task_users(t)
    step_to = any(u["step"] for u in users)
    extra = [u for u in users if not u["step"]]
    deliver = next((i for i, n in reversed(d.nodes.items()) if n["cls"] == "output"), None)
    for r in extra:
        node, typ = recipient_node(d, r["name"], own, r["use"] or (purpose if len(extra) == 1 and not step_to else ""), who=r["who"])
        if typ == "info":
            d.edge(last, node, "info", method="system")
            continue
        if deliver is None:
            deliver = d.node(f"o{t.id}", t.output or "（渡すもの未確認）", "output", medium=None if (t.output or "").strip() else "unknown")
            d.edge(last, deliver, "info")
        d.info_edge(deliver, node, method_of(r["how"]))
        d.edge(last, node, "work", handoff=handoff_of(steps[-1].get("action") or "", r["how"], side="push", method=method_of(r["how"])))

def build_task_flow(t):
    own = _own_wp(t)
    steps = t.steps
    cls = task_cls(t)
    if steps:
        d = build_steps_flow(steps, own, t.purpose or "")
        _extra_case_links(d, t, steps)
        first, last = "st1", f"st{steps[-1]['no']}"
        if cls != "task":
            for x in steps:
                d.nodes[f"st{x['no']}"]["cls"] = cls
    else:
        d = Diagram("LR")
        first = last = f"t{t.id}"
        d.node(first, t.name, cls, width=16, who={"actor": "team", "label": own, "state": "ok"})
        _add_task_info(d, t, first)
    others = Task.query.filter(Task.id != t.id).all() if t.id else []
    boxes = []

    def neighbor(name, side, hit=None):
        hit = hit if hit is not None else next(iter(_match_tasks(name, others, exclude=t, limit=1)), None)
        w = _task_ref(name, t, hit)
        node = next((nd for text, tid, nd in boxes if (hit is not None and tid == hit.id) or same_thing(name, text)), None)
        if node is None and side == "next":
            named = current_wp_text(name)
            node = next((i for i, n in d.nodes.items() if n["cls"] == "work" and n["text"] not in PLACEHOLDER_TEXTS and same_thing(name, n["text"])), None)
            if node is None and w["state"] == "ok" and w["actor"] != "team":
                node = next((i for i, n in d.nodes.items() if n["cls"] == "work" and n["who_state"] == "ok" and n["who"] and (n["who"] == w["label"] or named.startswith(n["who"]))), None)
            if node is not None:
                who = d.nodes[node]["who"]
                d.nodes[node]["text"] = (named.replace(who, "", 1).strip(" 　の") if who else "") or name
        if node is None:
            node = d.node(f"{side}{len(boxes)}", name, "work", who=w)
        boxes.append((name, hit.id if hit is not None else None, node))
        return node, hit

    for name in task_names(t.prev_task):
        node, hit = neighbor(name, "prev")
        d.edge(node, first, "work", handoff=_task_handoff(hit, t, d) if hit is not None else handoff_of(name, steps[0].get("action") if steps else ""))
    for name in task_names(t.next_task):
        node, hit = neighbor(name, "next")
        d.edge(last, node, "work", handoff=_task_handoff(t, hit, d) if hit is not None else handoff_of(steps[-1].get("action") if steps else "", name))
    every, written = others + [t], written_links()
    for o in others:
        if mentions_task(o.prev_task, t.name) and t in _match_tasks(o.prev_task, every, exclude=o):
            node, _ = neighbor(o.name, "next", o)
            d.edge(last, node, "work", handoff=_task_handoff(t, o, d), inferred=f"t{t.id}->t{o.id}" not in written)
        if mentions_task(o.next_task, t.name) and t in _match_tasks(o.next_task, every, exclude=o):
            node, _ = neighbor(o.name, "prev", o)
            d.edge(node, first, "work", handoff=_task_handoff(o, t, d), inferred=f"t{o.id}->t{t.id}" not in written)
    return d.derive_roles()

# ----- 凡例・閲覧と編集画面に渡すもの -----
def map_legend_html(d):
    vocab = map_vocab()
    nodes = list(d.nodes.values())
    edges = list(d.edges.values())
    parts = ['<div class="map-legend">', '<div class="legend-row lg-grammar"><span class="legend-title">箱と線</span>',
             '<span class="lg-item">' + shape_glyph("rect") + '作業</span>', '<span class="lg-item">' + shape_glyph("doc") + '情報</span>']
    if any(e["type"] == "work" for e in edges):
        parts.append('<span class="lg-item"><i class="lg-line lg-line-work"></i>仕事の流れ</span>')
    if any(e["type"] == "info" for e in edges):
        parts.append('<span class="lg-item"><i class="lg-line lg-line-info"></i>情報の流れ</span>')
    if any(e.get("inferred") for e in edges):
        parts.append('<span class="lg-item"><i class="lg-line lg-line-guess"></i>名前から推定したつながり</span>')
    parts.append("</div>")
    roles = [c for c in CLS_LABELS if any(n["cls"] == c for n in nodes)]
    if roles:
        parts.append('<div class="legend-row"><span class="legend-title">色</span>' + "".join(f'<span class="legend-chip lg-{c}">{CLS_LABELS[c]}</span>' for c in roles) + "</div>")
    media = [r for r in vocab["media"] if any(n["type"] == "info" and n["medium"] == r["key"] for n in nodes)]
    if media:
        parts.append('<div class="kind-legend"><span class="legend-title">媒体</span>' + "".join(
            f'<span class="kind-chip">{shape_glyph(r["shape"])}<i class="bi {r["icon"]} fi m-{r["key"]}"></i>{esc(r["label"])}</span>' for r in media) + "</div>")
    actors = [r for r in vocab["actors"] if any(n["type"] == "work" and n["cls"] in ("work", "warn") and n["actor"] == r["key"] and n["who_state"] != "unconfirmed" for n in nodes)]
    unconfirmed = any(n["type"] == "work" and n["cls"] == "work" and n["who_state"] == "unconfirmed" for n in nodes)
    if actors or unconfirmed:
        chips = "".join(f'<span class="kind-chip"><i class="bi {r["icon"]} fi a-{r["key"]}"></i>{esc(r["label"])}</span>' for r in actors)
        if unconfirmed:
            chips += '<span class="kind-chip"><i class="bi bi-question-circle fi who-unconfirmed"></i>だれ未確認</span>'
        parts.append('<div class="kind-legend"><span class="legend-title">だれ</span>' + chips + "</div>")
    parts.append("</div>")
    return "".join(parts)

def flow_view(d):
    if d is None or d.empty:
        return "", ""
    return d.render(), map_legend_html(d)

def editor_config():
    return {
        "vocab": vocab_public(), "cls_labels": CLS_LABELS, "cls_type": CLS_TYPE, "glyphs": _SHAPE_GLYPHS, "status_marks": STATUS_MARKS,
        "classify_url": url_for("bizmap.api_classify"),
    }

def kind_icon_html(text, hint=None):
    medium = medium_of(text, hint)
    row = vocab_row("media", medium)
    return f'<i class="bi {row["icon"]} fi m-{medium}"></i>' if row else ""

# ===== AI連携（社内 Azure OpenAI の ChatGPT API を Responses API で呼ぶ。なければルールベース） =====
AI_SETTING_LABELS = (("openai_api_url", "URL"), ("openai_api_key", "API キー"), ("openai_model", "モデル名"))


def _api_url():
    return (Setting.get("openai_api_url") or "").strip()

def _api_key():
    return (Setting.get("openai_api_key") or "").strip()

def _model():
    return (Setting.get("openai_model") or "").strip()

def ai_enabled():
    return bool(_api_url()) and bool(_api_key()) and bool(_model())

_LAST_ERROR = {"text": ""}
_NO_TOOLS = set()


def ai_status_text():
    missing = [label for key, label in AI_SETTING_LABELS if not (Setting.get(key) or "").strip()]
    if missing:
        return "ChatGPT API の " + "・".join(missing) + " が未設定（ルールベースで動作中）"
    text = f"ChatGPT API 連携 有効（モデル: {_model()}、Responses API）"
    if _LAST_ERROR["text"]:
        text += f"／直近のエラー: {_LAST_ERROR['text']}"
    return text

# ----- 共通: ChatGPT API（Responses API）呼び出し。標準ライブラリの urllib だけで呼ぶ -----
AI_TOOL_ROUNDS = 10
AI_TOOL_ROWS = 40


def _tools_flat(tool_defs):
    return [
        {"type": "function", "name": t["function"]["name"], "description": t["function"].get("description") or "",
         "parameters": t["function"].get("parameters") or {"type": "object", "properties": {}}}
        for t in (tool_defs or []) if t.get("type") == "function"
    ]


def _ai_request_body(system, items, max_tokens=None, with_reasoning=True, tools=None, tool_choice="auto"):
    body = {
        "model": _model(),
        "store": False,
        "instructions": system,
        "input": items,
        "max_output_tokens": int(max_tokens or conf.get("OPENAI_MAX_OUTPUT_TOKENS", 4000) or 4000),
    }
    if tools:
        body["tools"] = _tools_flat(tools)
        body["tool_choice"] = tool_choice
    effort = str(conf.get("OPENAI_REASONING_EFFORT", "low") or "").strip()
    if effort and with_reasoning:
        body["reasoning"] = {"effort": effort}
    return body


def _ai_post(body):
    req = urllib.request.Request(
        _api_url(),
        data=json.dumps(body, ensure_ascii=False).encode("utf-8"),
        method="POST",
        headers={"Content-Type": "application/json", "Authorization": "Bearer " + _api_key()},
    )
    timeout = float(conf.get("OPENAI_API_TIMEOUT", 90) or 90)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            status, raw = resp.status, resp.read().decode("utf-8", "replace")
    except urllib.error.HTTPError as e:
        try:
            raw = e.read().decode("utf-8", "replace")
        except Exception:
            raw = ""
        return e.code, raw
    try:
        return status, json.loads(raw)
    except ValueError:
        return status, raw


def _parse_response(payload):
    texts, calls, refusal = [], [], None
    for item in payload.get("output") or []:
        typ = item.get("type")
        if typ == "message":
            for part in item.get("content") or []:
                if part.get("type") == "output_text":
                    texts.append(part.get("text") or "")
                elif part.get("type") == "refusal":
                    refusal = part.get("refusal") or "（理由なし）"
        elif typ == "function_call":
            calls.append({"id": item.get("call_id") or item.get("id") or "", "name": item.get("name") or "", "arguments": item.get("arguments") or "{}"})
    content = "".join(texts)
    if not content and not calls and payload.get("output_text"):
        content = str(payload["output_text"])
    return content, calls, refusal


def _ai_turn(system, items, max_tokens=None, tools=None, tool_choice="auto"):
    key = (_api_url(), _model())
    if tools and key in _NO_TOOLS:
        tools = None
    with_reasoning = True
    body = _ai_request_body(system, items, max_tokens, tools=tools, tool_choice=tool_choice)
    try:
        status, payload = _ai_post(body)
        if status == 400 and "reasoning" in body and "reasoning" in str(payload):
            log.warning("ChatGPT API: reasoning が拒否されたため外して再試行します")
            with_reasoning = False
            status, payload = _ai_post(_ai_request_body(system, items, max_tokens, with_reasoning=False, tools=tools, tool_choice=tool_choice))
        if status == 400 and tools:
            log.warning("ChatGPT API: 道具（tools）つきの呼び出しが拒否されたため、道具なしで再試行します")
            plain = [i for i in items if not (isinstance(i, dict) and i.get("type") in ("function_call", "function_call_output"))] if isinstance(items, list) else items
            status, payload = _ai_post(_ai_request_body(system, plain, max_tokens, with_reasoning=with_reasoning))
            if status == 200 and isinstance(payload, dict) and not payload.get("error"):
                _NO_TOOLS.add(key)
    except OSError as e:
        log.warning("ChatGPT API: 接続できません: %s", e)
        _LAST_ERROR["text"] = f"接続できません（{e}）"
        return None, []

    if not isinstance(payload, dict):
        snippet = re.sub(r"\s+", " ", str(payload or ""))[:160]
        log.warning("ChatGPT API: HTTP %s %s", status, snippet)
        if status in (401, 403):
            _LAST_ERROR["text"] = f"HTTP {status}: APIキーが無効か、権限がありません"
        elif status == 404:
            _LAST_ERROR["text"] = "HTTP 404: URL が違います（管理画面の URL が Responses API のものか確認）"
        elif status == 429:
            _LAST_ERROR["text"] = "HTTP 429: 利用量の上限（レート制限）に達しました"
        elif status == 200:
            _LAST_ERROR["text"] = "応答が JSON ではありません（管理画面の URL を確認）"
        else:
            _LAST_ERROR["text"] = f"HTTP {status}: {snippet}"
        return None, []
    if payload.get("error"):
        err = payload["error"]
        msg = err.get("message") if isinstance(err, dict) else str(err)
        log.warning("ChatGPT API: エラー応答: %s", msg)
        _LAST_ERROR["text"] = f"エラー応答: {msg}"
        return None, []
    if payload.get("status") == "incomplete":
        reason = (payload.get("incomplete_details") or {}).get("reason", "")
        log.warning("ChatGPT API: 応答が途中で終わりました（%s）。OPENAI_MAX_OUTPUT_TOKENS を増やしてください", reason)
    content, calls, refusal = _parse_response(payload)
    if not content and not calls and refusal:
        log.warning("ChatGPT API: 応答が拒否されました: %s", refusal)
        _LAST_ERROR["text"] = "応答が拒否されました"
        return None, []
    _LAST_ERROR["text"] = ""
    return content, calls


def _tool_result(ok, content):
    return {"ok": ok, "llm_content": json.dumps(content, ensure_ascii=False, default=str)}


def _dispatch(handlers, name, arguments):
    fn = (handlers or {}).get(name)
    if fn is None:
        return _tool_result(False, {"error": f"未知の道具: {name}"})
    try:
        args = json.loads(arguments or "{}")
    except ValueError:
        return _tool_result(False, {"error": "道具の引数の JSON を読めませんでした"})
    if not isinstance(args, dict):
        args = {}
    try:
        return _tool_result(True, fn(args))
    except Exception as e:
        log.warning("道具 %s の実行に失敗: %s", name, e)
        return _tool_result(False, {"error": f"道具の実行に失敗しました: {e}"})


class _ToolGuard:
    LIMIT = 2

    def __init__(self):
        self.failed = {}
        self.done = set()
        self.repeats = 0

    @staticmethod
    def key(call):
        try:
            args = json.dumps(json.loads(call.get("arguments") or "{}"), ensure_ascii=False, sort_keys=True)
        except ValueError:
            args = call.get("arguments") or ""
        return (call.get("name"), args)

    def before(self, call):
        k = self.key(call)
        if k in self.failed:
            self.repeats += 1
            return _tool_result(False, {"error": "同じ呼び出しです。前回の失敗: " + self.failed[k] + "。引数を変えるか、別の道具を使うか、ここまでの結果で答えてください"})
        if k in self.done:
            self.repeats += 1
            return _tool_result(False, {"error": "同じ呼び出しです。結果はもう出ています。ここまでの結果で答えてください"})
        return None

    def after(self, call, result):
        k = self.key(call)
        if result.get("ok"):
            self.done.add(k)
        else:
            try:
                reason = json.loads(result.get("llm_content") or "{}").get("error") or "失敗"
            except (ValueError, AttributeError):
                reason = "失敗"
            self.failed[k] = str(reason)[:200]

    @property
    def stuck(self):
        return self.repeats >= self.LIMIT


def _call_ai(system, user_text, max_tokens=None, tools=None, handlers=None, tool_choice="auto"):
    if not tools:
        text, _ = _ai_turn(system, user_text, max_tokens)
        return text
    items = list(user_text) if isinstance(user_text, list) else [{"role": "user", "content": str(user_text)}]
    guard = _ToolGuard()
    retried_empty = False
    for round_no in range(C.AI_TOOL_ROUNDS):
        text, calls = _ai_turn(system, items, max_tokens, tools=tools, tool_choice=tool_choice if round_no == 0 else "auto")
        if text is None and not calls:
            return None
        if text:
            items.append({"role": "assistant", "content": text})
        for c in calls:
            items.append({"type": "function_call", "call_id": c["id"], "name": c["name"], "arguments": c["arguments"]})
        if not calls:
            if (text or "").strip():
                return text
            if not retried_empty:
                retried_empty = True
                continue
            break
        for c in calls:
            result = guard.before(c) or _dispatch(handlers, c["name"], c["arguments"])
            guard.after(c, result)
            items.append({"type": "function_call_output", "call_id": c["id"], "output": result["llm_content"]})
        if guard.stuck:
            break
    items.append({"role": "user", "content": "道具の往復はここまでです。ここまでの結果を使って、指示どおりの形で最終回答を返してください。"})
    text, _ = _ai_turn(system, items, max_tokens)
    return text


# ----- 自動化を検討: 聞き取りの会話と、その結果からの登録内容・ECRS の案 -----
INTAKE_PROMPT = """あなたは製造部「PC定型業務ゼロ化プロジェクト」の聞き取り担当です。本人と短い会話をしながら、その人が行っているPCの定型・定期業務の「現状の流れ（手順）」を組み立てます。

進め方:
- 1回の返事で質問は1つだけ。60字以内。やさしい言葉で、相手の答えに一言だけ反応してから聞く（例:「なるほど、まず実績を出すんですね。次に何をしますか？」）。
- 答えが選択で済むときは chips に2〜6個の選択肢を入れる（「分からない」「特にない」も入れる）。自由に書いてほしいときは chips を空にする。
- 聞く順番: ①最初に何をするか → ②次に何をするか（相手が「ここまで」と言うか、手順が {step_limit} 個になるまで繰り返す）→（交代勤務の人だけ）この業務をいつやるか（昼勤／夜勤／昼勤・夜勤の両方）→ ③情報の取り出し元（システム・ファイル・紙・人）と受け取り方（chips に「システムで見る／共有フォルダ／メールで届く／紙で届く／人に聞く」を入れてよい）→ ④できるものと、それを誰にどうやって渡し、何に使われるか（渡す相手・渡し方・使いみちを1つの質問で聞く。例:「できた『日次実績集計表』は、誰にどうやって渡し、何に使われますか？」。chips に渡し方（{methods}）を入れてよい）→ ⑤頻度 → ⑥1回にかかる時間。困っていることは、余裕があれば最後に1問だけ。
- 頻度を聞くときの chips は必ず次のとおり: {frequencies}
- 時間を聞くときの chips は必ず次のとおり: {durations}
- すでに state に入っていることは聞き直さない。相手がまとめて書いたり、手順書や画面の写真を送ったりしたら、読み取れることをすべて state に入れ、足りないことだけ聞く。
- 手順を聞いている間は、毎回「次に何をしますか？」と聞き、chips に「ここまで（手順はこれで全部）」を入れる。相手が「ここまで」「以上」「終わり」「最後」などと答えたら手順は確定とし、それまでは別のことを聞かない。
- 特定の手順について聞くときは、必ず番号と手順名を添える（例:「①の『装置のトラブル状況を確認する』は、何を見て確認しますか？」）。「確認には何を使いますか？」のように、どの手順の話か分からない聞き方はしない。
- 手順は「何を（対象）」「どこから・何で（システム・ファイル・紙・人）」「どうする」が分かる粒度で集める。「情報を集める」「資料を作る」「確認する」のように対象や手段が分からない答えには、次の手順に進む前に1つだけ掘り下げて聞く（例:「どんな情報を、どこから集めますか？」）。答えが返ったら、その手順の action を具体的に書き直す（例:「生産管理システムから前日の実績を集める」）。この掘り下げは手順の途中でもよい。
- 初めて「次に何をしますか？」と聞くときは「全部なら『ここまで』を押してください」と添える。
- 相手が「意味が分からない」「おかしい」と言ったら、謝らずに言い換えて、同じことを聞き直す。

使える道具（function tools）:
- person_history: 聞き取りの最初に1回だけ呼び、本人の登録済み業務・月次アンケートの答え・勤務形態を見てから質問を始める。
- list_known_systems: 手順にシステム・ファイル・帳票・できるものの名前が出たら1回呼び、既存の呼び方があればそれにそろえる（名寄せ）。できるもの（outputs と steps の output）も、同じものが登録済みならその呼び方にそろえる。
- check_duplicate: 業務名が決まったら1回呼ぶ。本人が同じ名前を登録済みなら、そのことを伝えて「前の登録を取り消して続けるか、別の業務として続けるか」を聞く。
- search_similar_tasks: 手順が2つそろったら1回呼び、他の人が同じ成果物を作っていれば「同じ職場で同じものを作っている人がいます。同じ流れですか？」と確かめる。
- 道具の結果は事実として使い、結果に無いことは言わない。同じ道具を同じ引数で2回呼ばない。道具を呼んだ往復でも、最後は必ず上の JSON だけで返す。
- 道具の結果は次のターンでは見えない。分かったこと（例: 二重登録なし、職場の常用システムは生産管理システム、同じ成果物を同じ職場の別の人も作成）は state.notes に短く残し、notes にあることは調べ直さない。渡されていない道具はもう調べ済み。
- 手順を聞く「次に何をしますか？」は回数に数えない。それ以外の質問が全部で {max_q} 回まで。remaining は「手順以外であと何回聞くつもりか」。手順が2つ以上、できるものと渡す相手（使いみち）、頻度、時間がそろったら done を true にし、reply は「この流れで合っていますか？」と確認の一言にする（chips は空）。使いみちを相手が分からないと答えたら unknowns に入れて先へ進む。
- 相手が「違う」「直したい」と言ったら、state を直して done は false のまま確認し直す。
- 写真が読めないときは、そう伝えて文字で教えてほしいと頼む。
- 推測で手順を作らない。相手が言ったことだけを入れる。
- 人の名前（氏名・〇〇さん）は state にも返事にも書かない。職場・役職・部署で書く（例: 製造1課の担当者、品質保証課 課長）。
- 渡す手順（送る・渡す・提出する）の output には、渡したものの名前を入れる。
- sources・outputs と steps の input・output には情報の名前だけを書く（「〜をメール送付」「〜を印刷」のような渡し方・作業は how や action に書く）。
- 引き継ぎの話が出たら、前の勤務からの引き継ぎは情報の取り出し元（sources と steps の input）、次の勤務への引き継ぎは渡す相手（recipients と steps の to）として扱う。

state は次の形（必ず全部のキーを返す。不明は "" か [] か null）:
{
  "title": "<業務名。『〜を〜する』で短く>",
  "steps": [{"no": 1, "action": "<その手順でやること>", "system": "<使うシステム・ファイル・ツール名。なければ空>", "input": "<元になる情報。なければ空>", "output": "<その手順でできるもの。なければ空>", "to": "<渡す相手（職場・役職・部署。氏名は書かない）。なければ空>", "how": "<to がある手順の渡し方（メール・手渡し・共有フォルダなど）。なければ空>", "use": "<渡す相手がそれを何に使うか。分からなければ空>", "how_from": "<この手順で情報を受け取る方法（メールで届く・PHSで聞く・システムで見るなど）。なければ空>"}],
  "sources": ["<情報の取り出し元（システム名・帳票名・人）>"],
  "outputs": ["<最終的にできるもの>"],
  "recipients": ["<最終的に渡す相手。人・部署を1つずつ>"],
  "purpose": "<できたものの使いみち・この業務をする理由（本人の言葉で）。分からなければ空>",
  "frequency": "<頻度。上の選択肢の中の1つ。不明なら null>",
  "duration": "<1回の時間。上の選択肢の中の1つ。不明なら null>",
  "shift": "<交代勤務の人だけ: 昼勤／夜勤／昼勤・夜勤の両方。普通勤務の人は「普通勤務」。不明なら null>",
  "pains": ["<困っていること>"],
  "unknowns": ["<相手が『分からない』と答えたこと>"],
  "notes": ["<道具で分かったこと・確認済みのことの短いメモ>"]
}

出力はこの JSON オブジェクトだけ（前後に説明を付けない）:
{"reply": "<相手への一言と次の質問>", "chips": ["<選択肢>"], "done": false, "remaining": <残り質問数>, "state": {...}}
日本語で。"""

FINALIZE_PROMPT = """あなたは製造部「PC定型業務ゼロ化プロジェクト」の業務分析アシスタントです。聞き取りで組み立てた1件のPCの定型・定期業務の現状の流れ（state と会話記録）を読み、改善チームが使う情報を JSON で返します。

プロジェクトの原則:
- 受けた案件をそのまま自動化しない。まずアウトプットと「なぜやるのか」を確認し、ECRS（排除・結合・入替・簡素化）でムリ・ムダ・ムラを取り除いてから、残る作業を自動化する。
- 目的が不明、できたものの利用者が分からない業務は廃止候補。同じ情報を何度も入力している、同じものを複数人が作っている業務は結合候補。
- AI の推定は改善チームが確認して確定する。断定せず、確認すべき点も示す。
- 勤務が「夜勤」「昼勤・夜勤の両方」の業務は、間接部門が不在の時間帯でも回る対策（自動配信、翌朝の確認に寄せる、引き継ぎ資料の自動作成など）を優先して考える。
- 本人の希望（wish）が「なくしたい」なら目的と利用者の確認を最優先にし排除の案から、「減らしたい」なら結合・入替・簡素化の案を出して残る作業に自動化の手段を付け、「まずは相談したい」なら改善チームがヒアリングで聞くべき質問を用意する。
- 使える道具: search_similar_tasks（登録済みの業務から似たものを探す。related はここで実際に見つかった業務名だけにする）、duplicate_outputs（同じできるものを複数人が作っている組）。結合の案はこの結果に基づいて書き、結果に無い業務名を作らない。道具は各1回まで。
- 「登録済みの内容」があるときは再分析。改善チームが直した内容を前提にして、分析と改善の案を出し直す。
- 人の名前（氏名・〇〇さん）は書かない。職場・役職・部署で書く。

次のキーを持つ JSON オブジェクトだけを返す（前後に説明文を付けない）:
{
  "name": "<業務名。『〜を〜する』で短く>",
  "purpose": "<目的（なぜやるのか・誰が何に使うか）。state.purpose と会話から書く。分からなければ「不明」>",
  "output": "<最終的にできるもの。search_similar_tasks で同じものが見つかったら、その呼び方にそろえる>",
  "output_users": ["<利用者。次の中から: {output_users}>"],
  "output_users_note": "<渡す相手の職場・部署・役職を「、」区切りで1つずつ（例: 課長、生産管理課）。氏名は書かない。なければ空>",
  "inputs": ["<情報のもと。次の中から: {input_sources}>"],
  "inputs_note": "<具体的なシステム名・ファイル名・帳票名を「、」区切りで1つずつ>",
  "info_from": "<情報の受け取り元（部署・システム・帳票）の名前を「、」区切りで。氏名は書かない>",
  "actors": "<関わる部署・役職を「、」区切り。氏名は書かない>",
  "prev_task": "<この業務の前に終わっている必要がある業務。なければ空>",
  "next_task": "<この業務の後に始まる業務。なければ空>",
  "tools": ["<次の中から: {tools}>"],
  "nature": ["<次の中から: {natures}>"],
  "pains": ["<次の中から: {pains}>"],
  "pains_note": "<困りごとの具体。なければ空>",
  "headcount": <同じ作業をしている人数。分からなければ 1>,
  "domain": "<与えられた業務領域リストの中から最も近い1つ>",
  "related": ["<候補リストにある関連業務名>"],
  "summary": "<この業務の要点と位置づけを2文以内で>",
  "ecrs": {"eliminate": "<排除の観点。該当なしなら null>", "combine": "<結合。該当なしなら null>", "rearrange": "<入替。該当なしなら null>", "simplify": "<簡素化。該当なしなら null>"},
  "automation": "<最適化後に残る作業をどう自動化するか（Python / Excel VBA / Power Query / RPA / BI / 生成AI など）。該当なしなら null>",
  "flags": ["<次の中から該当するもの: 目的不明, 廃止候補, 統合候補, 重複入力候補, 自動化候補, 判断業務, 要ヒアリング>"],
  "priority": "<高 | 中 | 低>",
  "questions": ["<改善チームが担当者に確認すべき質問（最大3件）>"],
  "tobe_steps": ["<改善後の流れの案を手順ごとに（最大6行）>"],
  "saving_minutes_per_month": <改善後に減る見込みの月間の分数。分からなければ null>,
  "first_impression": "<本人に見せる前向きな一言（60字以内。例: 転記が3回あるので、入力を1回にまとめられそうです）>",
  "hearing_questions": ["<「まずは相談したい」の人に改善チームが聞く質問（3〜5件）。それ以外は空>"]
}
日本語で、簡潔に。"""


def _intake_prompt():
    methods = "／".join(r["label"] for r in map_vocab()["methods"] if r["kind"] == "channel" and r["key"] != "other")
    return (INTAKE_PROMPT.replace("{frequencies}", "／".join(C.FREQUENCIES)).replace("{durations}", "／".join(C.DURATIONS))
            .replace("{max_q}", str(C.INTAKE_MAX_QUESTIONS)).replace("{step_limit}", str(C.INTAKE_STEP_LIMIT)).replace("{methods}", methods))


def _finalize_prompt():
    return (FINALIZE_PROMPT.replace("{output_users}", "／".join(C.OUTPUT_USERS)).replace("{input_sources}", "／".join(C.INPUT_SOURCES))
            .replace("{tools}", "／".join(C.TOOLS)).replace("{natures}", "／".join(C.NATURES)).replace("{pains}", "／".join(C.PAINS)))


def _intake_state_default():
    return {"title": "", "steps": [], "sources": [], "outputs": [], "recipients": [], "purpose": "", "frequency": None, "duration": None, "shift": None, "pains": [], "unknowns": [], "notes": [], "checked": []}


def _clean_state(data):
    base = _intake_state_default()
    if not isinstance(data, dict):
        return base
    base["title"] = mask_person_names(str(data.get("title") or "").strip())[:200]
    steps = []
    for x in data.get("steps") or []:
        if not isinstance(x, dict):
            continue
        action = str(x.get("action") or "").strip()
        if not action:
            continue
        field = lambda key, n: mask_person_names(str(x.get(key) or "").strip())[:n]
        steps.append({
            "no": len(steps) + 1, "action": mask_person_names(action)[:200], "system": field("system", 100), "input": field("input", 200),
            "output": field("output", 200), "to": field("to", 100), "how": field("how", 60), "use": field("use", 120), "how_from": field("how_from", 60),
        })
        if len(steps) >= C.INTAKE_STEP_LIMIT:
            break
    base["steps"] = steps
    for key in ("sources", "outputs", "recipients", "pains", "unknowns"):
        base[key] = [mask_person_names(str(v).strip())[:100] for v in (data.get(key) or []) if str(v).strip()][:10]
    base["purpose"] = mask_person_names(str(data.get("purpose") or "").strip())[:300]
    base["frequency"] = data.get("frequency") if data.get("frequency") in C.FREQUENCIES else None
    base["duration"] = data.get("duration") if data.get("duration") in C.DURATIONS else None
    base["shift"] = data.get("shift") if data.get("shift") in C.SHIFT_OPTIONS + ["普通勤務"] else None
    base["notes"] = [mask_person_names(str(v).strip())[:200] for v in (data.get("notes") or []) if str(v).strip()][-12:]
    base["checked"] = [str(v)[:120] for v in (data.get("checked") or []) if str(v).strip()][:20]
    return base


def ai_intake(messages, state, image=None, work_pattern="", handlers=None, tools=None, origin=""):
    recent = messages[-24:]
    items = []
    for i, m in enumerate(recent):
        role = "assistant" if m.get("role") == "ai" else "user"
        text = str(m.get("text") or "").strip() or "（空）"
        if role == "user" and image and i == len(recent) - 1:
            items.append({"role": "user", "content": [{"type": "input_text", "text": text}, {"type": "input_image", "image_url": image}]})
        else:
            items.append({"role": role, "content": text})
    if not items:
        start = "聞き取りを始めてください。" + (f"業務名は「{state['title']}」です。" if state.get("title") else "")
        if origin == "monthly":
            start += "この業務は、本人が月次アンケートで「なくしたい・自動化したい」と答えた業務です。最初にそのことへ一言ふれてから、最初の質問をしてください。"
        items.append({"role": "user", "content": start})
    elif image and items[-1]["role"] != "user":
        items.append({"role": "user", "content": [{"type": "input_text", "text": "この写真の手順を読み取ってください"}, {"type": "input_image", "image_url": image}]})
    if work_pattern == "交代勤務":
        pattern_note = "\n\n## 本人の勤務形態\n交代勤務（昼勤・夜勤あり）。手順が2つ以上そろったら一度だけ「この業務はいつやりますか？」と聞き、chips は必ず「昼勤」「夜勤」「昼勤・夜勤の両方」の3つ。答えを state.shift に入れる。"
    elif work_pattern:
        pattern_note = "\n\n## 本人の勤務形態\n普通勤務（日勤のみ）。勤務時間帯は聞かず、state.shift は「普通勤務」にする。"
    else:
        pattern_note = ""
    system = _intake_prompt() + pattern_note + "\n\n## 現在の state\n" + json.dumps(state, ensure_ascii=False)
    data = _extract_json(_call_ai(system, items, tools=INTAKE_TOOLS if tools is None else tools, handlers=handlers))
    if not isinstance(data, dict):
        return None
    try:
        remaining = max(0, int(data.get("remaining")))
    except (TypeError, ValueError):
        remaining = 0
    return {
        "reply": str(data.get("reply") or "").strip()[:400] or "続けてください。",
        "chips": [str(c).strip()[:40] for c in (data.get("chips") or []) if str(c).strip()][:6],
        "done": bool(data.get("done")),
        "remaining": remaining,
        "state": _clean_state(data.get("state") or state),
    }


def ai_finalize(state, messages, wish, domains, candidates, shift="", handlers=None, current=None):
    domain_names = [d.name for d in domains]
    cand = [{"業務名": t.name, "領域": t.domain.name if t.domain else "", "アウトプット": t.output or ""} for t in candidates[:40]]
    log_text = "\n".join(("AI: " if m.get("role") == "ai" else "本人: ") + str(m.get("text") or "") for m in messages[-40:])
    user_text = (
        "## 聞き取りの結果（state）\n" + json.dumps(state, ensure_ascii=False, indent=1)
        + "\n\n## 会話の記録\n" + (log_text or "（なし）")
        + "\n\n## 本人の希望\n" + (wish or "")
        + "\n\n## 勤務\n" + (shift or "不明")
        + ("\n\n## 登録済みの内容（再分析）\n" + json.dumps(current, ensure_ascii=False, indent=1) if current else "")
        + "\n\n## 業務領域リスト（domain はこの中から選ぶ）\n" + json.dumps(domain_names, ensure_ascii=False)
        + "\n\n## 関連候補の業務（related はこの中か、search_similar_tasks で見つかった業務から選ぶ）\n" + json.dumps(cand, ensure_ascii=False, indent=1)
    )
    data = _extract_json(_call_ai(_finalize_prompt(), user_text, tools=FINALIZE_TOOLS, handlers=handlers))
    return data if isinstance(data, dict) else None


# ----- 道具（aiagent_rag_bundle と同じ chat 形式の定義。送るときに平らな形に直す） -----
INTAKE_TOOLS = [
    {"type": "function", "function": {
        "name": "person_history",
        "description": "いま検討している本人の情報。登録済みの業務（名前・月間時間・希望）、月次アンケートで挙げた業務名、勤務形態。聞き取りの最初に1回だけ呼ぶ。",
        "parameters": {"type": "object", "properties": {}}}},
    {"type": "function", "function": {
        "name": "list_known_systems",
        "description": "これまでの登録に出てきたシステム・ファイル・帳票・できるものの名前と件数。手順やできるものの名前を既存の呼び方にそろえる（名寄せ）ために使う。",
        "parameters": {"type": "object", "properties": {"keyword": {"type": "string", "description": "絞り込みの語（任意）"}}}}},
    {"type": "function", "function": {
        "name": "search_similar_tasks",
        "description": "登録済みの業務から、名前・成果物・手順が似ているものを探す。同じ業務の二重登録や、他の人が同じ成果物を作っている重複を見つける。",
        "parameters": {"type": "object", "properties": {"query": {"type": "string", "description": "業務名や成果物、キーワード"}}, "required": ["query"]}}},
    {"type": "function", "function": {
        "name": "check_duplicate",
        "description": "本人が同じ名前の業務をすでに登録していないか確かめる。",
        "parameters": {"type": "object", "properties": {"name": {"type": "string", "description": "業務名"}}, "required": ["name"]}}},
]
FINALIZE_TOOLS = [
    INTAKE_TOOLS[2],
    {"type": "function", "function": {
        "name": "duplicate_outputs",
        "description": "同じできるもの（成果物）を複数の人が作っている組を返す。output を指定するとその成果物だけに絞る。",
        "parameters": {"type": "object", "properties": {"output": {"type": "string", "description": "成果物の名前（任意）"}}}}},
]


def _task_brief(t, steps=True):
    row = {
        "id": t.id, "業務名": t.name, "職場": t.workplace.name if t.workplace else "",
        "できるもの": t.output or "", "情報のもと": t.inputs_note or "・".join(t.inputs_list), "頻度": t.frequency or "",
        "月間時間": round((t.monthly_minutes or 0) / 60, 1), "希望": t.wish or "",
    }
    if steps:
        row["手順"] = [x.get("action", "") for x in t.steps][:4]
    return row


def _similar_tasks(query, exclude_id=None, limit=10):
    words = _content_words(query)
    nq = normalize_text(query)
    scored = []
    for t in Task.query.all():
        if exclude_id and t.id == exclude_id:
            continue
        tw = _content_words(t.name + " " + (t.output or "") + " " + " ".join(x.get("action", "") for x in t.steps))
        score = 2 * len(words & tw)
        nn = normalize_text(t.name)
        if nq and nn and (nq in nn or (len(nn) >= 3 and nn in nq)):
            score += 5
        if nq and normalize_text(t.output) and nq == normalize_text(t.output):
            score += 5
        if score >= 2:
            scored.append((score, t))
    scored.sort(key=lambda x: -x[0])
    return [t for _, t in scored[:limit]]


def _search_similar_tasks(args, exclude_id=None, owner_id=None):
    q = str(args.get("query") or "").strip()
    if not q:
        raise ValueError("query を指定してください")
    rows = []
    for t in _similar_tasks(q, exclude_id=exclude_id)[:C.AI_TOOL_ROWS]:
        row = _task_brief(t)
        if owner_id is not None:
            row["本人の業務"] = t.respondent_id == owner_id
        rows.append(row)
    return {"query": q, "found": len(rows), "業務": rows,
            "note": "「本人の業務」が true で同じ業務名なら二重登録。false（別の人）が同じできるものを作っていれば統合候補として伝える（名前は出さない）。"}


def intake_handlers(r, used=None):
    used = used if used is not None else set()
    def person_history(args):
        latest = monthly_for(r)
        monthly = None
        if latest is not None:
            monthly = {"月度": latest.period, "行っている業務": latest.answer("task_names"), "なくしたい業務": latest.answer("wants_tasks"),
                       "先月比": latest.answer("trend"), "減った理由": latest.answer("decreased_reason"), "増えた業務": latest.answer("new_tasks")}
        return {"職場": r.workplace.name if r.workplace else "", "勤務形態": r.work_pattern or "",
                "登録済みの業務": [_task_brief(t, steps=False) for t in r.tasks][:C.AI_TOOL_ROWS], "月次アンケート": monthly}

    def list_known_systems(args):
        kw = normalize_text(args.get("keyword") or "")
        counter = Counter()
        for t in Task.query.all():
            for x in t.steps:
                system = (x.get("system") or "").strip()
                if system and field_role(system)[0] != "line" and not _is_destination(system, x.get("action")):
                    counter[(system, "取り出し元")] += 1
                if (x.get("output") or "").strip():
                    counter[(x["output"].strip(), "できるもの")] += 1
            for v in _names(t.inputs_note):
                counter[(v, "取り出し元")] += 1
            if (t.info_from or "").strip():
                counter[(t.info_from.strip(), "取り出し元")] += 1
            if (t.output or "").strip():
                counter[(t.output.strip(), "できるもの")] += 1
        rows = [{"名前": k, "種類": kind, "件数": n} for (k, kind), n in counter.most_common() if not kw or kw in normalize_text(k)]
        return {"候補": rows[:C.AI_TOOL_ROWS], "分類の選択肢": C.INPUT_SOURCES}

    def check_duplicate(args):
        name = normalize_text(args.get("name") or "")
        if not name:
            raise ValueError("name を指定してください")
        own = [t for t in r.tasks if normalize_text(t.name) == name]
        others = [t for t in Task.query.filter(Task.respondent_id != r.id).all() if normalize_text(t.name) == name]
        return {"二重登録": bool(own), "本人の登録済み": [_task_brief(t, steps=False) for t in own],
                "他の人の同名の業務": [_task_brief(t, steps=False) for t in others][:C.AI_TOOL_ROWS]}

    def record(name, fn):
        def run(args):
            result = fn(args)
            used.add(name if name != "check_duplicate" else "check_duplicate:" + normalize_text(args.get("name") or ""))
            return result
        return run

    return {
        "person_history": record("person_history", person_history), "list_known_systems": list_known_systems,
        "search_similar_tasks": lambda args: _search_similar_tasks(args, owner_id=r.id), "check_duplicate": record("check_duplicate", check_duplicate),
    }


def intake_tools_for(state):
    checked = set(state.get("checked") or [])
    title_key = "check_duplicate:" + normalize_text(state.get("title") or "")
    out = []
    for t in INTAKE_TOOLS:
        name = t["function"]["name"]
        if name == "person_history" and name in checked:
            continue
        if name == "check_duplicate" and title_key in checked:
            continue
        out.append(t)
    return out


def finalize_handlers(task):
    def duplicate_outputs(args):
        key = normalize_text(args.get("output") or "")
        rows = []
        for grp in duplicate_output_groups():
            if key and normalize_text(grp["output"]) != key:
                continue
            rows.append({"できるもの": grp["output"], "件数": len(grp["tasks"]), "月間時間の合計": round(grp["minutes"] / 60, 1),
                         "業務": [_task_brief(t, steps=False) for t in grp["tasks"]][:10]})
        return {"同じできるものを複数人が作っている組": rows[:10]}

    return {"search_similar_tasks": lambda args: _search_similar_tasks(args, exclude_id=task.id, owner_id=task.respondent_id), "duplicate_outputs": duplicate_outputs}


def _extract_json(text, brackets="{}"):
    """AI の返答から JSON を取り出す（```json の囲みを外し、最初の開き括弧から最後の閉じ括弧まで）。配列なら brackets="[]"。読めなければ None"""
    if not text:
        return None
    text = text.strip()
    fence = re.search(r"```(?:json)?\s*(.*?)```", text, re.S)
    if fence:
        text = fence.group(1).strip()
    start = text.find(brackets[0])
    end = text.rfind(brackets[1])
    if start == -1 or end == -1:
        return None
    try:
        return json.loads(text[start : end + 1])
    except ValueError:
        return None

# ----- 業務（案件）の分析 -----
def task_payload(task):
    return {
        "業務名": task.name,
        "業務領域（推定）": task.domain.name if task.domain else "未選択",
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
        "自分で改善": task.self_improve or C.SELF_IMPROVE_UNCHECKED,
        "補足": task.note or "",
        "勤務": task.shift or "",
        "手順（聞き取り）": [f"{x['no']}. {x['action']}" + (f"［{x['system']}］" if x.get("system") else "") + (f" → {x['output']}" if x.get("output") else "") for x in task.steps],
    }

def task_state(task):
    return _clean_state({
        "title": task.name, "steps": task.steps, "sources": _names(task.inputs_note), "outputs": _names(task.output), "recipients": _names(task.output_users_note),
        "purpose": "" if task.purpose_unknown else task.purpose, "frequency": task.frequency, "duration": task.duration, "shift": task.shift or None,
        "pains": _names(task.pains_note), "unknowns": task.ai_extra_dict.get("unknowns") or [],
    })

# ----- ルールベース（APIキーなしでも同じ形式で動く） -----
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
_DOMAIN_GENERIC_WORDS = {"管理", "業務", "関する", "その他", "全般"}


def _domain_keywords(d):
    words = DOMAIN_KEYWORDS.get(d.name)
    if words is None:
        original = next((name for name, desc in DEFAULT_DOMAINS if desc and desc == (d.description or "").strip()), None)
        words = DOMAIN_KEYWORDS.get(original, [])
    own = re.split(r"[、,，・／/（）()\s]+", d.name + "、" + re.sub(r"(に関する|の)?業務$", "", (d.description or "").strip()))
    return list(dict.fromkeys(list(words) + [w for w in own if len(w) >= 2 and w not in _DOMAIN_GENERIC_WORDS]))

def rule_domain(task, domains):
    texts = [task.name, "" if task.purpose_unknown else task.purpose, task.output, task.note, task.inputs_note]
    texts += [x.get(k) or "" for x in task.steps for k in ("action", "system", "input", "output")]
    text = " ".join(t or "" for t in texts)
    best, best_score = None, 0
    for d in domains:
        score = sum(text.count(w) for w in _domain_keywords(d))
        if score > best_score:
            best, best_score = d, score
    return best

def rule_analyze(task, domains, candidate_tasks):
    found = rule_domain(task, domains)
    best = found.name if found else (task.domain.name if task.domain else None)

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
    if "転記・コピー" in nature or "集計・グラフ作成" in nature or "データ取得" in nature:
        flags.append("自動化候補")
        tool = "Python または Power Query"
        if "Excel" in task.tools_list and "VBA・マクロ" not in task.tools_list:
            tool = "Excel VBA / Power Query"
        if "社内システム" in task.tools_list:
            tool = "RPA またはシステムのデータ出力機能"
        ecrs["automation"] = f"最適化後に残る転記・集計は {tool} での自動化が候補です。"
    missing = task.missing_required
    if missing:
        flags.append("要ヒアリング")
        ecrs["questions"].append(f"必須項目（{'・'.join(missing)}）を担当者と確認してください。")

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

# ----- 次のアクション提案（AI） -----
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
    if not ai_enabled():
        return None
    try:
        text = _call_ai(ACTION_SYSTEM_PROMPT, json.dumps(context, ensure_ascii=False, indent=1))
    except Exception as e:
        log.exception("AI提案の生成に失敗: %s", e)
        return None
    items = _extract_json(text, "[]")
    if items is None:
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

# ===== 初期データ・デモデータ・既存データの移行 =====
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
    "ai_suggestions": "[]",
    "board_show_names": "1",
    "openai_api_url": "",
    "openai_api_key": "",
    "openai_model": "",
}

RENAMED_DEFAULTS = {
    "project_name": ("製造部 繰り返しPC業務ゼロ化ポータル", "製造部 PC定型業務ゼロ化ポータル"),
}

REMOVED_SETTINGS = ("survey_deadline", "contact", "portal_base_url", "monthly_deadline", "mail_subject", "mail_body", "board_message", "target_respondents", "deadline")

def seed_masters():
    changed = _backfill_excluded_at()
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
    for t in Task.query.filter(Task.wish.in_(list(C.LEGACY_WISHES))).all():
        t.wish = C.LEGACY_WISHES[t.wish]
        changed = True
    patterns = {emp_key(e.emp_id): e.work_pattern for e in Employee.query.filter(Employee.work_pattern != "").all()}
    if patterns:
        for model in (Respondent, MonthlyResponse):
            for row in model.query.filter(db.or_(model.work_pattern.is_(None), model.work_pattern == "")).all():
                if emp_key(row.emp_id) in patterns:
                    row.work_pattern = patterns[emp_key(row.emp_id)]
                    changed = True
    changed = _backfill_member_since() or changed
    changed = _merge_employees() or changed
    changed = _sync_members() or changed
    changed = _migrate_monthly_questions() or changed
    changed = _migrate_task_minutes() or changed
    changed = _migrate_self_improve() or changed
    for m in FlowMap.query.filter(FlowMap.kind != "info").all():
        db.session.delete(m)
        changed = True
    if changed:
        db.session.commit()

def _sync_members():
    answered = {}
    for m in MonthlyResponse.query.order_by(MonthlyResponse.created_at, MonthlyResponse.id).all():
        if norm_emp_id(m.emp_id):
            answered[norm_emp_id(m.emp_id).lower()] = (m.emp_id, m.name, m.workplace_id, m.work_pattern)
    for r in Respondent.query.all():
        if norm_emp_id(r.emp_id) and r.tasks:
            answered.setdefault(norm_emp_id(r.emp_id).lower(), (r.emp_id, r.name, r.workplace_id, r.work_pattern))
    existing = {emp_key(e.emp_id): e for e in Employee.query.all()}
    changed = False
    for key, (emp_id, name, wp, pat) in answered.items():
        if key not in existing:
            db.session.add(Employee(emp_id=norm_emp_id(emp_id), name=name or "", workplace_id=wp, work_pattern=pat or "", source="answer"))
            changed = True
    for key, e in existing.items():
        if e.source == "ldap" and key not in answered:
            db.session.delete(e)
            changed = True
    return changed

_SOURCE_RANK = {"manual": 0, "csv": 1, "answer": 2, "ldap": 3}
_RESPONDENT_FIELDS = ("name", "workplace_id", "work_pattern", "c1_wish", "c2_skills", "c3_learn", "c4_hours", "c5_targets", "c6_share", "d1", "d2", "d3", "d4")


def _merge_employees():
    groups, monthly, resps = defaultdict(list), defaultdict(list), defaultdict(list)
    for e in Employee.query.order_by(Employee.id).all():
        groups[emp_key(e.emp_id)].append(e)
    for m in MonthlyResponse.query.order_by(MonthlyResponse.created_at, MonthlyResponse.id).all():
        if emp_key(m.emp_id):
            monthly[emp_key(m.emp_id)].append(m)
    for r in Respondent.query.order_by(Respondent.id).all():
        if emp_key(r.emp_id):
            resps[emp_key(r.emp_id)].append(r)
    changed = False
    for key in dict.fromkeys(list(groups) + list(monthly) + list(resps)):
        rows = sorted(groups.get(key, []), key=lambda e: (e.emp_id != norm_emp_id(e.emp_id), _SOURCE_RANK.get(e.source, 9), e.id))
        used = Counter(norm_emp_id(x.emp_id) for x in monthly.get(key, []) + resps.get(key, []))
        spelling = used.most_common(1)[0][0] if used else norm_emp_id(rows[0].emp_id)
        if rows:
            changed = _merge_employee_rows(rows, spelling) or changed
        changed = _merge_answer_rows(monthly.get(key, []), resps.get(key, []), spelling) or changed
    if changed:
        db.session.flush()
    return changed

def _merge_employee_rows(rows, spelling):
    keep, others = rows[0], rows[1:]
    if others:
        named = sorted((e for e in rows if (e.name or "").strip()), key=lambda e: e.updated_at or datetime.min, reverse=True)
        named.sort(key=lambda e: _SOURCE_RANK.get(e.source, 9) > 1)
        keep.name = named[0].name.strip() if named else ""
        excluded = [e.excluded_at for e in rows if e.active is False]
        keep.active = not excluded
        keep.excluded_at = None if keep.active or None in excluded else min(excluded)
        keep.is_demo = all(e.is_demo for e in rows)
        for field in ("department", "email", "work_pattern", "workplace_id"):
            if not getattr(keep, field):
                setattr(keep, field, next((getattr(e, field) for e in others if getattr(e, field)), getattr(keep, field)))
        keep.created_at = min((e.created_at for e in rows if e.created_at), default=keep.created_at)
        for e in others:
            db.session.delete(e)
        db.session.flush()
    if others or keep.emp_id != spelling:
        keep.emp_id = spelling
        return True
    return False

def _merge_answer_rows(monthly_rows, resp_rows, spelling):
    changed, periods, kept = False, {}, []
    for m in monthly_rows:
        if m.period in periods:
            first = periods[m.period]
            log.warning("社員IDのまとめ: 同じ月度の月次の回答が重なっていたため、後から送られた回答を消しました（月度 %s、社員ID %s、消した回答 #%s %s、残した回答 #%s %s）",
                        m.period, m.emp_id, m.id, m.created_at, first.id, first.created_at)
            db.session.delete(m)
            changed = True
        else:
            periods[m.period] = m
            kept.append(m)
    keep = min(resp_rows, key=lambda r: (norm_emp_id(r.emp_id) != spelling, r.id)) if resp_rows else None
    for r in resp_rows:
        if r is keep:
            continue
        for t in list(r.tasks):
            t.respondent = keep
        for m in MonthlyResponse.query.filter_by(respondent_id=r.id).all():
            m.respondent_id = keep.id
        for field in _RESPONDENT_FIELDS:
            if not getattr(keep, field) and getattr(r, field):
                setattr(keep, field, getattr(r, field))
        keep.part_c_done, keep.part_d_done = bool(keep.part_c_done or r.part_c_done), bool(keep.part_d_done or r.part_d_done)
        keep.is_demo = bool(keep.is_demo and r.is_demo)
        keep.created_at = min((x for x in (keep.created_at, r.created_at) if x), default=None)
        db.session.flush()
        db.session.delete(r)
        changed = True
    if changed:
        db.session.flush()
    for x in kept + ([keep] if keep is not None else []):
        if x.emp_id != spelling:
            x.emp_id = spelling
            changed = True
    return changed

def _backfill_excluded_at():
    from sqlalchemy import text

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")
    result = db.session.execute(text("UPDATE employees SET excluded_at = COALESCE(updated_at, :now) WHERE active = 0 AND excluded_at IS NULL"), {"now": now})
    return bool(result.rowcount)

def _backfill_member_since():
    rows = Employee.query.filter(Employee.created_at.is_(None)).all()
    if not rows:
        return False
    first = {}

    def seen(emp_id, at):
        key = emp_key(emp_id)
        if key and at and (key not in first or at < first[key]):
            first[key] = at

    for emp_id, at in db.session.query(MonthlyResponse.emp_id, MonthlyResponse.created_at).all():
        seen(emp_id, at)
    for emp_id, at in db.session.query(Respondent.emp_id, Task.created_at).join(Task, Task.respondent_id == Respondent.id).all():
        seen(emp_id, at)
    for e in rows:
        e.created_at = min([x for x in (first.get(emp_key(e.emp_id)), e.updated_at) if x], default=None) or datetime.now()
    return True

def _migrate_task_minutes():
    changed = False
    for t in Task.query.filter(db.or_(Task.frequency == "不定期", Task.ai_fields.like("%headcount%"))).all():
        if "headcount" in t.ai_fields_list and not t.verified:
            t.ai_extra = json.dumps(dict(t.ai_extra_dict, headcount_guess=t.headcount), ensure_ascii=False)
            t.headcount = 1
            t.ai_fields = C.join_list([k for k in t.ai_fields_list if k != "headcount"])
            changed = True
        before = t.monthly_minutes
        if t.compute_minutes() != before:
            changed = True
    return changed

def _migrate_self_improve():
    changed = False
    for t in Task.query.filter(Task.self_improve == C.SELF_IMPROVE_NO).all():
        if t.steps:
            t.self_improve = ""
            changed = True
    return changed

def _stance_from_old(ask, self_answer):
    if self_answer == "取り組みたい":
        return C.STANCE_SELF
    if self_answer == "教えてもらえれば取り組みたい":
        return C.STANCE_TOGETHER
    if ask in ("頼みたい", "相談したい"):
        return C.STANCE_ASK
    return ""

_STANCE_OLD_TO_NEW = {"教えてもらいながら自分で取り組みたい": STANCE_TOGETHER, "改善チームに頼みたい": STANCE_ASK, "今は特にない": ""}


_HOURS_WEEK_TO_SHIFT = {"1時間未満": "30分未満", "1〜3時間": "30分未満", "3〜5時間": "30分〜1時間", "5〜10時間": "1〜2時間", "10時間以上": "2〜4時間"}
_PREVIOUS_TEXTS = {
    "has": {"今月、PC定型業務（繰り返しの手作業）はありますか", "PCでおこなう定型業務（繰り返しの手動作業）はありますか"},
    "hours": {"1週間あたり、PC定型業務にかかっている時間は", "1シフト（1出勤）あたり、PCの定型業務を行っている時間は", "1シフト（1出勤）あたり、PCの定型・定期業務を行っている時間は"},
    "trend": {"先月と比べて、PC定型業務は"},
    "wants": {"なくしたい・自動化したい業務はありますか", "なくしたい・自動化したいPC定型業務はありますか"},
    "stance": {"なくしたい業務の改善は、どう進めたいですか"},
    "new_tasks": {"今月、新たに発生したPC定型業務", "増えたPC定型業務名"},
    "wants_tasks": {"あれば、その業務名（1行に1件）"},
}


def _core_in_order(qs):
    core_order = [q["key"] for q in C.MONTHLY_CORE_QUESTIONS]
    slots = [i for i, q in enumerate(qs) if q.get("key") in core_order]
    core_sorted = sorted((qs[i] for i in slots), key=lambda q: core_order.index(q["key"]))
    if [qs[i]["key"] for i in slots] == [q["key"] for q in core_sorted]:
        return qs, False
    qs = list(qs)
    for i, q in zip(slots, core_sorted):
        qs[i] = q
    return qs, True

def _migrate_monthly_questions():
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
                if q.get("type") != base["type"]:
                    q["type"] = base["type"]
                    if base.get("placeholder"):
                        q["placeholder"] = base["placeholder"]
                    touched = True
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
        present = {q.get("key") for q in new_qs}
        for idx, cq in enumerate(C.MONTHLY_CORE_QUESTIONS):
            if cq["key"] in present:
                continue
            before = {c["key"] for c in C.MONTHLY_CORE_QUESTIONS[:idx]}
            pos = max([i for i, q in enumerate(new_qs) if q.get("key") in before], default=-1) + 1
            item = dict(cq, options=list(cq.get("options", [])))
            if cq.get("depends"):
                item["depends"] = dict(cq["depends"])
            new_qs.insert(pos, item)
            present.add(cq["key"])
            touched = True
        new_qs, reordered = _core_in_order(new_qs)
        touched = touched or reordered
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
        if answers.get("has") == "少しある":
            answers["has"] = "ある"
            touched = True
        if answers.get("has") == "ない" and answers.get("hours"):
            answers["hours"] = ""
            touched = True
        if answers.get("trend") == "初回なので分からない":
            answers["trend"] = ""
            touched = True
        if answers.get("stance") in _STANCE_OLD_TO_NEW:
            answers["stance"] = _STANCE_OLD_TO_NEW[answers["stance"]]
            touched = True
        if answers.get("wants") == "ない" and answers.get("stance"):
            answers["stance"] = ""
            touched = True
        if touched:
            r.answers_json = json.dumps(answers, ensure_ascii=False)
            changed = True
    return changed

# ----- デモデータ -----
DEMO_RESPONDENTS = [
    ("田中 一郎", "製造1課", "教えてもらえれば取り組みたい", ["Excel関数"], ["Excel VBA", "RPA"], "1〜2時間", "できる"),
    ("佐藤 花子", "製造1課", "ぜひ取り組みたい", ["Excel関数", "ピボットテーブル", "Excel VBA"], ["Python", "Power BI"], "2〜4時間", "できる"),
    ("鈴木 健", "製造2課", "興味はあるが時間がない", ["Excel関数"], ["Excel VBA"], "1時間未満", "内容による"),
    ("高橋 美咲", "組立課", "教えてもらえれば取り組みたい", ["Excel関数"], ["Power Query", "RPA"], "1〜2時間", "できる"),
    ("伊藤 大輔", "検査・品質保証課", "ぜひ取り組みたい", ["Excel関数", "ピボットテーブル", "Access"], ["Python", "生成AI"], "2〜4時間", "できる"),
    ("渡辺 由美", "生産管理課", "今は考えていない", ["Excel関数", "ピボットテーブル"], [], "1時間未満", "難しい"),
    ("山本 剛", "設備保全課", "教えてもらえれば取り組みたい", ["特になし"], ["Excel関数", "Excel VBA"], "1〜2時間", "内容による"),
    ("中村 翔太", "製造2課", "興味はあるが時間がない", ["Excel関数", "ピボットテーブル"], ["Power BI"], "1時間未満", "内容による"),
]

DEMO_TASKS = [
    dict(name="日報の実績をExcelの集計表に転記する", who="田中 一郎", domain="製造実行・管理", purpose="日次の生産実績を職場の朝会で確認するため", output="日次実績集計表（Excel）",
         sources=["紙の日報"], freq="毎日", dur="15〜30分", head=2, wish="減らしたい", self=True, prev="各ラインの日報記入", next="週報の作成",
         pains=["時間がかかる", "日報の数字を日次表・週報・月報に3回入れている"],
         steps=[("各ラインから手書きの日報を集める", "紙の日報", "前日の生産実績", "", "", ""),
                ("日報の数字を日次実績集計表に転記する", "Excel（日次実績集計表）", "日報", "", "", ""),
                ("不良件数と稼働時間を手で集計する", "Excel（日次実績集計表）", "", "日次実績集計表", "", ""),
                ("朝会用に印刷して課長に渡す", "", "", "印刷した集計表", "課長", "")]),
    dict(name="週報の作成", who="田中 一郎", purpose="週次の生産状況を課長に報告するため", output="週報（Excel）",
         sources=["日次実績集計表（Excel）"], freq="毎週", dur="30分〜1時間", head=1, wish="減らしたい", self=False, prev="日次実績集計表の転記", next="月報の作成",
         pains=["前週の数字を集めるのに時間がかかる"],
         steps=[("日次実績集計表から1週間分の数字を週報に転記する", "Excel（週報）", "日次実績集計表", "", "", ""),
                ("前週比のコメントを書く", "Excel（週報）", "", "週報", "", ""),
                ("課長にメールで送る", "メール", "", "", "課長", "")]),
    dict(name="残業予定表の提出", who="田中 一郎", purpose="翌日の残業人員を総務に伝えるため", output="残業予定表（Excel）",
         sources=["各ラインリーダー（口頭）"], freq="毎日", dur="15分未満", head=1, wish="減らしたい", self=False, prev="", next="",
         pains=["夕方に急ぎで集めることが多い"],
         steps=[("各ラインリーダーに翌日の残業予定を聞く", "口頭", "翌日の残業予定", "", "", ""),
                ("残業予定表に人数と時間を入力する", "Excel（残業予定表）", "", "残業予定表（Excel）", "", ""),
                ("総務にメールで送る", "メール", "", "", "総務", "")]),
    dict(name="生産実績の月報作成", who="佐藤 花子", purpose="月次の生産実績・稼働率を部長会議に報告するため", output="月報（Excel）",
         sources=["生産管理システム", "週報（Excel）"], freq="毎月", dur="2時間以上", head=1, wish="減らしたい", self=True, prev="週報の作成", next="部長会議",
         pains=["数字が多くて時間がかかる", "集計のミスが起きやすい"],
         steps=[("生産管理システムから月間の生産実績をダウンロードする", "生産管理システム", "月間の生産実績", "実績データ（CSV）", "", ""),
                ("週報4週分の数字と合わせて稼働率を集計する", "Excel（月報集計表）", "週報（Excel）", "月報集計表", "", ""),
                ("グラフを付けて月報にまとめる", "Excel（月報）", "月報集計表", "月報（Excel）", "", ""),
                ("部長と生産管理課にメールで送る", "メール", "", "", "部長、生産管理課", "")]),
    dict(name="不良品発生報告のメール配信", who="佐藤 花子", domain="品質管理", purpose="不良発生を関係部署に共有するため", output="不良報告メール",
         sources=["検査員の不良報告書（紙）"], freq="週に数回", dur="15〜30分", head=1, wish="なくしたい", self=True, prev="品質システムへの不良登録", next="",
         pains=["品質システムに入力した内容と同じことをメールにも書いている"],
         steps=[("検査員から紙の不良報告書を受け取る", "紙の不良報告書", "", "", "", ""),
                ("不良の内容を不良報告メールに書き写す", "", "紙の不良報告書", "不良報告メール", "", ""),
                ("生産管理課と課長にメールで送る", "メール", "", "", "生産管理課、課長", "")]),
    dict(name="出勤簿の集計", who="佐藤 花子", domain="人員・労務管理", purpose="月次の勤怠を総務に提出するため", output="勤怠集計表",
         sources=["紙の出勤簿"], freq="毎月", dur="1〜2時間", head=1, wish="減らしたい", self=True, prev="", next="総務での給与計算",
         pains=["人数が多く時間がかかる", "写し間違いが起きやすい"],
         steps=[("紙の出勤簿を全員分回収する", "紙の出勤簿", "", "", "", ""),
                ("出勤日数と残業時間をExcelに転記する", "Excel（勤怠集計表）", "紙の出勤簿", "", "", ""),
                ("月の合計を集計する", "Excel（勤怠集計表）", "", "勤怠集計表", "", ""),
                ("総務に提出する", "", "", "", "総務", "社内便")]),
    dict(name="安全パトロール結果の入力", who="佐藤 花子", purpose="パトロール指摘事項を記録し是正を管理するため", output="指摘事項台帳",
         sources=["紙のチェックシート"], freq="毎週", dur="15〜30分", head=1, wish="減らしたい", self=False, prev="安全パトロール", next="是正確認",
         pains=["特にない"],
         steps=[("パトロールの指摘を書いたチェックシートを受け取る", "紙のチェックシート", "", "", "", ""),
                ("指摘事項台帳に指摘と担当を入力する", "Excel（指摘事項台帳）", "紙のチェックシート", "指摘事項台帳", "", ""),
                ("是正が済んだ項目を更新する", "Excel（指摘事項台帳）", "", "", "", ""),
                ("課長と職場のみんなに知らせる", "", "", "", "課長、職場内", "共有フォルダ")]),
    dict(name="段取り替え時間の記録と集計", who="鈴木 健", purpose="段取り時間の推移を把握し、生産性の改善テーマを決めるため", output="段取り時間推移グラフ",
         sources=["紙の記録用紙"], freq="毎日", dur="15〜30分", head=1, wish="減らしたい", self=True, prev="段取り替え作業", next="改善会議",
         pains=["毎日の入力と集計に時間がかかる"],
         steps=[("段取り替えの開始と終了の時刻を記録用紙に書く", "紙の記録用紙", "", "", "", ""),
                ("記録用紙の時刻をExcelの段取り時間表に入力する", "Excel（段取り時間表）", "紙の記録用紙", "段取り時間表", "", ""),
                ("段取り工数の推移グラフを更新する", "Excel（段取り時間表）", "", "段取り時間推移グラフ", "", ""),
                ("改善会議で課長に見せる", "", "", "", "課長", "会議")]),
    dict(name="設備日常点検表のファイリング確認", who="鈴木 健", purpose="点検漏れがないか確認するため", output="点検実施チェック表",
         sources=["紙の日常点検表"], freq="毎日", dur="15分未満", head=1, wish="まずは相談したい", self=False, prev="日常点検", next="",
         pains=["特にない"],
         steps=[("各設備の日常点検表を集める", "紙の日常点検表", "", "", "", ""),
                ("点検漏れがないか見比べて確認する", "", "紙の日常点検表", "", "", ""),
                ("点検実施チェック表に済の印を付ける", "Excel（点検実施チェック表）", "", "点検実施チェック表", "", ""),
                ("点検漏れがあれば班長に伝える", "", "", "", "班長", "口頭")]),
    dict(name="原材料の払出し依頼票の作成", who="鈴木 健", domain="生産計画・管理", purpose="翌日分の材料を倉庫に依頼するため", output="払出し依頼票",
         sources=["週次生産計画表（Excel）", "生産管理システム"], freq="毎日", dur="15〜30分", head=2, wish="なくしたい", self=True, prev="週次生産計画の受領", next="倉庫での払出し",
         pains=["生産計画の数字を手で入れ直している"],
         steps=[("週次生産計画表から翌日分の品番と数量を確認する", "週次生産計画表（Excel）", "", "", "", ""),
                ("生産管理システムで品番ごとの材料を調べる", "生産管理システム", "", "", "", ""),
                ("払出し依頼票に品番と数量を入力して印刷する", "Excel（払出し依頼票）", "", "払出し依頼票", "", ""),
                ("倉庫に持っていく", "", "", "", "倉庫", "手渡し")]),
    dict(name="課内会議資料の作成", who="中村 翔太", purpose="週次の課内会議で実績と課題を共有するため", output="会議資料（PowerPoint）",
         sources=["週報（Excel）"], freq="毎週", dur="1〜2時間", head=1, wish="減らしたい", self=False, prev="週報の作成", next="課内会議",
         pains=["資料づくりに時間がかかる"],
         steps=[("週報から実績の数字を拾う", "週報（Excel）", "", "", "", ""),
                ("実績のグラフをExcelで作る", "Excel（実績グラフ）", "", "実績グラフ", "", ""),
                ("課題とあわせて会議資料にまとめる", "PowerPoint（会議資料）", "実績グラフ", "会議資料（PowerPoint）", "", ""),
                ("課内会議で全員に配る", "", "", "", "課内メンバー", "会議")]),
    dict(name="生産性指標の報告（部へ）", who="中村 翔太", purpose="部の生産性会議に課の指標を提出するため", output="生産性指標シート",
         sources=["月報（Excel）"], freq="毎月", dur="30分〜1時間", head=1, wish="なくしたい", self=False, prev="月報の作成", next="部の生産性会議",
         pains=["月報と同じ数字を別のフォーマットに入れ直している"],
         steps=[("月報から課の数字を拾う", "月報（Excel）", "", "", "", ""),
                ("部の生産性指標シートに入れ直す", "Excel（生産性指標シート）", "", "生産性指標シート", "", ""),
                ("部長にメールで提出する", "メール", "", "", "部長", "")]),
    dict(name="組立実績の日次入力（生産管理システム）", who="高橋 美咲", domain="製造実行・管理", purpose="進捗管理と在庫引き落としのため", output="システム上の実績データ",
         sources=["紙の実績票"], freq="毎日", dur="30分〜1時間", head=1, wish="減らしたい", self=True, prev="組立作業の日報記入", next="生産管理課の進捗確認",
         pains=["手書きの数字が読みにくく時間がかかる", "入力ミスが起きやすい"],
         steps=[("組立作業者の手書き実績票を回収する", "紙の実績票", "組立数・不良数", "", "", ""),
                ("生産管理システムに品番ごとの実績を入力する", "生産管理システム", "実績票", "システム上の実績データ", "生産管理課", ""),
                ("入力後に紙の実績票をファイリングする", "", "", "", "", "")]),
    dict(name="外注先への納期確認メール", who="高橋 美咲", purpose="外注部品の納期遅れを早めに把握するため", output="納期回答一覧",
         sources=["発注一覧（Excel）", "外注先からの回答メール"], freq="週に数回", dur="30分〜1時間", head=1, wish="減らしたい", self=True, prev="発注", next="組立計画の調整",
         pains=["急ぎの問い合わせで割り込まれる", "人によって聞き方が違う"],
         steps=[("発注一覧から納期が近い部品を探す", "発注一覧（Excel）", "", "", "", ""),
                ("外注先に納期確認のメールを送る", "メール", "", "", "", ""),
                ("届いた回答を納期回答一覧に入力する", "Excel（納期回答一覧）", "外注先からの回答メール", "納期回答一覧", "", ""),
                ("遅れそうな部品を班長に伝える", "", "", "", "班長", "口頭")]),
    dict(name="作業手順書の改訂管理台帳の更新", who="高橋 美咲", purpose="手順書の版数を管理するため", output="改訂管理台帳",
         sources=["手順書を改訂した人（口頭）"], freq="不定期", dur="15〜30分", head=1, wish="なくしたい", self=False, prev="手順書の改訂", next="",
         pains=["文書管理システムでも版数が分かるので二重管理になっている"],
         steps=[("手順書を改訂した人から改訂の内容を聞く", "口頭", "改訂の内容", "", "", ""),
                ("改訂管理台帳に版数と改訂日を入力する", "Excel（改訂管理台帳）", "", "改訂管理台帳", "", ""),
                ("職場のみんなに改訂を知らせる", "", "", "", "職場内", "掲示")]),
    dict(name="日次実績のグラフ更新（掲示用）", who="高橋 美咲", purpose="", output="掲示板の実績グラフ",
         sources=["日次実績集計表（Excel）"], freq="毎日", dur="15分未満", head=1, wish="まずは相談したい", self=False, prev="日次実績集計表", next="",
         pains=["前任者から引き継いだが、誰が見ているか分からない"], unknowns=["誰が見ているか（渡す相手・使いみち）"],
         steps=[("日次実績集計表から前日の出来高を写す", "日次実績集計表（Excel）", "", "", "", ""),
                ("掲示用の実績グラフをExcelで更新して印刷する", "", "", "掲示板の実績グラフ", "", ""),
                ("掲示板に貼り出す", "", "", "", "", "")]),
    dict(name="検査成績書の作成", who="伊藤 大輔", domain="品質管理", purpose="出荷時に顧客へ品質を証明するため", output="検査成績書（PDF）",
         sources=["品質システム"], recipients=["顧客"], freq="毎日", dur="30分〜1時間", head=2, wish="減らしたい", self=True, prev="出荷検査", next="出荷",
         pains=["測定値が多くて時間がかかる", "写し間違いが起きやすい"],
         steps=[("品質システムから出荷ロットの検査データを出力する", "品質システム", "出荷ロットの検査データ", "検査データ（CSV）", "", ""),
                ("成績書の様式に測定値を転記する", "Excel（検査成績書の様式）", "検査データ（CSV）", "", "", ""),
                ("PDFにして営業にメールで送る", "メール", "", "検査成績書（PDF）", "営業", "")]),
    dict(name="不良データの集計とパレート図作成", who="伊藤 大輔", purpose="月次品質会議で重点不良を決めるため", output="不良パレート図",
         sources=["品質システム"], freq="毎月", dur="1〜2時間", head=1, wish="減らしたい", self=True, prev="不良登録", next="品質会議",
         pains=["グラフを作るのに時間がかかる"],
         steps=[("品質システムから1か月分の不良データをダウンロードする", "品質システム", "不良データ", "不良データ（CSV）", "", ""),
                ("不良項目ごとに件数を集計する", "Excel（不良集計表）", "不良データ（CSV）", "不良集計表", "", ""),
                ("パレート図を作る", "Excel（不良集計表）", "不良集計表", "不良パレート図", "", ""),
                ("品質会議で課長と製造各課に見せる", "", "", "", "課長、製造各課", "会議")]),
    dict(name="測定器の校正期限チェック", who="伊藤 大輔", purpose="校正切れの測定器を使わないため", output="校正期限一覧",
         sources=["校正台帳（Excel）"], freq="毎週", dur="15分未満", head=1, wish="減らしたい", self=True, prev="", next="校正依頼",
         pains=["特にない"],
         steps=[("校正台帳から期限が近い測定器を探す", "校正台帳（Excel）", "", "", "", ""),
                ("校正期限一覧にまとめる", "Excel（校正期限一覧）", "", "校正期限一覧", "", ""),
                ("各職場に期限が近いことをメールで知らせる", "メール", "", "", "各職場", "")]),
    dict(name="クレーム対応の進捗報告", who="伊藤 大輔", purpose="顧客クレームの対応状況を上司・営業に共有するため", output="進捗報告メール",
         sources=["顧客からのクレームメール", "担当者からの聞き取り"], freq="週に数回", dur="15〜30分", head=1, wish="減らしたい", self=False, prev="クレーム受付", next="顧客への回答",
         pains=["急ぎの問い合わせで作業が中断する"],
         steps=[("顧客からのクレームメールで内容を確かめる", "メール", "顧客からのクレームメール", "", "", ""),
                ("担当者から対応の状況を聞く", "口頭", "対応の状況", "", "", ""),
                ("クレーム管理表を更新する", "Excel（クレーム管理表）", "", "クレーム管理表", "", ""),
                ("進捗報告メールを書いて課長と営業に送る", "メール", "", "進捗報告メール", "課長、営業", "")]),
    dict(name="週次生産計画の作成と配布", who="渡辺 由美", domain="生産計画・管理", purpose="各課に翌週の生産順序を指示するため", output="週次生産計画表",
         sources=["受注一覧（Excel）", "生産管理システム"], freq="毎週", dur="2時間以上", head=1, wish="減らしたい", self=False, prev="受注一覧の更新", next="各課の段取り・材料手配",
         pains=["計画の作成に時間がかかる", "急な変更で割り込まれる"],
         steps=[("受注一覧から翌週分の受注を抜き出す", "受注一覧（Excel）", "", "", "", ""),
                ("生産管理システムで在庫と設備の空きを確認する", "生産管理システム", "", "", "", ""),
                ("生産順序を決めて週次生産計画表を作る", "Excel（週次生産計画表）", "", "週次生産計画表", "", ""),
                ("製造各課にメールで配布する", "メール", "", "", "製造各課", "")]),
    dict(name="受注情報のExcel転記", who="渡辺 由美", domain="生産計画・管理", purpose="営業からの受注を計画表に反映するため", output="受注一覧（Excel）",
         sources=["営業からの受注メール"], recipients=["生産管理課の計画担当"], freq="毎日", dur="30分〜1時間", head=1, wish="減らしたい", self=False, prev="営業の受注入力", next="週次生産計画の作成",
         pains=["営業が販売システムに入れた情報を、メールで受けて転記している", "入力ミスが起きやすい"],
         steps=[("営業から届いた受注メールを開く", "メール", "受注メール", "", "", ""),
                ("品番・数量・納期を受注一覧に転記する", "Excel（受注一覧）", "受注メール", "受注一覧（Excel）", "", "")]),
    dict(name="在庫報告書の作成", who="渡辺 由美", purpose="月末在庫を経理に報告するため", output="在庫報告書",
         sources=["生産管理システム"], freq="毎月", dur="1〜2時間", head=1, wish="減らしたい", self=False, prev="月末棚卸", next="経理の決算処理",
         pains=["月末に作業が集中して時間がかかる"],
         steps=[("生産管理システムから月末の在庫データを出力する", "生産管理システム", "月末の在庫データ", "在庫データ（CSV）", "", ""),
                ("品目別に在庫金額を集計する", "Excel（在庫報告書）", "在庫データ（CSV）", "", "", ""),
                ("在庫報告書にまとめて経理に提出する", "", "", "在庫報告書", "経理", "メール")]),
    dict(name="設備故障記録の台帳入力", who="山本 剛", purpose="故障履歴を残し予防保全に活かすため", output="故障履歴台帳",
         sources=["オペレーターからの故障連絡（口頭）"], freq="週に数回", dur="15〜30分", head=1, wish="減らしたい", self=True, prev="故障対応", next="月次保全報告書の作成",
         pains=["人によって書き方が違う"],
         steps=[("オペレーターから故障の連絡を受ける", "PHS", "故障の内容", "", "", ""),
                ("修理の内容を作業メモに書く", "紙の作業メモ", "", "作業メモ", "", ""),
                ("作業メモを見て故障履歴台帳に入力する", "Excel（故障履歴台帳）", "作業メモ", "故障履歴台帳", "", ""),
                ("台帳を課長と共有する", "", "", "", "課長", "共有フォルダ")]),
    dict(name="保全部品の発注依頼書作成", who="山本 剛", purpose="部品を購買に発注してもらうため", output="発注依頼書",
         sources=["保全担当からの聞き取り", "部品台帳（Excel）"], freq="週に数回", dur="15〜30分", head=1, wish="減らしたい", self=False, prev="故障対応・点検", next="購買での発注",
         pains=["部品台帳と同じ内容を発注依頼書に書き写している"],
         steps=[("保全担当から必要な部品を聞く", "口頭", "必要な部品", "", "", ""),
                ("部品台帳で品番と価格を調べる", "部品台帳（Excel）", "", "", "", ""),
                ("発注依頼書に品番・数量を書き写す", "Excel（発注依頼書）", "", "発注依頼書", "", ""),
                ("購買にメールで送る", "メール", "", "", "購買", "")]),
    dict(name="月次保全報告書の作成", who="山本 剛", purpose="設備稼働率と故障件数を部長に報告するため", output="月次保全報告書",
         sources=["故障履歴台帳（Excel）"], freq="毎月", dur="1〜2時間", head=1, wish="減らしたい", self=True, prev="故障履歴台帳の入力", next="部長への報告",
         pains=["グラフ作りに時間がかかる"],
         steps=[("故障履歴台帳から1か月分の故障件数を数える", "故障履歴台帳（Excel）", "", "", "", ""),
                ("設備稼働率を計算してグラフにする", "Excel（月次保全報告書）", "", "", "", ""),
                ("月次保全報告書にまとめて部長に提出する", "", "", "月次保全報告書", "部長", "メール")]),
]

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


DEMO_INTAKE = {
    "日報の実績をExcelの集計表に転記する": [
        {"role": "ai", "text": "「日報の実績をExcelの集計表に転記する」ですね。最初に何をしますか？"},
        {"role": "user", "text": "各ラインから手書きの日報を集める"},
        {"role": "ai", "text": "なるほど。次にその日報をどうしますか？"},
        {"role": "user", "text": "日次実績集計表（Excel）に数字を転記して、不良件数と稼働時間を手で集計する"},
        {"role": "ai", "text": "集計表ができたら、誰に渡しますか？"},
        {"role": "user", "text": "印刷して課長に。朝会で使う"},
        {"role": "ai", "text": "これはどのくらいの頻度ですか？"},
        {"role": "user", "text": "毎日"},
        {"role": "ai", "text": "1回にかかる時間はどのくらいですか？"},
        {"role": "user", "text": "15〜30分"},
        {"role": "ai", "text": "この流れで合っていますか？"},
        {"role": "user", "text": "はい、合っています"},
    ],
}
DEMO_AI_EXTRA = {
    "日報の実績をExcelの集計表に転記する": {
        "tobe_steps": ["各ラインでタブレットから日報を直接入力する", "集計表は入力データから自動で更新する", "朝会は画面で確認し、印刷をやめる"],
        "saving_minutes_per_month": 300,
        "first_impression": "転記と手集計が重なっているので、入力を1回にまとめられそうです",
        "hearing_questions": [],
    },
}

DEMO_MASTERS_KEY = "demo_masters"


def _demo_master(cache, model, name, sort, created):
    obj = cache.get(name)
    if obj is None:
        obj = model(name=name, sort=sort)
        db.session.add(obj)
        db.session.flush()
        cache[name] = obj
        created.append(obj.id)
    return obj

def _demo_state(row):
    steps = [{"no": i, "action": a, "system": s, "input": inp, "output": o, "to": to, "how": how} for i, (a, s, inp, o, to, how) in enumerate(row["steps"], 1)]
    return _clean_state({
        "title": row["name"], "steps": steps, "sources": row["sources"], "outputs": [row["output"]], "recipients": row.get("recipients") or [n for x in steps for n in _names(x["to"])],
        "purpose": row["purpose"], "frequency": row["freq"], "duration": row["dur"], "pains": row["pains"], "unknowns": row.get("unknowns") or [],
    })

def demo_loaded():
    return any(db.session.query(m.id).filter_by(is_demo=True).first() is not None for m in (Task, Respondent, MonthlyResponse, Employee))

def seed_demo():
    clear_demo()
    wp_by_name = {w.name: w for w in Workplace.query.all()}
    dom_by_name = {d.name: d for d in Domain.query.all()}
    created = {"workplaces": [], "domains": [], "surveys": []}
    base = datetime.now() - timedelta(days=20)

    resp_by_name = {}
    for i, (name, wp, c1, skills, learn, c4, c6) in enumerate(DEMO_RESPONDENTS):
        workplace = _demo_master(wp_by_name, Workplace, wp, 100 + i, created["workplaces"])
        emp_id = f"DEMO-{i + 1:03d}"
        r = Respondent.query.filter(db.func.lower(Respondent.emp_id) == emp_id.lower()).first()
        if r is None:
            d = DEMO_D.get(name, ("", "", "", ""))
            r = Respondent(
                emp_id=emp_id,
                name=name, workplace=workplace, is_demo=True, work_pattern=("交代勤務" if i % 2 else "普通勤務"),
                c1_wish=c1, c2_skills=C.join_list(skills), c3_learn=C.join_list(learn), c4_hours=c4, c6_share=c6,
                part_c_done=True, part_d_done=any(d),
                d1=d[0], d2=d[1], d3=d[2], d4=d[3],
                created_at=base + timedelta(days=i),
            )
            db.session.add(r)
        resp_by_name[name] = r
    db.session.flush()

    made = []
    for i, row in enumerate(DEMO_TASKS):
        r = resp_by_name[row["who"]]
        state = _demo_state(row)
        t = Task(
            respondent=r, workplace_id=r.workplace_id, is_demo=True, name=row["name"], frequency=row["freq"], duration=row["dur"], headcount=row["head"],
            wish=row["wish"], self_improve=C.SELF_IMPROVE_YES if row["self"] else "", prev_task=row["prev"], next_task=row["next"],
            shift=("普通勤務" if r.work_pattern != "交代勤務" else C.SHIFT_OPTIONS[i % 3]),
            steps_json=json.dumps(state["steps"], ensure_ascii=False), intake_json=json.dumps(DEMO_INTAKE.get(row["name"], []), ensure_ascii=False),
            created_at=base + timedelta(days=i % 15, hours=i),
        )
        if row.get("domain"):
            t.domain = _demo_master(dom_by_name, Domain, row["domain"], 100 + i, created["domains"])
        for k, v in DEMO_FLOWS.get(row["name"], {}).items():
            setattr(t, k, v)
        for k, v in DEMO_PROGRESS.get(row["name"], {}).items():
            setattr(t, k, v)
        _seed_stage_history(t, datetime.now())
        db.session.add(t)
        made.append((t, state))
    db.session.flush()

    for t, state in made:
        _finalize_task(t, state, t.intake, use_ai=False)
        if t.name in DEMO_AI_EXTRA:
            t.ai_extra = json.dumps(dict(t.ai_extra_dict, **DEMO_AI_EXTRA[t.name], ai_plan_demo=True), ensure_ascii=False)
    _seed_employees_and_monthly(resp_by_name, wp_by_name, created)
    Setting.set(DEMO_MASTERS_KEY, json.dumps(created))
    db.session.commit()
    return len(made)

DEMO_EXTRA_EMPLOYEES = [
    ("DEMO-009", "小林 直樹", "製造1課"),
    ("DEMO-010", "加藤 恵", "組立課"),
    ("DEMO-011", "吉田 悠", "検査・品質保証課"),
    ("DEMO-012", "松本 玲奈", "生産管理課"),
]


def recent_periods(n=4):
    cur = current_period()
    return [shift_period(cur, -k) for k in range(n - 1, -1, -1)]

def _seed_employees_and_monthly(resp_by_name, wp_by_name, created):
    rows = [(r.emp_id, name, r.workplace, r.work_pattern or "") for name, r in resp_by_name.items()]
    rows += [(emp_id, name, wp_by_name.get(wp), "交代勤務" if j % 2 else "普通勤務") for j, (emp_id, name, wp) in enumerate(DEMO_EXTRA_EMPLOYEES)]
    employees = []
    for i, (emp_id, name, workplace, pattern) in enumerate(rows):
        emp = find_employee(emp_id)
        if emp is None:
            emp = Employee(
                emp_id=emp_id, name=name, department=workplace.name if workplace else "", work_pattern=pattern,
                workplace_id=workplace.id if workplace else None, email=f"user{i + 1:02d}@example.local", active=True, source="csv", is_demo=True,
            )
            db.session.add(emp)
        employees.append(emp)
    db.session.flush()

    resp_by_emp = {r.emp_id: r for r in resp_by_name.values()}
    periods = recent_periods(4)
    last = len(periods) - 1
    hours_options = list(C.MONTHLY_HOURS_OPTIONS)
    stance_options = [C.STANCE_SELF, C.STANCE_TOGETHER, C.STANCE_ASK]
    surveys = {}
    for period in periods:
        survey = MonthlySurvey.query.filter_by(period=period).first()
        if survey is None:
            survey = MonthlySurvey(period=period, is_active=True)
            survey.set_questions(default_questions())
            db.session.add(survey)
            db.session.flush()
            created["surveys"].append(survey.id)
        surveys[period] = survey
    answered = {(p, e) for p, e in db.session.query(MonthlyResponse.period, MonthlyResponse.emp_id).all()}
    for i, emp in enumerate(employees):
        for p_idx, period in enumerate(periods):
            if ((i * 7 + p_idx * 3) % 9 == 0 and p_idx != last) or (period, emp.emp_id) in answered:
                continue
            level = max(0, (i % 3) + 2 - p_idx)
            prev_level = max(0, (i % 3) + 2 - (p_idx - 1))
            has = "ない" if level == 0 else "ある"
            hours = "" if not level else hours_options[min(len(hours_options) - 1, level * 2 + (1 if i % 4 == 0 else 0))]
            if p_idx == 0 or level == prev_level:
                trend = "変わらない"
            elif level < prev_level:
                trend = "減った"
            elif (i + p_idx) % 7 == 0:
                trend = "増えた"
            else:
                trend = "変わらない"
            wants = "ある" if level >= 1 and (i + p_idx) % 3 != 0 else "ない"
            stance_idx = max(0, min(2, (i % 3) - (1 if p_idx >= 2 else 0)))
            r = resp_by_emp.get(emp.emp_id)
            answers = {
                "has": has,
                "hours": hours,
                "task_names": ("\n".join(t.name for t in r.tasks[:2]) if r and r.tasks else "日報の集計\n出勤簿の転記") if has == "ある" else "",
                "trend": trend,
                "decreased_reason": "日報の転記を自動化した" if trend == "減った" else "",
                "wants": wants,
                "wants_tasks": ((r.tasks[0].name if r and r.tasks else "日報の実績の転記") if wants == "ある" else ""),
                "stance": stance_options[stance_idx] if wants == "ある" else "",
                "new_tasks": "新しい報告フォーマットへの転記" if (trend == "増えた" or (trend == "変わらない" and p_idx == last and i % 5 == 0)) else "",
            }
            if has == "ない":
                answers = {k: ("ない" if k == "has" else "") for k in answers}
            comment = ""
            if p_idx == last and i % 4 == 1:
                comment = ["日報の転記が減って朝の時間に余裕ができました", "Excelの集計を自分で自動化してみたい", "新しい帳票が増えたので見直してほしい"][i % 3]
            yy, mm = (int(x) for x in period.split("-"))
            db.session.add(
                MonthlyResponse(
                    period=period, emp_id=emp.emp_id, name=emp.name, workplace_id=emp.workplace_id, work_pattern=emp.work_pattern or "",
                    respondent_id=r.id if r else None, survey_id=surveys[period].id,
                    answers_json=json.dumps(answers, ensure_ascii=False), comment=comment,
                    is_demo=True, created_at=datetime(yy, mm, 5, 9, 0),
                )
            )

def _clear_demo_masters():
    try:
        made = json.loads(Setting.get(DEMO_MASTERS_KEY, "") or "{}")
    except ValueError:
        made = {}
    for sid in made.get("surveys") or []:
        s = db.session.get(MonthlySurvey, sid)
        used = s is not None and MonthlyResponse.query.filter(db.or_(MonthlyResponse.survey_id == s.id, MonthlyResponse.period == s.period)).first() is not None
        if s is not None and not used and s.questions == default_questions():
            db.session.delete(s)
    for did in made.get("domains") or []:
        d = db.session.get(Domain, did)
        if d is not None and Task.query.filter_by(domain_id=d.id).first() is None and Domain.query.filter_by(parent_id=d.id).first() is None:
            db.session.delete(d)
    for wid in made.get("workplaces") or []:
        w = db.session.get(Workplace, wid)
        if w is not None and not any(m.query.filter_by(workplace_id=w.id).first() is not None for m in (Respondent, Task, Employee, MonthlyResponse)):
            db.session.delete(w)
    row = db.session.get(Setting, DEMO_MASTERS_KEY)
    if row is not None:
        db.session.delete(row)

DEMO_ID_RX = re.compile(r"^demo-\d+$", re.I)


def is_demo_id(emp_id):
    return bool(DEMO_ID_RX.match(norm_emp_id(emp_id)))

def demo_ids_with_real_data():
    found = OrderedDict()

    def add(emp_id, name, kind):
        row = found.setdefault(emp_key(emp_id), {"emp_id": norm_emp_id(emp_id), "name": "", "tasks": 0, "monthly": 0})
        row["name"] = row["name"] or (name or "")
        row[kind] += 1

    for t in Task.query.join(Respondent, Task.respondent_id == Respondent.id).filter(Respondent.emp_id.like("demo-%")).order_by(Respondent.emp_id, Task.id).all():
        if not t.is_demo and is_demo_id(t.respondent.emp_id):
            add(t.respondent.emp_id, t.respondent.name, "tasks")
    for m in MonthlyResponse.query.filter(MonthlyResponse.emp_id.like("demo-%")).order_by(MonthlyResponse.emp_id, MonthlyResponse.id).all():
        if not m.is_demo and is_demo_id(m.emp_id):
            add(m.emp_id, m.name, "monthly")
    for row in found.values():
        row["label"] = "デモの社員IDで本物の業務あり" if row["tasks"] else "デモの社員IDで本物の回答あり"
        row["detail"] = "・".join(x for x in (f"業務 {row['tasks']} 件" if row["tasks"] else "", f"月次 {row['monthly']} 回" if row["monthly"] else "") if x)
    return found

def clear_demo():
    db.session.flush()
    task_ids = [i for (i,) in db.session.query(Task.id).filter_by(is_demo=True).all()]
    if task_ids:
        StageLog.query.filter(StageLog.task_id.in_(task_ids)).delete(synchronize_session=False)
        Task.query.filter(Task.id.in_(task_ids)).delete(synchronize_session=False)
    MonthlyResponse.query.filter_by(is_demo=True).delete(synchronize_session=False)
    db.session.expire_all()
    kept = set()
    for r in Respondent.query.filter_by(is_demo=True).all():
        if Task.query.filter_by(respondent_id=r.id).first() is not None or MonthlyResponse.query.filter_by(respondent_id=r.id).first() is not None:
            r.is_demo = False
            kept.add(norm_emp_id(r.emp_id).lower())
        else:
            db.session.delete(r)
    for e in Employee.query.filter_by(is_demo=True).all():
        if e.emp_id.lower() in kept or MonthlyResponse.query.filter(db.func.lower(MonthlyResponse.emp_id) == e.emp_id.lower()).first() is not None:
            e.is_demo = False
        else:
            db.session.delete(e)
    _clear_demo_masters()
    db.session.commit()
    return list(demo_ids_with_real_data().values())

# ===== 画面: ログイン =====
bp_auth = Blueprint("auth", __name__)

DEFAULT_ADMIN_ID = "admin"
DEFAULT_ADMIN_PASSWORD = "adminpass"


def admin_credentials():
    return (
        str(conf.get("ADMIN_ID", "") or DEFAULT_ADMIN_ID).strip(),
        str(conf.get("ADMIN_PASSWORD", "") or DEFAULT_ADMIN_PASSWORD),
    )

def authenticate(user_id, password):
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
    session.clear()
    flash("ログアウトしました。", "secondary")
    return redirect(url_for("board.index"))

# ===== 画面: ダッシュボード =====
bp_main = Blueprint("main", __name__)


def _resolve_urls(suggestions):
    for s in suggestions:
        endpoint, params = s.get("url", (None, {}))
        try:
            s["href"] = url_for(endpoint, **params) if endpoint else None
        except Exception:
            s["href"] = None
    return suggestions

def _rounded_series(rows, *keys):
    return {key: [round(r[key], 1) for r in rows] for key in keys}

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
    monthly_rows = monthly_summary()
    charts["monthly"] = {
        "labels": [r["label"] for r in monthly_rows],
        **_rounded_series(monthly_rows, "has_rate", "decreased_rate", "self_rate"),
    }
    return render_template(
        "index.html",
        k=k,
        charts_json=json.dumps(charts, ensure_ascii=False),
        suggestions=suggestions,
        ai_suggestions=ai_suggestions,
        workplaces=workplace_stats(),
        recent=recent,
        monthly_latest=monthly_rows[-1] if monthly_rows else None,
        active_count=Task.query.filter_by(status="対応中").count(),
    )

@bp_main.post("/ai/suggest")
def ai_suggest():
    if not ai_enabled():
        flash("AI連携が無効です。管理画面の「AI連携」で URL・API キー・モデル名を設定してください。", "warning")
        return redirect(url_for("main.index"))
    k = kpis()
    have, learn = skills_summary()
    tasks = Task.query.all()
    context = {
        "登録業務件数": k["task_count"],
        "回答した人": k["respondent_count"],
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
            {"職場": w["workplace"].name, "回答した人": w["respondents"], "件数": w["tasks"], "月間時間（時間）": round(w["minutes"] / 60, 1)}
            for w in workplace_stats()
        ],
        "重複入力の指摘（入力元別件数）": [{"入力元": src, "件数": len(ts)} for src, ts in duplicate_input_tasks()],
        "同じアウトプットの重複作成": [{"アウトプット": g["output"], "件数": len(g["tasks"])} for g in duplicate_output_groups()[:5]],
        "目的不明の業務": [t.name for t in tasks if t.purpose_unknown][:10],
        "月間時間上位の案件候補": [
            {"業務": t.name, "月間時間（時間）": round((t.monthly_minutes or 0) / 60, 1), "改善希望": t.wish}
            for t in sorted([t for t in tasks if t.is_candidate], key=lambda x: -(x.monthly_minutes or 0))[:8]
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

# ===== 画面: 自動化を検討（1画面・JSON API） =====
bp_survey = Blueprint("survey", __name__, url_prefix="/survey")


def _current():
    rid = session.get("respondent_id")
    return db.session.get(Respondent, rid) if rid else None

def _domains():
    return Domain.query.filter_by(parent_id=None).order_by(Domain.sort, Domain.id).all()

def _workplaces():
    return Workplace.query.order_by(Workplace.sort, Workplace.id).all()

class _JsonForm(dict):
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

def _respondent_json(r):
    return {
        "id": r.id,
        "emp_id": r.emp_id or "",
        "name": r.name,
        "workplace_id": r.workplace_id,
        "workplace": r.workplace.name if r.workplace else "",
        "work_pattern": r.work_pattern or "",
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
        "steps": t.steps,
        "shift": t.shift or "",
        "monthly_name": t.monthly_name or "",
        "stage": t.stage_name + ("（保留中）" if t.is_on_hold else ""),
        "can_delete": _withdrawable(t),
    }

def _withdrawable(t):
    return t.stage_key == "start" and not t.is_on_hold and not _team_touched(t)

def monthly_task_names(text, vocab=None):
    vocab = vocab or map_vocab()
    out, seen = [], set()
    for raw in task_names(text):
        name = mask_person_names(raw, vocab).strip()
        key = name_key(name)
        if key and key not in seen:
            seen.add(key)
            out.append(name)
    return out

def monthly_for(r):
    if r is None or not emp_key(r.emp_id):
        return None
    handed = session.get("monthly_handoff") if has_request_context() else None
    row = db.session.get(MonthlyResponse, handed) if handed else None
    if row is not None and emp_key(row.emp_id) == emp_key(r.emp_id):
        return row
    return MonthlyResponse.query.filter(same_emp_id(MonthlyResponse.emp_id, r.emp_id)).order_by(MonthlyResponse.period.desc(), MonthlyResponse.id.desc()).first()

def _suggested_tasks(r):
    row = monthly_for(r)
    if row is None:
        return [], []
    vocab = map_vocab()
    wants = monthly_task_names(row.answer("wants_tasks"), vocab)
    done = {pending_key(v, vocab) for t in r.tasks for v in (t.name, t.monthly_name) if v}
    keep, seen = [], set()
    for n in wants + monthly_task_names(row.answer("task_names"), vocab):
        k = pending_key(n, vocab)
        if k not in done and k not in seen:
            seen.add(k)
            keep.append(n)
    keep = keep[:8]
    return keep, [n for n in keep if n in wants]

def _stance_json(emp_id):
    stance, row = person_stance(emp_id)
    if not stance:
        return None
    return {"value": stance, "period": row.period, "period_label": period_label(row.period), "self_improve": stance in C.MONTHLY_KPI["self_positive"]}

def _state(r):
    if r is None:
        return {"respondent": None, "tasks": [], "suggested": [], "suggested_wants": [], "stance": None}
    tasks = Task.query.filter_by(respondent_id=r.id).order_by(Task.created_at).all()
    suggested, wants = _suggested_tasks(r)
    return {"respondent": _respondent_json(r), "tasks": [_task_json(t) for t in tasks], "suggested": suggested, "suggested_wants": wants, "stance": _stance_json(r.emp_id)}

def _monthly_context(emp, src, m=None):
    key = emp_key(emp)
    if src != "monthly" or not key:
        return None
    if m:
        row = db.session.get(MonthlyResponse, m)
    else:
        row = db.session.get(MonthlyResponse, session["monthly_last"]) if session.get("monthly_last") else None
        if row is not None and emp_key(row.emp_id) != key:
            row = None
        if row is None:
            row = MonthlyResponse.query.filter(db.func.lower(MonthlyResponse.emp_id) == key).order_by(MonthlyResponse.period.desc(), MonthlyResponse.id.desc()).first()
    if row is None or emp_key(row.emp_id) != key or row.answer("wants") != "ある":
        return None
    wants = monthly_task_names(row.answer("wants_tasks"))
    if not wants:
        return None
    stance = row.answer("stance")
    wp, pattern = (row.workplace_id, row.work_pattern or "") if row.period >= newest_period(row.emp_id) else current_place(row.emp_id, row)
    return {"id": row.id, "period": row.period, "period_label": period_label(row.period), "wants_tasks": wants, "stance": stance,
            "self_improve": stance in C.MONTHLY_KPI["self_positive"] and row.period >= stance_since(), "workplace_id": wp, "work_pattern": pattern}

def _options():
    return {
        "workplaces": [{"id": w.id, "name": w.name} for w in _workplaces()],
        "person": C.PERSON_FIELDS,
        "frequencies": C.FREQUENCIES,
        "durations": C.DURATIONS,
        "wishes": C.WISHES,
        "work_patterns": C.WORK_PATTERNS,
        "shift_options": C.SHIFT_OPTIONS,
        "frequency_hints": C.FREQUENCY_HINTS,
        "self_improve_yes": C.SELF_IMPROVE_YES,
        "skills": C.SKILLS,
        "skill_none": C.SKILL_NONE,
        "max_questions": C.INTAKE_MAX_QUESTIONS,
    }

@bp_survey.route("/")
def index():
    emp = norm_emp_id(request.args.get("emp"))
    task = (request.args.get("task") or "").strip()[:200]
    src = (request.args.get("src") or "").strip()
    session.pop("respondent_id", None)
    monthly = _monthly_context(emp, src, request.args.get("m", type=int))
    if monthly is not None:
        session.pop("monthly_last", None)
        session["monthly_handoff"] = monthly["id"]
    else:
        session.pop("monthly_handoff", None)
    return render_template("survey/index.html", state=_state(None), options=_options(), emp=emp, prefill=task, monthly=monthly)

@bp_survey.route("/start")
def start():
    emp = norm_emp_id(request.args.get("emp"))
    task = (request.args.get("task") or "").strip()[:200]
    src = (request.args.get("src") or "").strip()
    m = request.args.get("m", type=int)
    return redirect(url_for("survey.index", emp=emp or None, task=task or None, src=src or None, m=m or None))

# ----- API: あなたのこと -----
@bp_survey.post("/api/respondent")
def api_respondent():
    f = _JsonForm(request.get_json(silent=True) or {})
    workplace_id = f.get("workplace_id", type=int)
    work_pattern = (f.get("work_pattern") or "").strip()
    info = lookup(f.get("emp_id"))
    emp_id = info["emp_id"] if info else norm_emp_id(f.get("emp_id"))
    name = info["name"] if info else ""
    errors = {}
    if not emp_id:
        errors["emp_id"] = "入力してください"
    elif not info:
        errors["emp_id"] = EMP_NOT_FOUND
    elif not info["active"]:
        errors["emp_id"] = EMP_EXCLUDED
    if not workplace_id:
        errors["workplace_id"] = "選択してください"
    elif db.session.get(Workplace, workplace_id) is None:
        errors["workplace_id"] = "職場を選んでください"
    if work_pattern not in C.WORK_PATTERNS:
        errors["work_pattern"] = "選択してください"
    if errors:
        return jsonify({"success": False, "errors": errors}), 400
    newest = current_period() >= newest_period(emp_id)
    r = Respondent.query.filter(same_emp_id(Respondent.emp_id, emp_id)).order_by(Respondent.id).first()
    resumed = r is not None
    if r is None:
        r = Respondent(emp_id=emp_id)
        db.session.add(r)
    r.emp_id = emp_id
    r.name = name or r.name or ""
    if newest or not r.workplace_id:
        r.workplace_id = workplace_id
        r.work_pattern = work_pattern
    remember_person(emp_id, name, workplace_id if newest else None, work_pattern if newest else None)
    db.session.commit()
    session["respondent_id"] = r.id
    return jsonify({"success": True, "resumed": resumed, **_state(r)})

@bp_survey.post("/api/reset")
def api_reset():
    session.pop("respondent_id", None)
    return jsonify({"success": True})

# ----- API: 検討する業務（AI の聞き取り → 登録） -----
_IMAGE_PREFIX = re.compile(r"^data:image/(png|jpe?g|webp|gif);base64,")
_SOURCE_WORDS = [("紙", "紙の帳票"), ("帳票", "紙の帳票"), ("手書き", "紙の帳票"), ("メール", "メール"), ("Excel", "他のExcel"), ("エクセル", "他のExcel"), ("聞", "人から聞く"), ("口頭", "人から聞く")]
_SYSTEM_WORDS = [("生産管理", "生産管理システム"), ("品質", "品質システム"), ("設備", "設備データ"), ("モニタ", "設備データ")]
_SCREEN_RX = re.compile(r"モニタ|画面|端末|ディスプレイ|タブレット|表示器|表示盤|パネル|HMI", re.I)
_MEDIUM_SOURCES = {"paper": "紙の帳票", "excel": "他のExcel", "mail": "メール", "verbal": "人から聞く"}
_METHOD_SOURCES = {"mail": "メール", "verbal": "人から聞く", "phone": "人から聞く", "phs": "人から聞く", "meeting": "人から聞く", "hand": "紙の帳票", "fax": "紙の帳票"}
_FOLDER_SOURCE = "他のExcel"
_SAVE_RX = re.compile(r"保存|格納|保管|置く|置い|置き|入れ|アップ|上げ")
_TOOL_WORDS = [("Excel", "Excel"), ("エクセル", "Excel"), ("Access", "Access"), ("メール", "メール"), ("システム", "社内システム"), ("紙", "紙"), ("手書き", "紙"), ("印刷", "紙"), ("マクロ", "VBA・マクロ"), ("VBA", "VBA・マクロ"), ("RPA", "RPA")]
NATURE_WORDS = [
    ("転記・コピー", r"転記|書き写|写す|写し|コピー|貼り付|貼りつ|入力|入れ直|打ち込|記入|登録"),
    ("集計・グラフ作成", r"集計(?!表)|グラフ|合計|計算|数え|平均|パレート|ピボット"),
    ("資料作成", r"作成|まとめ|作る|書く|コメント|資料を"),
    ("確認・チェック", r"確認|確かめ|照合|見比べ|突き合わせ|探す|チェック(する|して|を)|点検(する|して)"),
    ("データ取得", r"出力|(?<![提送差届])出す|ダウンロード|抽出|取り出|取得|集める|回収|拾う|抜き出|調べ"),
    ("メール・配布", r"送る|送付|送信|配信|配布|配る|渡す|提出|掲示|貼り出|知らせ|伝え|回覧|メール(する|して|した)|共有.*(メール|チャット|Teams)|(メール|チャット|Teams).*共有"),
    ("台帳更新", r"(台帳|管理表|一覧)[^、。]{0,12}(入力|更新|記入|追加|登録)|ファイリング"),
    ("判断を伴う", r"判断|決め|判定|調整|振り分"),
]
PAIN_WORDS = [
    ("時間がかかる", r"時間が?かか|時間を取ら|手間|面倒|遅い|遅く|長い|集中して|追われ|残業"),
    ("ミスが起きやすい", r"ミス|間違|まちが|誤り|誤記|漏れ|抜け|見落と"),
    ("同じ情報を何度も入力する", r"何度も|何回も|[2-9二三]回|二重|重複|書き写|転記|入れ直|打ち直|再入力|同じ(情報|内容|数字|こと|もの)"),
    ("手順が人によって違う", r"人によ|人それぞれ|属人|バラバラ|ばらばら|やり方が違|決まっていない|その人しか"),
    ("急ぎで割り込まれる", r"急ぎ|急な|急に|割り込|突発|至急|中断"),
    ("何のためか分からない", r"何のため|目的が(分から|わから|不明)|誰が見て|だれが見て|見ているか分から|使われているか|必要か分から|いらないかも"),
]
_NO_PAIN_RX = re.compile(r"(特に)?(ない|なし|無し|ありません|ないです)。?")


def _guess(text, table, default):
    for word, value in table:
        if word.lower() in (text or "").lower():
            return value
    return default


def _ordered(values, choices):
    return [c for c in choices if c in values]


def _source_category(name, method=""):
    name = (name or "").strip()
    medium = medium_of(name) if name else "unknown"
    method = method or (method_of(name) if name else "")
    if medium in ("system", "web") or (medium == "unknown" and _SCREEN_RX.search(vnorm(name))):
        return _guess(name, _SYSTEM_WORDS, "その他")
    found = _MEDIUM_SOURCES.get(medium) or _guess(name, _SOURCE_WORDS, None)
    if not found and method == "folder" and medium in ("unknown", "doc"):
        return _FOLDER_SOURCE
    return found or _METHOD_SOURCES.get(method)


def _mentions_folder(text):
    row = vocab_row("methods", "folder")
    return bool(row and row.get("_rx") and row["_rx"].search(_masked(text, row)))


def _step_natures(x):
    text = (x.get("action") or "") + " " + (x.get("how") or "")
    found = [cat for cat, rx in NATURE_WORDS if re.search(rx, text)]
    if _SAVE_RX.search(x.get("action") or "") and _mentions_folder(text + " " + (x.get("system") or "")):
        found.append("台帳更新")
    return found


def _state_sources(state):
    made = [v for v in [x.get("output") or "" for x in state["steps"]] + list(state["outputs"]) if v.strip()]
    internal = lambda v: any(same_thing(v, m) for m in made)
    found = [(v, "") for v in state["sources"]]
    for x in state["steps"]:
        system, inp, action, to = [(x.get(k) or "").strip() for k in ("system", "input", "action", "to")]
        role, m = field_role(system) if system else ("box", "")
        if system and role == "line":
            if inp and not to and not internal(inp):
                found.append((inp, m))
        elif system and not internal(system) and not _is_destination(system, action):
            found.append((system, m))
        elif inp and not to and not internal(inp):
            found.append((inp, method_of(x.get("how_from") or "")))
    return found


def map_pains(texts):
    found = [cat for cat, rx in PAIN_WORDS if any(re.search(rx, t or "") for t in texts)]
    if not found and any(_NO_PAIN_RX.fullmatch((t or "").strip()) for t in texts):
        found = ["特になし"]
    return found


def rule_categories(task, state):
    steps = state["steps"]
    sources = _state_sources(state)
    inputs = [c for c in (_source_category(n, m) for n, m in sources) if c]
    if sources and not inputs:
        inputs = ["その他"]
    names = list(state["recipients"]) + [n for x in steps for n in _names(x.get("to"))]
    unknown = any(re.search(r"渡す|使う|使い|相手|利用", u) for u in state.get("unknowns") or [])
    users = users_from_names(names, task.workplace.name if task.workplace else "") or (["分からない"] if unknown else [])
    texts = [x.get(k) or "" for x in steps for k in ("system", "action", "how", "how_from")] + list(state["sources"])
    tools = [v for word, v in _TOOL_WORDS if any(word.lower() in t.lower() for t in texts)]
    nature = [cat for x in steps for cat in _step_natures(x)]
    if any((x.get("to") or "").strip() for x in steps):
        nature.append("メール・配布")
    return {
        "inputs": _ordered(inputs, C.INPUT_SOURCES), "output_users": list(dict.fromkeys(users)), "tools": _ordered(tools, C.TOOLS),
        "nature": _ordered(nature, C.NATURES), "pains": map_pains(state["pains"]),
    }


def _clean_messages(items):
    out = []
    for m in items or []:
        if isinstance(m, dict):
            text = str(m.get("text") or "").strip()[:2000]
            if text:
                out.append({"role": "ai" if m.get("role") == "ai" else "user", "text": text})
    return out[-40:]


def _need_respondent(emp_id=None):
    r = _current()
    if r is None:
        return None, (jsonify({"success": False, "message": "先に「あなたのこと」を保存してください", "errors": {"_": "respondent"}}), 400)
    if emp_id is not None and emp_key(emp_id) != emp_key(r.emp_id):
        return None, (jsonify({"success": False, "message": "画面の社員IDと登録先の人が違います。「あなたのこと」を入れ直してください", "errors": {"_": "respondent"}}), 409)
    emp = find_employee(r.emp_id)
    if emp is not None and emp.active is False:
        session.pop("respondent_id", None)
        return None, (jsonify({"success": False, "message": EMP_EXCLUDED, "errors": {"_": "respondent", "emp_id": EMP_EXCLUDED}}), 400)
    return r, None


@bp_survey.post("/api/intake")
def api_intake():
    data = request.get_json(silent=True) or {}
    r, err = _need_respondent(data.get("emp_id"))
    if err:
        return err
    messages = _clean_messages(data.get("messages"))
    state = _clean_state(data.get("state") or {})
    if not state["title"] and data.get("title"):
        state["title"] = mask_person_names(str(data.get("title")).strip())[:200]
    image = data.get("image") or None
    if image is not None and (not isinstance(image, str) or len(image) > 6_000_000 or not _IMAGE_PREFIX.match(image)):
        return jsonify({"success": False, "message": "写真は PNG・JPEG・WebP・GIF で、4MB までにしてください"}), 400
    if not ai_enabled():
        return jsonify({"success": False, "message": "AI連携が設定されていません。管理画面の「AI連携」で URL・API キー・モデル名を設定してください"}), 400
    used = set()
    origin = "monthly" if data.get("origin") == "monthly" else ""
    result = ai_intake(messages, state, image, r.work_pattern or "", handlers=intake_handlers(r, used), tools=intake_tools_for(state), origin=origin)
    if result is None:
        return jsonify({"success": False, "message": "AI から応答がありませんでした。もう一度送ってください（" + (_LAST_ERROR["text"] or "原因不明") + "）"}), 502
    result["state"]["checked"] = sorted(set(state["checked"]) | used)[:20]
    steps = result["state"]["steps"]
    try:
        shown = max(0, min(int(data.get("shown") or 0), len(steps)))
    except (TypeError, ValueError):
        shown = 0
    own = r.workplace.name if r.workplace else ""
    purpose = result["state"].get("purpose") or ""
    result["frames"] = [build_steps_flow(steps[:i], own, purpose).render() for i in range(shown + 1, len(steps) + 1)]
    result["diagram"] = build_steps_flow(steps, own, purpose).render() if steps else ""
    return jsonify({"success": True, **result})


def _fill_from_state(task, state, only_missing=False):
    if not (task.output or "").strip() and state["outputs"]:
        task.output = "、".join(state["outputs"])[:500]
    if not (task.inputs_note or "").strip() and state["sources"]:
        task.inputs_note = "、".join(state["sources"])[:500]
    if not (task.output_users_note or "").strip() and state["recipients"]:
        task.output_users_note = "、".join(state["recipients"])[:500]
    if not (task.pains_note or "").strip() and state["pains"]:
        task.pains_note = "、".join(state["pains"])[:500]
    if (state.get("purpose") or "").strip() and (task.purpose or "").strip() in ("", "不明"):
        task.purpose = state["purpose"].strip()[:500]
    if only_missing:
        return []
    if not (task.purpose or "").strip():
        task.purpose = "不明"
    return _fill_categories(task, state)


def _fill_categories(task, state):
    filled = []
    for key, values in rule_categories(task, state).items():
        if values and not getattr(task, key):
            setattr(task, key, C.join_list(values))
            filled.append(key)
    return filled


def unmark_estimate(task, key):
    task.ai_fields = C.join_list([k for k in task.ai_fields_list if k != key])
    extra = task.ai_extra_dict
    if key in (extra.get("rule_fields") or []):
        extra["rule_fields"] = [k for k in extra["rule_fields"] if k != key]
        task.ai_extra = json.dumps(extra, ensure_ascii=False)


def _finalize_task(task, state, messages, refresh=False, use_ai=True):
    task.compute_minutes()
    domains = _domains()
    extra = task.ai_extra_dict if refresh else {}
    fields = task.ai_fields_list if refresh else []
    rules = list(extra.get("rule_fields") or [])

    def may(key):
        current = task.domain_id if key == "domain" else getattr(task, key)
        return not refresh or not str(current or "").strip() or (key in fields and not task.verified)

    def mark(key, by_rule):
        if key not in fields:
            fields.append(key)
        if by_rule and key not in rules:
            rules.append(key)
        elif not by_rule and key in rules:
            rules.remove(key)

    def set_domain(d, by_rule):
        if d is not None and (not task.domain_id or ("domain" in fields and not task.verified)):
            task.domain = d
            mark("domain", by_rule)

    _fill_from_state(task, state, only_missing=True)
    data = None
    if use_ai and ai_enabled():
        cands = related_tasks_for(task, limit=40)
        data = ai_finalize(state, messages, task.wish, domains, cands, task.shift or "", handlers=finalize_handlers(task), current=task_payload(task) if refresh else None)
    if data is None:
        if any(extra.get(k) for k in C.AI_PLAN_KEYS) and not extra.get("ai_plan_at") and task.ai_source == "openai" and task.ai_at:
            extra["ai_plan_at"] = task.ai_at.strftime("%Y/%m/%d %H:%M")
        for key in _fill_from_state(task, state):
            mark(key, True)
        set_domain(rule_domain(task, domains), True)
        result = rule_analyze(task, domains, related_tasks_for(task, limit=40))
        result["source"] = "rules"
        task.apply_ai_result(result)
        extra.update(first_impression=extra.get("first_impression") or "受け付けました。改善チームが内容を確認して連絡します", unknowns=state.get("unknowns") or [])
    else:
        for key in ("purpose", "output", "output_users_note", "inputs_note", "info_from", "actors", "prev_task", "next_task", "pains_note"):
            value = mask_person_names(str(data.get(key) or "").strip())[:500]
            if value and may(key):
                setattr(task, key, value)
                mark(key, False)
        for key, allowed in (("output_users", C.OUTPUT_USERS), ("inputs", C.INPUT_SOURCES), ("tools", C.TOOLS), ("nature", C.NATURES), ("pains", C.PAINS)):
            values = [v for v in (data.get(key) or []) if isinstance(v, str) and v in allowed]
            if values and may(key):
                setattr(task, key, C.join_list(values))
                mark(key, False)
        if not (task.name or "").strip() and data.get("name"):
            task.name = mask_person_names(str(data["name"]).strip())[:200]
        _fill_from_state(task, state, only_missing=True)
        for key in _fill_categories(task, state):
            mark(key, True)
        ai_domain = next((d for d in domains if d.name == data.get("domain")), None)
        est = ai_domain or rule_domain(task, domains)
        set_domain(est, ai_domain is None)
        real = {n for (n,) in db.session.query(Task.name).filter(Task.id != task.id).all()}
        ecrs = data.get("ecrs") if isinstance(data.get("ecrs"), dict) else {}
        task.apply_ai_result({
            "domain": est.name if est else None,
            "related": list(dict.fromkeys(x for x in (data.get("related") or []) if isinstance(x, str) and x in real))[:5],
            "summary": str(data.get("summary") or ""),
            "ecrs": {
                "eliminate": ecrs.get("eliminate"), "combine": ecrs.get("combine"), "rearrange": ecrs.get("rearrange"), "simplify": ecrs.get("simplify"),
                "automation": data.get("automation"), "questions": [q for q in (data.get("questions") or []) if isinstance(q, str)][:3],
            },
            "flags": [f for f in (data.get("flags") or []) if f in C.AI_FLAGS],
            "priority": data.get("priority") if data.get("priority") in ("高", "中", "低") else "中",
            "source": "openai",
        })
        saving = data.get("saving_minutes_per_month")
        clean = lambda v, n: mask_person_names(str(v).strip())[:n]
        extra.update(
            tobe_steps=[clean(x, 200) for x in (data.get("tobe_steps") or []) if str(x).strip()][:6],
            saving_minutes_per_month=int(saving) if isinstance(saving, (int, float)) and saving > 0 else None,
            first_impression=clean(data.get("first_impression") or "", 120) or extra.get("first_impression") or "",
            hearing_questions=[clean(q, 200) for q in (data.get("hearing_questions") or []) if str(q).strip()][:5],
            unknowns=state.get("unknowns") or [],
            ai_plan_at=datetime.now().strftime("%Y/%m/%d %H:%M"),
        )
        extra.pop("ai_plan_demo", None)
        try:
            guess = int(data.get("headcount") or 1)
        except (TypeError, ValueError):
            guess = 1
        if guess > 1:
            extra["headcount_guess"] = guess
        else:
            extra.pop("headcount_guess", None)
    task.ai_fields = C.join_list(fields)
    extra["rule_fields"] = [k for k in rules if k in fields]
    extra.pop("ai_before_edit", None)
    task.ai_extra = json.dumps(extra, ensure_ascii=False)


@bp_survey.post("/api/tasks")
def api_task_create():
    data = request.get_json(silent=True) or {}
    r, err = _need_respondent(data.get("emp_id"))
    if err:
        return err
    state = _clean_state(data.get("state") or {})
    messages = [dict(m, text=mask_person_names(m["text"])) for m in _clean_messages(data.get("messages"))]
    errors = {}
    name = mask_person_names(str(data.get("name") or "").strip() or state["title"])[:200]
    if not name:
        errors["name"] = "業務名を入力してください"
    if not state["steps"]:
        errors["steps"] = "手順を1つ以上聞き取ってから登録してください"
    frequency = data.get("frequency") or state["frequency"]
    duration = data.get("duration") or state["duration"]
    if frequency not in C.FREQUENCIES:
        errors["frequency"] = "頻度を選んでください"
    if duration not in C.DURATIONS:
        errors["duration"] = "1回の時間を選んでください"
    wish = data.get("wish")
    if wish not in C.WISHES:
        errors["wish"] = "この業務をどうしたいかを選んでください"
    shift = str(data.get("shift") or state.get("shift") or "").strip()
    if (r.work_pattern or "") == "交代勤務":
        if shift not in C.SHIFT_OPTIONS:
            errors["shift"] = "この業務をいつやるか（昼勤／夜勤／両方）を選んでください"
    else:
        shift = "普通勤務"
    if errors:
        return jsonify({"success": False, "errors": errors, "message": next(iter(errors.values()))}), 400
    dup = next((t for t in r.tasks if normalize_text(t.name) == normalize_text(name)), None)
    if dup is not None and not data.get("force"):
        return jsonify({
            "success": False, "message": f"同じ名前の業務「{dup.name}」を登録済みです",
            "duplicate": {"id": dup.id, "name": dup.name, "hours": round((dup.monthly_minutes or 0) / 60, 1), "created_at": dup.created_at.strftime("%Y/%m/%d") if dup.created_at else "",
                          "can_delete": _withdrawable(dup)},
        }), 409
    try:
        headcount = max(1, int(data.get("headcount") or 1))
    except (TypeError, ValueError):
        headcount = 1
    task = Task(
        respondent_id=r.id, workplace_id=r.workplace_id, name=name, frequency=frequency, duration=duration, headcount=headcount, wish=wish, shift=shift,
        self_improve=C.SELF_IMPROVE_YES if data.get("self_improve") else "", monthly_name=mask_person_names(str(data.get("monthly_name") or "").strip())[:200],
        steps_json=json.dumps(state["steps"], ensure_ascii=False), intake_json=json.dumps(messages, ensure_ascii=False),
    )
    if data.get("self_improve"):
        skills = [v for v in (data.get("skills") or []) if v in C.SKILLS]
        learn = [v for v in (data.get("learn") or []) if v in C.SKILLS]
        skills, learn = [[v for v in vals if v != C.SKILL_NONE] if len(vals) > 1 else vals for vals in (skills, learn)]
        r.c2_skills = C.join_list(skills)
        r.c3_learn = C.join_list(learn)
        r.part_c_done = True
        task.assignment = "自主改善"
    db.session.add(task)
    db.session.flush()
    task.compute_minutes()
    _finalize_task(task, state, messages)
    if find_employee(r.emp_id) is None:
        register_member(r.emp_id, r.name, r.workplace_id, r.work_pattern)
    db.session.commit()
    return jsonify({"success": True, "task": _task_json(task), "first_impression": task.ai_extra_dict.get("first_impression") or "", **_state(r)})


@bp_survey.delete("/api/tasks/<int:task_id>")
def api_task_delete(task_id):
    r, err = _need_respondent((request.get_json(silent=True) or {}).get("emp_id"))
    if err:
        return err
    task = db.session.get(Task, task_id)
    if task is None or task.respondent_id != r.id:
        abort(404)
    if not _withdrawable(task):
        return jsonify({"success": False, "message": "改善チームが対応を始めているため取り消せません"}), 409
    db.session.delete(task)
    db.session.commit()
    return jsonify({"success": True, **_state(r)})


# ===== 画面: 改善案件 =====
bp_cases = Blueprint("cases", __name__, url_prefix="/cases")


def _get_or_404(model, ident):
    obj = db.session.get(model, ident)
    if obj is None:
        abort(404)
    return obj

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
    task = _get_or_404(Task, task_id)
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
        domains=_domains(),
        **dict(zip(("task_flow", "task_legend"), flow_view(build_task_flow(task)))),
    )

def _update_task_info(task, f):
    changed = False
    if "domain_id" in f:
        d = db.session.get(Domain, f.get("domain_id", type=int) or 0)
        if (d.id if d else None) != task.domain_id:
            task.domain = d
            unmark_estimate(task, "domain")
            changed = True
    minutes = False
    if "headcount" in f:
        try:
            head = max(1, int(f.get("headcount") or 1))
        except ValueError:
            head = task.headcount or 1
        if head != (task.headcount or 1):
            task.headcount = head
            minutes = True
    for key, choices in (("frequency", C.FREQUENCIES), ("duration", C.DURATIONS)):
        value = f.get(key)
        if value in choices and value != getattr(task, key):
            setattr(task, key, value)
            unmark_estimate(task, key)
            minutes = True
    if minutes:
        task.compute_minutes()
        changed = True
    for key in ("purpose", "output", "output_users_note", "inputs_note", "prev_task", "next_task", "actors"):
        if key in f:
            value = mask_person_names((f.get(key) or "").replace("\r\n", "\n").strip())[:500]
            if value != (getattr(task, key) or "").strip():
                setattr(task, key, value)
                unmark_estimate(task, key)
                changed = True
    for key, choices in (("output_users", C.OUTPUT_USERS), ("inputs", C.INPUT_SOURCES)):
        if key in f:
            values = _ordered(f.getlist(key), choices)
            if set(values) != set(getattr(task, key + "_list")):
                setattr(task, key, C.join_list(values))
                unmark_estimate(task, key)
                changed = True
    return changed

def _after_info_edit(task):
    extra = task.ai_extra_dict
    extra["info_edited_at"] = datetime.now().strftime("%Y/%m/%d %H:%M")
    kind = ""
    if task.ai_source in ("openai", "claude"):
        extra.setdefault("ai_before_edit", task.ai_at.strftime("%Y/%m/%d %H:%M") if task.ai_at else "日付不明")
        kind = "ai"
    elif task.ai_source:
        result = rule_analyze(task, _domains(), related_tasks_for(task, limit=40))
        result["source"] = "rules"
        task.apply_ai_result(result)
        kind = "rules"
    task.ai_extra = json.dumps(extra, ensure_ascii=False)
    return kind

@bp_cases.post("/<int:task_id>/update")
def update(task_id):
    task = _get_or_404(Task, task_id)
    f = request.form
    info_changed = _update_task_info(task, f)
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
    refreshed = _after_info_edit(task) if info_changed else ""
    db.session.commit()
    flash("案件を更新しました（業務情報を直したので、ルールベースの分析もやり直しました）。" if refreshed == "rules" else "案件を更新しました。", "success")
    if refreshed == "ai":
        flash("業務情報を直しました。AI 分析を新しくするには「再分析」を押してください", "warning")
    return redirect(url_for("cases.detail", task_id=task.id))

@bp_cases.post("/<int:task_id>/stage")
def set_stage(task_id):
    task = _get_or_404(Task, task_id)
    f = request.form
    key = f.get("stage")
    if key not in C.STAGE_INDEX:
        step = 1 if f.get("direction") == "next" else -1
        idx = min(max(task.stage_index + step, 0), len(C.STAGE_KEYS) - 1)
        key = C.STAGE_KEYS[idx]
    if task.is_on_hold:
        task.status = "対応中"
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
    task = _get_or_404(Task, task_id)
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
    task = _get_or_404(Task, task_id)
    guess = task.domain_guess
    task.verified = not task.verified
    task.verified_by = (request.form.get("verified_by") or "推進メンバー").strip() if task.verified else ""
    if task.verified and guess and request.form.get("adopt_ai_domain"):
        d = Domain.query.filter_by(name=guess).first()
        if d:
            task.domain = d
            unmark_estimate(task, "domain")
    db.session.commit()
    flash("担当者確認を更新しました。" if task.verified else "担当者確認を解除しました。", "success")
    return redirect(url_for("cases.detail", task_id=task.id))

@bp_cases.post("/<int:task_id>/reanalyze")
def reanalyze(task_id):
    task = _get_or_404(Task, task_id)
    _finalize_task(task, task_state(task), task.intake, refresh=True)
    db.session.commit()
    if ai_enabled() and task.ai_source != "openai":
        flash("AI から使える応答がなかったため、ルールベースで分析しました（" + (_LAST_ERROR["text"] or "応答を読めませんでした") + "）。"
              + ("前回の AI の案は日付を付けて分けて表示しています。" if task.ai_plan_note else ""), "warning")
    else:
        flash("AI分析を更新しました（" + ("ChatGPT API" if task.ai_source == "openai" else "ルールベース") + "）。", "success")
    return redirect(url_for("cases.detail", task_id=task.id))

def _csv_response(buf, filename):
    return Response(buf.getvalue().encode("utf-8-sig"), mimetype="text/csv", headers={"Content-Disposition": "attachment; filename=" + filename})

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
            "担当者確認", "AI領域", "AI優先度", "AIフラグ", "AI要約", "登録日", "勤務", "手順（聞き取り）", "AI推定の項目",
            "分からなかったこと", "改善後の流れ（AI案）", *[f"ECRSメモ {label}" for _key, label, _q in C.ECRS], "改善内容・結果", "確認者",
        ]
    )
    for t in Task.query.order_by(Task.id).all():
        extra = t.ai_extra_dict
        w.writerow(
            [
                t.id, t.name, t.respondent.name if t.respondent else "", t.workplace.name if t.workplace else "",
                t.domain.name if t.domain else "", t.purpose, t.output, "・".join(t.output_users_list), t.output_users_note,
                "・".join(t.inputs_list), t.inputs_note, t.prev_task, t.next_task,
                t.actors, t.info_from,
                t.frequency, t.duration, t.headcount,
                t.monthly_minutes, "・".join(t.tools_list), "・".join(t.nature_list), "・".join(t.pains_list), t.pains_note,
                t.wish, t.self_improve or C.SELF_IMPROVE_UNCHECKED, t.note, t.status, t.stage_name, t.stage_since.strftime("%Y-%m-%d") if t.stage_since else "", t.progress_note,
                t.decision, t.assignment, t.assignee, "・".join(str(s) for s in t.steps_done_list),
                "" if t.after_minutes is None else t.after_minutes, t.saved_minutes, "済" if t.verified else "",
                t.ai_domain or "", t.ai_priority or "", "・".join(t.ai_flags_list), (f"［{t.ai_stale_note}］" if t.ai_stale_note else "") + (t.ai_summary or ""),
                t.created_at.strftime("%Y-%m-%d %H:%M") if t.created_at else "",
                t.shift or "", "\n".join(_step_line(x) for x in t.steps), "・".join(t.ai_fields_list),
                "・".join(extra.get("unknowns") or []), (f"［{t.ai_plan_note}］" if t.ai_plan_note and extra.get("tobe_steps") else "") + " → ".join(extra.get("tobe_steps") or []),
                *[getattr(t, "ecrs_" + key) or "" for key, _label, _q in C.ECRS],
                t.result_note or "", t.verified_by or "",
            ]
        )
    return _csv_response(buf, "tasks_{}.csv".format(datetime.now().strftime("%Y%m%d")))

def _step_line(x):
    line = f"{x.get('no')}. {x.get('action') or ''}" + (f"［{x['system']}］" if x.get("system") else "") + (f" → {x['output']}" if x.get("output") else "")
    if x.get("to"):
        line += f"（{x['to']}へ" + (f"・{x['how']}" if x.get("how") else "") + "）"
    return line

# ===== 画面: 業務マップ =====
bp_bizmap = Blueprint("bizmap", __name__, url_prefix="/bizmap")

MAX_TASKS = 80
KINDS = ("info",)
KIND_LABELS = {"info": "業務マップ"}


def _filtered_tasks(domain_id, workplace_id):
    q = Task.query
    if domain_id:
        d = db.session.get(Domain, domain_id)
        ids = [domain_id] + ([c.id for c in d.children] if d else [])
        q = q.filter(Task.domain_id.in_(ids))
    if workplace_id:
        q = q.filter(Task.workplace_id == workplace_id)
    tasks = q.order_by(Task.workplace_id, Task.created_at).all()
    return tasks[:MAX_TASKS], len(tasks) > MAX_TASKS

def _build(tasks, workplace_id=None, domain_id=None):
    d = build_info_flow(tasks)
    return d if domain_id else add_pending_tasks(d, workplace_id, tasks)

def _scope_label(workplace_id, domain_id):
    parts = []
    if workplace_id:
        w = db.session.get(Workplace, workplace_id)
        parts.append(w.name if w else f"職場 {workplace_id}")
    if domain_id:
        d = db.session.get(Domain, domain_id)
        parts.append(d.name if d else f"領域 {domain_id}")
    return "／".join(parts) if parts else "製造部全体"

def _graph_data(row):
    if not row:
        return None
    try:
        return json.loads(row.data or "{}")
    except ValueError:
        return None

def _merged_graph(kind, workplace_id, domain_id):
    tasks, truncated = _filtered_tasks(domain_id, workplace_id)
    auto = _build(tasks, workplace_id, domain_id).to_graph()
    scope = scope_key(workplace_id, domain_id)
    row = FlowMap.query.filter_by(kind=kind, scope=scope).first()
    merged = merge_graph(auto, _graph_data(row))
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
    tasks, truncated = _filtered_tasks(domain_id, workplace_id)

    flows, saved_at, kind_legends = {}, {}, {}
    for kind in KINDS:
        merged = _merged_graph(kind, workplace_id, domain_id)
        flows[kind], kind_legends[kind] = flow_view(graph_to_diagram(merged) if merged["nodes"] else None)
        saved_at[kind] = merged["meta"]["saved_at"]
        pending_count = len([n for n in merged["nodes"] if n.get("cls") == "pending"])
    rows = domain_analysis()
    return render_template(
        "bizmap/index.html",
        flows=flows,
        kind_legends=kind_legends,
        pending_count=pending_count,
        saved_at=saved_at,
        task_count=len(tasks),
        flow_info_count=len([t for t in tasks if t.has_flow_info or t.steps]),
        source_rows=source_rows(tasks),
        handoff_rows=handoff_rows(tasks),
        truncated=truncated,
        max_tasks=MAX_TASKS,
        workplaces=_workplaces(),
        domains=_domains(),
        args=request.args,
        rows=rows,
        dup_outputs=duplicate_output_groups(),
        dup_inputs=duplicate_input_tasks(),
        flow_rows=flow_rows(),
    )

def _kind_and_scope(payload=None):
    src = request.args if payload is None else payload
    kind = src.get("kind") or "info"
    if kind not in KINDS:
        abort(404)
    if payload is None:
        return kind, request.args.get("workplace_id", type=int), request.args.get("domain_id", type=int)
    return kind, payload.get("workplace_id") or None, payload.get("domain_id") or None

@bp_bizmap.route("/edit")
def edit():
    kind, workplace_id, domain_id = _kind_and_scope()
    return render_template(
        "bizmap/edit.html",
        kind=kind,
        kind_label=KIND_LABELS[kind],
        workplace_id=workplace_id,
        domain_id=domain_id,
        scope_label=_scope_label(workplace_id, domain_id),
        kinds_json=json.dumps(editor_config(), ensure_ascii=False).replace("</", "<\\/"),
    )

# ----- グラフ API（エディタ用） -----
@bp_bizmap.post("/api/classify")
def api_classify():
    data = request.get_json(silent=True) or {}
    out = []
    for it in (data.get("items") or [])[:300]:
        if not isinstance(it, dict):
            continue
        if it.get("type") == "info":
            out.append({"id": it.get("id"), "medium": medium_of(str(it.get("text") or "")[:200])})
        else:
            w = who_of(str(it.get("who") or "")[:100], str(it.get("workplace") or "")[:100])
            out.append({"id": it.get("id"), "actor": w["actor"], "who": w["label"], "who_state": w["state"]})
    return jsonify({"items": out})


@bp_bizmap.route("/api/graph", methods=["GET"])
def api_graph_get():
    kind, workplace_id, domain_id = _kind_and_scope()
    return jsonify(_merged_graph(kind, workplace_id, domain_id))

@bp_bizmap.route("/api/graph", methods=["POST"])
def api_graph_save():
    payload = request.get_json(silent=True) or {}
    kind, workplace_id, domain_id = _kind_and_scope(payload)
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
    kind, workplace_id, domain_id = _kind_and_scope()
    row = FlowMap.query.filter_by(kind=kind, scope=scope_key(workplace_id, domain_id)).first()
    if row:
        db.session.delete(row)
        db.session.commit()
    return jsonify({"ok": True})

# ----- 手修正の業務データへの書き戻し -----
def _saved_graph(kind, workplace_id, domain_id):
    return _graph_data(FlowMap.query.filter_by(kind=kind, scope=scope_key(workplace_id, domain_id)).first())

def _all_tasks(saved):
    return {t.id: t for t in Task.query.all()} if saved else {}

@bp_bizmap.route("/api/writeback", methods=["GET"])
def api_writeback_candidates():
    kind, workplace_id, domain_id = _kind_and_scope()
    saved = _saved_graph(kind, workplace_id, domain_id)
    cands = writeback_candidates(saved, _all_tasks(saved))
    return jsonify({"candidates": [{"id": c["id"], "task_id": c["task_id"], "text": c["text"], "checked": c.get("checked", True)} for c in cands]})

@bp_bizmap.route("/api/writeback", methods=["POST"])
def api_writeback_apply():
    payload = request.get_json(silent=True) or {}
    kind, workplace_id, domain_id = _kind_and_scope(payload)
    selected = set(str(x) for x in (payload.get("ids") or []))
    saved = _saved_graph(kind, workplace_id, domain_id)
    tasks_by_id = _all_tasks(saved)
    cands = writeback_candidates(saved, tasks_by_id)
    applied, ai_stale = apply_writeback(cands, selected, tasks_by_id)
    db.session.commit()
    return jsonify({"ok": True, "applied": applied, "selected": len(selected), "ai_stale": ai_stale})

# ----- 業務領域ごとのページ -----
@bp_bizmap.route("/domain/<int:domain_id>")
def domain(domain_id):
    d = _get_or_404(Domain, domain_id)
    row = next((r for r in domain_analysis() if r["domain"].id == d.id), None)
    ids = [d.id] + [c.id for c in d.children]
    tasks = Task.query.filter(Task.domain_id.in_(ids)).all()
    tasks.sort(key=lambda t: t.monthly_minutes or 0, reverse=True)
    merged = _merged_graph("info", None, d.id)
    info_flow, info_legend = flow_view(graph_to_diagram(merged) if merged["nodes"] else None)
    return render_template(
        "bizmap/domain.html",
        d=d,
        row=row,
        tasks=tasks,
        missing=[t for t in tasks if not t.required_filled],
        unverified=[t for t in tasks if not t.verified],
        flows=[t for t in tasks if (t.prev_task or "").strip() or (t.next_task or "").strip()],
        info_flow=info_flow,
        info_legend=info_legend,
    )

# ===== 画面: DX人財情報 =====
bp_people = Blueprint("people", __name__, url_prefix="/people")


@bp_people.route("/")
def index():
    workplace_id = request.args.get("workplace_id", type=int)
    wish = (request.args.get("wish") or "").strip()
    everyone = dx_people()
    people = dx_people(workplace_id=workplace_id, wish=wish or None) if (workplace_id or wish) else everyone
    have, learn = skills_summary(everyone)
    latest_period = max((p["latest_period"] for p in everyone if p["latest_period"]), default="")
    in_latest = [p["monthly"][0] for p in everyone if latest_period and p["latest_period"] == latest_period]
    summary = {
        "total": len(everyone),
        "positive": len([p for p in everyone if p["positive"]]),
        "self_now": len([m for m in in_latest if m.answer("stance") in C.MONTHLY_KPI["self_positive"]]),
        "period_label": period_label(latest_period) if latest_period else "月次の回答はまだありません",
        "learners": len([p for p in everyone if p["learn"]]),
    }
    positive_counts = Counter((p["work_pattern"] or "未設定") for p in everyone if p["positive"])
    summary["positive_by_pattern"] = {k: positive_counts[k] for k in list(C.WORK_PATTERNS) + ["未設定"] if positive_counts.get(k)}
    latest_stances = Counter(m.answer("stance") for m in in_latest if m.answer("stance"))
    stance_options = next((q["options"] for q in C.MONTHLY_CORE_QUESTIONS if q["key"] == "stance"), [])
    wish_counts = [("検討: " + opt, len([p for p in everyone if p["wish"] == opt])) for opt in C.DX_WISH_OPTIONS]
    wish_counts += [("月次: " + opt, latest_stances.get(opt, 0)) for opt in stance_options]
    return render_template(
        "people/index.html",
        people=people,
        summary=summary,
        wish_counts=wish_counts,
        have=have.most_common(),
        learn=learn.most_common(),
        workplaces=_workplaces(),
        args={"workplace_id": workplace_id, "wish": wish},
    )

# ===== 画面: 管理 =====
bp_admin = Blueprint("admin", __name__, url_prefix="/admin")


@bp_admin.route("/")
def index():
    return render_template(
        "admin/index.html",
        workplaces=_workplaces(),
        workplace_usage=_workplace_usage(),
        domains=_domains(),
        ai_status=ai_status_text(),
        ldap_status=ldap_status_text(),
        demo_count=Task.query.filter_by(is_demo=True).count(),
        answered_count=len(answered_people()),
        task_count=Task.query.count(),
        auth_status=auth_status_text(),
    )

@bp_admin.post("/settings")
def save_settings():
    for key in ("project_name", "catchphrase", "sub_message"):
        Setting.set(key, (request.form.get(key) or "").strip())
    Setting.set("board_show_names", "1" if request.form.get("board_show_names") else "0")
    db.session.commit()
    flash("設定を保存しました。", "success")
    return redirect(url_for("admin.index"))


def _save_ai_form():
    for key, _label in AI_SETTING_LABELS:
        Setting.set(key, (request.form.get(key) or "").strip())
    db.session.commit()
    _LAST_ERROR["text"] = ""
    _NO_TOOLS.clear()


@bp_admin.post("/ai")
def save_ai_settings():
    _save_ai_form()
    if ai_enabled():
        flash("AI連携の設定を保存しました。", "success")
    else:
        flash("AI連携の設定を保存しました。URL・API キー・モデル名がそろうまではルールベースで動きます。", "success")
    return redirect(url_for("admin.index"))


@bp_admin.post("/ai/test")
def test_ai_settings():
    _save_ai_form()
    if not ai_enabled():
        flash("URL・API キー・モデル名をすべて入力してから接続テストしてください。", "warning")
        return redirect(url_for("admin.index"))
    text = _call_ai("接続確認です。", "「OK」とだけ返してください。")
    if text is None:
        flash("接続できませんでした: " + (_LAST_ERROR["text"] or "原因不明"), "danger")
        return redirect(url_for("admin.index"))
    called = []
    ping = [{"type": "function", "function": {"name": "ping", "description": "接続確認用。呼ぶと pong を返す", "parameters": {"type": "object", "properties": {}}}}]
    text2 = _call_ai("接続確認です。道具 ping を1回呼び、その結果を見てから「OK」とだけ返してください。", "ping を呼んでください。",
                     tools=ping, handlers={"ping": lambda args: called.append(1) or {"pong": True}})
    if called and text2:
        flash("接続できました。応答: " + " ".join(text.split())[:60] + "／道具の呼び出し（tool calling）も通りました", "success")
    else:
        flash("接続できました。応答: " + " ".join(text.split())[:60] + "／道具の呼び出しは通りませんでした（" + (_LAST_ERROR["text"] or "tools が中継されていない可能性") + "）。聞き取りは道具なしで動きます", "warning")
    return redirect(url_for("admin.index"))

# ----- 職場・業務領域 -----
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

_WORKPLACE_USERS = (("業務", "件", Task), ("月次の回答", "件", MonthlyResponse), ("対象者", "人", Employee), ("「自動化を検討」の回答者", "人", Respondent))


def _workplace_usage():
    usage = defaultdict(list)
    for label, unit, model in _WORKPLACE_USERS:
        for wp_id, n in db.session.query(model.workplace_id, db.func.count(model.id)).filter(model.workplace_id.isnot(None)).group_by(model.workplace_id).all():
            usage[wp_id].append(f"{label} {n} {unit}")
    return usage

@bp_admin.post("/workplace/<int:workplace_id>/update")
def update_workplace(workplace_id):
    w = _get_or_404(Workplace, workplace_id)
    name = (request.form.get("name") or "").strip()
    if not name:
        flash("職場名を入力してください。", "danger")
    elif Workplace.query.filter(Workplace.name == name, Workplace.id != w.id).first():
        flash("同じ名前の職場があります。", "warning")
    elif name != w.name:
        remember_workplace_name(w.id, w.name, name)
        w.name = name
        db.session.commit()
        flash("職場名を変更しました。前の名前で書かれた業務の相手・前後の業務・受け取り元も、業務マップではこの職場として扱います。", "success")
    return redirect(url_for("admin.index"))

@bp_admin.post("/workplace/<int:workplace_id>/delete")
def delete_workplace(workplace_id):
    w = _get_or_404(Workplace, workplace_id)
    used = _workplace_usage().get(w.id)
    if used:
        flash(f"「{w.name}」は使われているため削除できません（{'・'.join(used)}）。名前を変えるときは、職場名を直して保存してください。", "warning")
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
    d = _get_or_404(Domain, domain_id)
    name = (request.form.get("name") or "").strip()
    if name and name != d.name:
        Task.query.filter_by(ai_domain=d.name).update({"ai_domain": name}, synchronize_session=False)
        d.name = name
    d.description = (request.form.get("description") or "").strip()
    d.sort = request.form.get("sort", type=int) or d.sort
    db.session.commit()
    flash("業務領域を更新しました。", "success")
    return redirect(url_for("admin.index"))

@bp_admin.post("/domain/<int:domain_id>/delete")
def delete_domain(domain_id):
    d = _get_or_404(Domain, domain_id)
    if d.tasks or d.children:
        flash("業務が紐づいている領域は削除できません。先に業務の領域を変更してください。", "warning")
    else:
        db.session.delete(d)
        db.session.commit()
        flash("業務領域を削除しました。", "info")
    return redirect(url_for("admin.index"))

# ----- 対象者一覧（employees テーブル） -----
def _member_payload():
    rows = member_rows()
    periods = member_periods()
    return rows, periods, member_stats(rows, periods)

@bp_admin.route("/employees")
def employees():
    rows, periods, stats = _member_payload()
    return render_template(
        "admin/employees.html",
        rows=rows,
        periods=periods,
        stats=stats,
        active_count=len([r for r in rows if r["emp"].active]),
        ldap_status=ldap_status_text(),
        workplaces=_workplaces(),
        demo_real=demo_ids_with_real_data(),
    )

def _upsert_employee(emp_id, name, department, email, workplace_id=None, source="manual", work_pattern=None):
    emp_id = norm_emp_id(emp_id)
    emp = find_employee(emp_id)
    created = emp is None
    if emp is None:
        emp = Employee(emp_id=emp_id, name=name)
        db.session.add(emp)
    emp.name = name or emp.name
    emp.department = department if department is not None else emp.department
    emp.email = email if email is not None else emp.email
    emp.source = source
    if work_pattern in C.WORK_PATTERNS:
        emp.work_pattern = work_pattern
    if workplace_id:
        emp.workplace_id = workplace_id
    else:
        wp = workplace_for(emp.department)
        if wp is not None:
            emp.workplace_id = wp.id
    return created

@bp_admin.post("/employees/add")
def add_employee():
    emp_id = norm_emp_id(request.form.get("emp_id"))
    name = (request.form.get("name") or "").strip()
    if not emp_id or not name:
        flash("社員IDと氏名を入力してください。", "danger")
        return redirect(url_for("admin.employees"))
    created = _upsert_employee(
        emp_id, name, (request.form.get("department") or "").strip(), (request.form.get("email") or "").strip(),
        request.form.get("workplace_id", type=int), "manual", (request.form.get("work_pattern") or "").strip() or None,
    )
    db.session.commit()
    flash("社員を追加しました。" if created else "社員情報を更新しました。", "success")
    return redirect(url_for("admin.employees"))

@bp_admin.post("/employees/<int:employee_id>/toggle")
def toggle_employee(employee_id):
    e = _get_or_404(Employee, employee_id)
    data = request.get_json(silent=True) if request.is_json else None
    was_active = bool(e.active)
    e.active = (not data["excluded"]) if isinstance(data, dict) and "excluded" in data else not e.active
    if e.active:
        e.excluded_at = None
    elif was_active:
        e.excluded_at = datetime.now()
    db.session.commit()
    if _wants_json():
        rows, periods, stats = _member_payload()
        active = len([r for r in rows if r["emp"].active])
        return jsonify({"success": True, "active": e.active, "active_count": active, "excluded_count": len(rows) - active,
                        "stats": [dict(period=p, **stats[p]) for p in periods]})
    return redirect(url_for("admin.employees"))

@bp_admin.post("/employees/<int:employee_id>/delete")
def delete_employee(employee_id):
    e = _get_or_404(Employee, employee_id)
    key = emp_key(e.emp_id)
    if any(emp_key(x) == key for model in (MonthlyResponse, Respondent) for (x,) in db.session.query(model.emp_id).distinct().all()):
        flash(f"{e.emp_id} は月次アンケートの回答か「自動化を検討」の登録があるため削除できません（削除しても次の起動で一覧に戻ります）。回答率の分母から外すときは「対象外」にチェックしてください。", "warning")
        return redirect(url_for("admin.employees"))
    db.session.delete(e)
    db.session.commit()
    flash("社員を削除しました。", "info")
    return redirect(url_for("admin.employees"))

_HEADER_ALIASES = {
    "emp_id": ["社員id", "社員番号", "社員コード", "emp_id", "employee_id", "employeeid", "id"],
    "name": ["氏名", "名前", "name", "displayname"],
    "department": ["所属", "部署", "職場", "department", "dept"],
    "email": ["メール", "メールアドレス", "email", "mail", "e-mail"],
    "work_pattern": ["勤務形態", "勤務", "work_pattern", "shift"],
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
        emp_id = norm_emp_id(row.get(cols["emp_id"]))
        name = (row.get(cols["name"]) or "").strip()
        if not emp_id or not name:
            continue
        dept = (row.get(cols["department"]) or "").strip() if "department" in cols else None
        email = (row.get(cols["email"]) or "").strip() if "email" in cols else None
        pattern = (row.get(cols["work_pattern"]) or "").strip() if "work_pattern" in cols else None
        if _upsert_employee(emp_id, name, dept, email, None, "csv", pattern):
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
    w.writerow(["社員ID", "氏名", "所属", "職場", "勤務形態", "メール", "対象", "元"])
    for e in Employee.query.order_by(Employee.workplace_id, Employee.emp_id).all():
        w.writerow([e.emp_id, e.name, e.department, e.workplace.name if e.workplace else "", e.work_pattern or "", e.email, "1" if e.active else "0", e.source])
    return _csv_response(buf, "employees.csv")

# ----- 業務マップの区分 -----
@bp_admin.route("/map-vocab")
def map_vocab_page():
    return render_template(
        "admin/map_vocab.html", vocab_json=json.dumps(vocab_public(), ensure_ascii=False).replace("</", "<\\/"),
        choices_json=json.dumps({"icons": ICON_CHOICES, "shapes": SHAPE_CHOICES, "directions": DIRECTION_CHOICES, "kinds": METHOD_KIND_CHOICES, "users": USER_CHOICES}, ensure_ascii=False),
    )

@bp_admin.post("/map-vocab")
def map_vocab_save():
    data = request.get_json(silent=True) or {}
    if data.get("reset"):
        Setting.set(MAP_VOCAB_KEY, json.dumps({"_seq": map_vocab().get("_seq", 0)}))
    else:
        posted = data.get("vocab") if isinstance(data.get("vocab"), dict) else {}
        posted = dict(posted, _seq=max(int(map_vocab().get("_seq") or 0), int(posted.get("_seq") or 0) if str(posted.get("_seq") or "0").isdigit() else 0))
        Setting.set(MAP_VOCAB_KEY, json.dumps(_clean_vocab(posted), ensure_ascii=False))
    db.session.commit()
    return jsonify({"success": True, "vocab": vocab_public()})

@bp_admin.post("/map-vocab/try")
def map_vocab_try():
    data = request.get_json(silent=True) or {}
    vocab = _compile_vocab(_clean_vocab(data.get("vocab"))) if isinstance(data.get("vocab"), dict) else map_vocab()
    text = str(data.get("text") or "")[:200]
    medium = medium_of(text, vocab=vocab)
    who = who_of(text, str(data.get("workplace") or ""), vocab)
    role, method = field_role(text, vocab)
    handoff = handoff_of(text, vocab=vocab)
    name = lambda g, k: next((r["label"] for r in vocab[g] if r["key"] == k), "")
    return jsonify({
        "medium": name("media", medium), "who": (name("actors", who["actor"]) or "未確認") + (f"（{who['label']}）" if who["label"] else ""),
        "method": name("methods", method) or "（当てはまらない）",
        "role": {"line": "線の方法にする", "both": "情報の箱にして、線の方法にもする", "box": "情報の箱にする"}[role],
        "handoff": vocab["handoff"][handoff]["label"] if handoff in vocab["handoff"] else "（決まらない）",
    })

# ----- デモデータ・エクスポート -----
@bp_admin.post("/demo/load")
def load_demo():
    n = seed_demo()
    flash(f"デモデータを投入しました（業務 {n} 件、社員マスタと月次アンケート4か月分を含む）。", "success")
    return redirect(url_for("main.index"))

@bp_admin.post("/demo/clear")
def remove_demo():
    kept = clear_demo()
    flash("デモデータを削除しました。", "info")
    if kept:
        who = "、".join(" ".join(x for x in (p["emp_id"], p["name"]) if x) + f"（{p['detail']}）" for p in kept)
        flash(f"デモの社員IDで本物の業務・回答がある人は消さずに残しました: {who}。対象者一覧のその行に印を付けています。回答率に入れないときは対象外にしてください。", "warning")
    return redirect(url_for("admin.index"))

@bp_admin.route("/export/respondents.csv")
def export_respondents():
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["社員ID", "氏名", "職場", "勤務形態", "自分で改善したい", "意欲（自動化を検討）", "月次の進め方（空でない直近の答え）", "その月度", "月次で自分で・一緒にと答えた回数",
                "使える技術", "学びたい技術", "登録業務数", "月間時間(分)", "一緒にやってみたい業務数", "伴走・希望の業務"])
    for p in dx_people():
        w.writerow(
            [
                p["emp_id"], p["name"], p["workplace"].name if p["workplace"] else "", p["work_pattern"], "○" if p["positive"] else "", wish_text(p),
                p["stance_latest"], stance_period_text(p), p["self_positive_count"] if p["monthly"] else "",
                "・".join(p["skills"]) or (C.SKILL_NONE if p["skills_none"] else ""), "・".join(p["learn"]) or (C.SKILL_NONE if p["learn_none"] else ""),
                p["task_count"], p["minutes"], p["checked_count"], "・".join(f"{x['task'].name}（{x['label']}）" for x in p["target_tasks"]),
            ]
        )
    return _csv_response(buf, "respondents_{}.csv".format(datetime.now().strftime("%Y%m%d")))

# ===== 画面: 月次アンケート =====
bp_monthly = Blueprint("monthly", __name__, url_prefix="/monthly")

_PERIOD_RE = re.compile(r"^\d{4}-(0[1-9]|1[0-2])$")
QUESTION_TYPES = ("choice", "text", "list")


def current_period():
    return datetime.now().strftime("%Y-%m")

def parse_period(value):
    value = (value or "").strip()
    return value if _PERIOD_RE.match(value) else current_period()

def answerable_periods():
    cur = current_period()
    prev = shift_period(cur, -1)
    periods = [cur]
    if MonthlySurvey.query.filter_by(period=prev).first() is not None:
        periods.append(prev)
    return periods

def public_period(value):
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

def default_questions():
    return [dict(q, options=list(q.get("options", []))) for q in C.MONTHLY_CORE_QUESTIONS]

def questions_for_new(period):
    prev = MonthlySurvey.query.filter(MonthlySurvey.period < period).order_by(MonthlySurvey.period.desc()).first()
    return prev.questions if prev else default_questions()

def get_or_create_survey(period):
    survey = MonthlySurvey.query.filter_by(period=period).first()
    if survey is None:
        survey = MonthlySurvey(period=period, is_active=True)
        survey.set_questions(questions_for_new(period))
        db.session.add(survey)
        db.session.commit()
    return survey

def _wants_json():
    return request.is_json or "application/json" in (request.headers.get("Accept") or "")

# ----- 回答（ログイン不要） -----
@bp_monthly.route("/")
def form():
    period = public_period(request.args.get("period"))
    survey = get_or_create_survey(period)
    emp_id = norm_emp_id(request.args.get("emp"))
    info = lookup(emp_id)
    existing = MonthlyResponse.query.filter(MonthlyResponse.period == period, same_emp_id(MonthlyResponse.emp_id, info["emp_id"])).first() if info else None
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
    try:
        workplace_id = int(data.get("workplace_id") or 0) or None
    except (TypeError, ValueError):
        workplace_id = None
    work_pattern = (data.get("work_pattern") or "").strip()
    info = lookup(data.get("emp_id"))
    emp_id = info["emp_id"] if info else norm_emp_id(data.get("emp_id"))
    name = info["name"] if info else ""

    errors = {}
    if not emp_id:
        errors["emp_id"] = "社員IDを入力してください"
    elif not info:
        errors["emp_id"] = EMP_NOT_FOUND
    elif not info["active"]:
        errors["emp_id"] = EMP_EXCLUDED
    if not workplace_id or db.session.get(Workplace, workplace_id) is None:
        errors["workplace_id"] = "職場を選んでください"
    if work_pattern not in C.WORK_PATTERNS:
        errors["work_pattern"] = "勤務形態を選んでください（普通勤務／交代勤務）"
    raw = data.get("answers") if request.is_json else None
    answers = {}
    for q in survey.questions:
        key = q["key"]
        if isinstance(raw, dict):
            value = raw.get(key)
        elif q["type"] == "list" and hasattr(data, "getlist"):
            value = data.getlist("ans-" + key)
        else:
            value = data.get("ans-" + key)
        if isinstance(value, list):
            value = "\n".join(str(v).strip() for v in value if str(v).strip())
        value = value.strip() if isinstance(value, str) else ""
        if q["type"] == "choice":
            if value and value not in q.get("options", []):
                errors[key] = "選択肢から選んでください"
            elif not value and q.get("required"):
                errors[key] = "選択してください"
        elif not value and q.get("required"):
            errors[key] = "入力してください"
        answers[key] = value[:1000]
    for q in survey.questions:
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
    if not errors and MonthlyResponse.query.filter(MonthlyResponse.period == period, same_emp_id(MonthlyResponse.emp_id, emp_id)).first():
        errors["answered"] = f"{period_label(period)}はすでに回答済みです（1人1回）"
    if errors:
        message = next(iter(errors.values()))
        if _wants_json():
            return jsonify({"success": False, "message": message, "errors": errors}), 400
        flash(message, "danger")
        return redirect(url_for("monthly.form", period=period, emp=emp_id or None))

    newest = period >= newest_period(emp_id)
    row = MonthlyResponse(
        period=period, emp_id=emp_id, name=name, workplace_id=workplace_id, work_pattern=work_pattern, survey_id=survey.id,
        answers_json=json.dumps(answers, ensure_ascii=False), comment=comment,
    )
    respondent = Respondent.query.filter(same_emp_id(Respondent.emp_id, emp_id)).order_by(Respondent.id).first()
    row.respondent_id = respondent.id if respondent else None
    if respondent is not None and newest:
        respondent.workplace_id = workplace_id
        respondent.work_pattern = work_pattern
    register_member(emp_id, name, workplace_id if newest else None, work_pattern if newest else "")
    db.session.add(row)
    db.session.commit()
    wants = answers.get("wants") == "ある"
    if _wants_json():
        session.pop("monthly_last", None)
        return jsonify(
            {
                "success": True,
                "message": f"{period_label(period)}の回答を受け付けました。ご協力ありがとうございました。",
                "wants": wants,
                "wants_tasks": monthly_task_names(answers.get("wants_tasks")),
                "register_url": url_for("survey.start", emp=emp_id, src="monthly", m=row.id),
            }
        )
    session["monthly_last"] = row.id
    return redirect(url_for("monthly.done", period=period))

@bp_monthly.route("/done")
def done():
    period = public_period(request.args.get("period"))
    last = session.pop("monthly_last", None)
    row = db.session.get(MonthlyResponse, last) if last else None
    if row is not None and row.period != period:
        row = MonthlyResponse.query.filter(MonthlyResponse.period == period, same_emp_id(MonthlyResponse.emp_id, row.emp_id)).first()
    wants = monthly_task_names(row.answer("wants_tasks")) if row is not None and row.answer("wants") == "ある" else []
    return render_template("monthly/done.html", period_label=period_label(period), row=row, wants_tasks=wants)

@bp_monthly.route("/api/employee/<emp_id>")
def api_employee(emp_id):
    period = public_period(request.args.get("period"))
    emp_id = norm_emp_id(emp_id)
    info = lookup(emp_id)
    if not info:
        return jsonify({"found": False, "emp_id": emp_id, "answered": False})
    answered = MonthlyResponse.query.filter(MonthlyResponse.period == period, same_emp_id(MonthlyResponse.emp_id, info["emp_id"])).first() is not None
    if not info["active"]:
        return jsonify({"found": True, "excluded": True, "message": EMP_EXCLUDED, "answered": answered, "emp_id": info["emp_id"]})
    return jsonify({"found": True, "answered": answered, "emp_id": info["emp_id"]})

# ----- 集計（要ログイン） -----
@bp_monthly.route("/results")
def results():
    rows = monthly_summary()
    surveys = MonthlySurvey.query.order_by(MonthlySurvey.period.desc()).all()
    default_period = rows[-1]["period"] if rows else (surveys[0].period if surveys else current_period())
    selected = parse_period(request.args.get("period") or default_period)
    survey = MonthlySurvey.query.filter_by(period=selected).first()
    responses = MonthlyResponse.query.filter_by(period=selected).order_by(MonthlyResponse.workplace_id, MonthlyResponse.name).all()
    stats = [s for s in (monthly_question_stats(survey, responses) if survey else []) if s["q"]["key"] not in C.MONTHLY_CORE_KEYS or s["q"]["type"] != "choice" or s["q"]["key"] == "hours"]
    summary = next((r for r in rows if r["period"] == selected), None)
    trends = monthly_question_trends()
    trends_json = [{"key": t["q"]["key"], "labels": t["labels"], "options": [x["option"] for x in t["series"]], "percents": [x["percents"] for x in t["series"]]} for t in trends]
    chart = {
        "labels": [r["label"] for r in rows],
        "n": [r["n"] for r in rows],
        **_rounded_series(rows, "hours_month", "has_rate", "decreased_rate", "wants_rate", "ask_rate", "self_rate", "together_rate", "self_only_rate"),
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
        wp_trends=monthly_workplace_trends(), pattern_trends=monthly_pattern_trends(),
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
    w.writerow(["月度", "社員ID", "氏名", "職場", "勤務形態"] + [texts[k] for k in keys] + ["コメント", "回答日時"])
    for r in MonthlyResponse.query.order_by(MonthlyResponse.period, MonthlyResponse.workplace_id, MonthlyResponse.name).all():
        a = r.answers
        w.writerow(
            [r.period, r.emp_id, r.name, r.workplace.name if r.workplace else "", r.work_pattern or ""]
            + [a.get(k, "") for k in keys]
            + [r.comment or "", r.created_at.strftime("%Y-%m-%d %H:%M") if r.created_at else ""]
        )
    return _csv_response(buf, "monthly_{}.csv".format(datetime.now().strftime("%Y%m%d")))

# ----- 設問の編集（要ログイン） -----
@bp_monthly.route("/admin")
def survey_admin():
    period = parse_period(request.args.get("period"))
    survey = MonthlySurvey.query.filter_by(period=period).first()
    surveys = MonthlySurvey.query.order_by(MonthlySurvey.period.desc()).all()
    counts = dict(db.session.query(MonthlyResponse.period, db.func.count(MonthlyResponse.id)).group_by(MonthlyResponse.period).all())
    return render_template(
        "monthly/admin.html",
        period=period,
        period_label=period_label(period),
        survey=survey,
        is_active=survey.is_active if survey is not None else True,
        questions_json=json.dumps(survey.questions if survey is not None else questions_for_new(period), ensure_ascii=False),
        core_keys=sorted(C.MONTHLY_CORE_KEYS),
        surveys=surveys,
        counts=counts,
        response_count=counts.get(period, 0),
        prev_period=shift_period(period, -1),
        current_period=current_period(),
    )

@bp_monthly.route("/admin/questions")
def survey_admin_questions():
    which = (request.args.get("period") or "default").strip()
    if which == "default":
        return jsonify({"questions": default_questions(), "label": "基本設問"})
    period = parse_period(which)
    survey = MonthlySurvey.query.filter_by(period=period).first()
    if survey is None:
        return jsonify({"questions": None, "message": f"{period_label(period)}の設問はありません"})
    return jsonify({"questions": survey.questions, "label": survey.period_label})

def _clean_questions(items):
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
    cleaned, _ = _core_in_order(cleaned)
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
    survey = _get_or_404(MonthlySurvey, survey_id)
    if MonthlyResponse.query.filter_by(period=survey.period).count():
        flash("回答がある月度の設問は削除できません。", "warning")
    else:
        db.session.delete(survey)
        db.session.commit()
        flash(f"{survey.period_label}の設問セットを削除しました。", "success")
    return redirect(url_for("monthly.survey_admin", period=current_period()))

# ===== 画面: 活動状況 =====
bp_board = Blueprint("board", __name__, url_prefix="/board")


@bp_board.route("/")
def index():
    workplace_id = request.args.get("workplace_id", type=int)
    q = Task.query
    if workplace_id:
        q = q.filter(Task.workplace_id == workplace_id)
    tasks = q.all()
    workplaces = _workplaces()
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

# ===== テンプレート =====
TEMPLATES = {
    'admin/employees.html': r'''{% extends "base.html" %}
{% block title %}対象者一覧{% endblock %}
{% block content %}
<nav aria-label="breadcrumb"><ol class="breadcrumb small mb-2">
  <li class="breadcrumb-item"><a href="{{ url_for('admin.index') }}">管理</a></li>
  <li class="breadcrumb-item active">対象者一覧</li>
</ol></nav>
<div class="d-flex flex-wrap justify-content-between align-items-end gap-2 mb-2">
  <div>
    <h2 class="mb-1">対象者一覧</h2>
    <p class="text-muted mb-0">月次アンケートか「自動化を検討」に一度でも回答した人が、社員IDごとに1行で自動で載ります。退職や他部門への異動で回答率の分母から外す人は「対象外」にチェックします。</p>
  </div>
  <div class="d-flex gap-2">
    <a class="btn btn-sm btn-outline-secondary" href="{{ url_for('monthly.results') }}">月次集計</a>
    <a class="btn btn-sm btn-outline-secondary" href="{{ url_for('admin.export_employees') }}">CSV</a>
  </div>
</div>

<div class="card mb-3"><div class="card-body">
  <div class="d-flex flex-wrap justify-content-between align-items-baseline gap-2 mb-2">
    <h5 class="card-title mb-0">対象者 <span id="member-active">{{ active_count }}</span> 人 <span class="small text-muted fw-normal">（対象外 <span id="member-excluded">{{ rows|length - active_count }}</span> 人）</span></h5>
    <span class="small text-muted"><i class="bi bi-check-lg member-yes"></i> その月に回答　<span class="member-before">－</span> 対象者になる前の月（初めて回答した月か、ここで追加した月より前。その月の分母に入りません）</span>
  </div>
  {% if rows %}
  <div class="member-wrap">
  <table class="table table-sm align-middle mb-0 member-table">
    <thead><tr>
      <th class="text-center">対象外</th><th>社員ID</th><th>氏名</th><th class="text-end">自動化を検討</th>
      {% for p in periods %}<th class="text-center member-month">{{ p[:4] }}<br>{{ p[5:]|int }}月</th>{% endfor %}
    </tr></thead>
    <tbody>
    {% for r in rows %}
    <tr class="{% if not r.emp.active %}is-excluded{% endif %}">
      <td class="text-center"><input class="form-check-input member-exclude" type="checkbox" data-url="{{ url_for('admin.toggle_employee', employee_id=r.emp.id) }}" {% if not r.emp.active %}checked{% endif %} aria-label="{{ r.emp.emp_id }} を対象外にする"></td>
      <td>{{ r.emp.emp_id }}{% set dr = demo_real.get(C.emp_key(r.emp.emp_id)) %}{% if dr %} <span class="badge bg-warning-subtle text-warning-emphasis" title="{{ dr.detail }}。デモデータを消してもこの人は残ります。回答率に入れないときは対象外にしてください">{{ dr.label }}</span>{% endif %}</td>
      <td>{{ r.emp.name or '' }}{% if r.deletable %}
        <form method="post" action="{{ url_for('admin.delete_employee', employee_id=r.emp.id) }}" class="d-inline" data-confirm="{{ r.emp.emp_id }} を対象者一覧から削除しますか？（まだ回答のない人だけ削除できます）"><button class="btn btn-link btn-sm text-danger p-0 ms-1">削除</button></form>{% endif %}</td>
      <td class="text-end">{{ r.considered }} 回</td>
      {% for p in periods %}<td class="text-center member-cell">{% if p in r.answered %}<i class="bi bi-check-lg member-yes" title="回答あり"></i>{% elif r.since and p < r.since %}<span class="member-before">－</span>{% endif %}</td>{% endfor %}
    </tr>
    {% endfor %}
    </tbody>
    <tfoot><tr>
      <th colspan="4" class="text-end">回答 ／ 対象（回答率）</th>
      {% for p in periods %}{% set st = stats[p] %}<td class="text-center member-stat" data-period="{{ p }}">{{ st.answered }}／{{ st.members }}{% if st.rate is not none %}<br>{{ '%.0f'|format(st.rate) }}%{% endif %}</td>{% endfor %}
    </tr></tfoot>
  </table>
  </div>
  <p class="small text-muted mt-2 mb-0">下の行の「回答」は対象者（対象外を除く）のうちその月に回答した人の数で、月次集計・ダッシュボードの回答率と同じ数え方です。対象外の人は回答できず、対象外にした月からは回答率に入りません（それより前の月はそのまま数えます）。削除できるのは、まだ回答も「自動化を検討」の登録もない人だけです。</p>
  {% else %}
  <p class="text-muted mb-0">まだ回答した人がいません。月次アンケートか「自動化を検討」に回答すると、ここに載ります。</p>
  {% endif %}
</div></div>

<div class="row g-3 mb-3">
  <div class="col-lg-6">
    <div class="card h-100"><div class="card-body">
      <h5 class="card-title">社員IDの照会（LDAP）</h5>
      <p class="mb-2">{{ ldap_status }}</p>
      <p class="small text-muted mb-0">回答するときは、社員IDが存在するかだけを確かめます。<code>config.py</code> に <code>LDAP_API_URL</code>（例: <code>http://ldap-api.local/employees/{id}</code>）と、レスポンスの項目名 <code>LDAP_API_FIELD_NAME</code> などを設定します（README 参照）。</p>
      <form class="d-flex gap-2 mt-3" onsubmit="return false">
        <input type="text" id="test-emp" class="form-control form-control-sm" placeholder="社員IDで照会テスト">
        <button type="button" class="btn btn-sm btn-outline-primary text-nowrap" id="test-btn">照会</button>
      </form>
      <div id="test-msg" class="form-text"></div>
    </div></div>
  </div>
  <div class="col-lg-6">
    <div class="card h-100"><div class="card-body">
      <h5 class="card-title">回答前に対象者へ入れる（任意）</h5>
      <p class="small text-muted">まだ回答していない人も分母に入れたいときだけ使います。</p>
      <form method="post" action="{{ url_for('admin.add_employee') }}" class="row g-2 align-items-end">
        <div class="col-6 col-md-3"><label class="form-label small mb-0">社員ID</label><input type="text" name="emp_id" class="form-control form-control-sm" required></div>
        <div class="col-6 col-md-3"><label class="form-label small mb-0">氏名</label><input type="text" name="name" class="form-control form-control-sm" required></div>
        <div class="col-6 col-md-3"><label class="form-label small mb-0">職場</label>
          <select name="workplace_id" class="form-select form-select-sm"><option value="">（未設定）</option>{% for w in workplaces %}<option value="{{ w.id }}">{{ w.name }}</option>{% endfor %}</select></div>
        <div class="col-6 col-md-3"><label class="form-label small mb-0">勤務形態</label>
          <select name="work_pattern" class="form-select form-select-sm"><option value="">（本人が選ぶ）</option>{% for w in C.WORK_PATTERNS %}<option value="{{ w }}">{{ w }}</option>{% endfor %}</select></div>
        <div class="col-12"><button class="btn btn-sm btn-primary">追加</button></div>
      </form>
      <form method="post" action="{{ url_for('admin.import_employees') }}" enctype="multipart/form-data" class="d-flex gap-2 align-items-center mt-3">
        <input type="file" name="file" accept=".csv,text/csv" class="form-control form-control-sm" required>
        <button class="btn btn-sm btn-outline-primary text-nowrap">CSV取込</button>
      </form>
      <p class="small text-muted mt-2 mb-0">CSV は1行目に「社員ID」「氏名」（任意で「所属」「メール」「勤務形態」）。同じ社員IDは上書きします。文字コードは UTF-8 または Shift_JIS。</p>
    </div></div>
  </div>
</div>
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
    msg.textContent = j.found ? ("見つかりました: " + j.emp_id + (j.excluded ? "（対象外。回答できません）" : "")) : "見つかりませんでした";
  } catch (e) { msg.textContent = "照会に失敗しました"; }
});
document.querySelectorAll(".member-exclude").forEach(function (box) {
  box.addEventListener("change", async function () {
    var row = box.closest("tr");
    box.disabled = true;
    try {
      var r = await fetch(box.dataset.url, { method: "POST", headers: { "Content-Type": "application/json", Accept: "application/json" }, body: JSON.stringify({ excluded: box.checked }) });
      var j = await r.json();
      if (!j.success) throw new Error("failed");
      row.classList.toggle("is-excluded", box.checked);
      document.getElementById("member-active").textContent = j.active_count;
      document.getElementById("member-excluded").textContent = j.excluded_count;
      j.stats.forEach(function (st) {
        var cell = document.querySelector('.member-stat[data-period="' + st.period + '"]');
        if (cell) cell.innerHTML = st.answered + "／" + st.members + (st.rate === null ? "" : "<br>" + Math.round(st.rate) + "%");
      });
    } catch (e) {
      box.checked = !box.checked;
      alert("保存できませんでした。通信状態を確認してください。");
    }
    box.disabled = false;
  });
});
</script>
{% endblock %}
''',
    'admin/map_vocab.html': r'''{% extends "base.html" %}
{% block title %}業務マップの区分{% endblock %}
{% block content %}
<nav aria-label="breadcrumb"><ol class="breadcrumb small mb-2">
  <li class="breadcrumb-item"><a href="{{ url_for('admin.index') }}">管理</a></li>
  <li class="breadcrumb-item active">業務マップの区分</li>
</ol></nav>
<h2 class="mb-1">業務マップの区分</h2>
<p class="text-muted">業務マップは「作業」「情報」の箱と、「仕事の流れ」「情報の流れ」の線だけで描きます。ここでは、作業のだれ・情報の媒体・情報の流れの方法・仕事の流れの引き継ぎ方の区分と、名前から見分けるための語を決めます。業務マップ・案件の流れ図・聞き取り中の図・編集画面の判定がそろって変わります（編集画面で手で選んだものは残ります）。</p>
<div class="alert alert-light border small">
  判定に使う語は「、」で区切ります。英字の語は単語として見分け、全角の英数字は半角とみなします。「～課」のように「～」で始めると語尾だけで見分けます（後ろが文末・かっこ・ひらがなのとき）。「除く語」の部分は判定に使いません（例: Word の除く語に「パスワード」）。上から順に見て、最初に当てはまった区分になります（▲▼で順番を変えられます）。<span class="badge text-bg-light border">組み込み</span> の行は判定と書き戻しに使うため消せませんが、名前や語は変えられます。
</div>
<div id="mv-app">
  <div class="card mb-3 mv-card"><div class="card-body">
    <h5 class="card-title">作業のだれ（担い手）</h5>
    <p class="small text-muted mb-2">作業の箱に付く「だれがやるか」。渡す相手・前後の作業の担い手を名前から見分けます。氏名は出さず、職場・役職・部署で表します。先に職場マスタの職場名で判定し、当たらなければここの語で判定します。「利用者の区分」は、業務データへ反映するときの利用者（職場内・上司・他部署・顧客・社外）です。</p>
    <div class="table-responsive"><table class="table table-sm align-middle mv-table" data-group="actors"><thead><tr><th></th><th>名前</th><th>判定に使う語</th><th>除く語</th><th>アイコン</th><th>利用者の区分</th><th></th></tr></thead><tbody></tbody></table></div>
    <button type="button" class="btn btn-sm btn-outline-primary mv-add" data-group="actors"><i class="bi bi-plus-lg"></i> 区分を追加</button>
  </div></div>
  <div class="card mb-3 mv-card"><div class="card-body">
    <h5 class="card-title">情報の媒体</h5>
    <p class="small text-muted mb-2">情報の箱の形とアイコンになる「どんな形の情報か」。どれにも当たらないものは「不明」になります。</p>
    <div class="table-responsive"><table class="table table-sm align-middle mv-table" data-group="media"><thead><tr><th></th><th>名前</th><th>判定に使う語</th><th>除く語</th><th>アイコン</th><th>形</th><th></th></tr></thead><tbody></tbody></table></div>
    <button type="button" class="btn btn-sm btn-outline-primary mv-add" data-group="media"><i class="bi bi-plus-lg"></i> 媒体を追加</button>
  </div></div>
  <div class="card mb-3 mv-card"><div class="card-body">
    <h5 class="card-title">情報の流れの方法</h5>
    <p class="small text-muted mb-2">情報の線のラベルになる「どうやって渡すか」。「渡し方」は線だけ、「置き場」（システム・共有フォルダ）は情報の箱を置いたうえで線の方法にもなります。既定の向きは仕事の流れの引き継ぎ方を決めるのに、箱がないときの媒体は口頭のやりとりなどで置く情報の箱に使います。</p>
    <div class="table-responsive"><table class="table table-sm align-middle mv-table" data-group="methods"><thead><tr><th></th><th>名前</th><th>判定に使う語</th><th>除く語</th><th>種類</th><th>既定の向き</th><th>箱がないときの媒体</th><th></th></tr></thead><tbody></tbody></table></div>
    <button type="button" class="btn btn-sm btn-outline-primary mv-add" data-group="methods"><i class="bi bi-plus-lg"></i> 方法を追加</button>
  </div></div>
  <div class="card mb-3 mv-card"><div class="card-body">
    <h5 class="card-title">仕事の流れの引き継ぎ方</h5>
    <p class="small text-muted mb-2">作業から次の作業へ仕事が移るときの2つのパターン。手順の文や渡し方にここの語（動詞）があれば、その引き継ぎ方になります。同じ人が続けて行う作業の間は「引き継ぎなし」です。</p>
    <div class="row g-2">
      <div class="col-md-6"><label class="form-label small mb-0">前の作業から伝えに行く</label><input type="text" class="form-control form-control-sm mb-1" id="mv-push-label" maxlength="20"><input type="text" class="form-control form-control-sm" id="mv-push-words" placeholder="判定に使う語（例: 渡す、送る）"></div>
      <div class="col-md-6"><label class="form-label small mb-0">後ろの作業が取りに行く</label><input type="text" class="form-control form-control-sm mb-1" id="mv-pull-label" maxlength="20"><input type="text" class="form-control form-control-sm" id="mv-pull-words" placeholder="判定に使う語（例: 集める、取りに）"></div>
    </div>
  </div></div>
  <div class="card mb-3 mv-card"><div class="card-body">
    <h5 class="card-title">判定のお試し</h5>
    <p class="small text-muted mb-2">名前や文を入れると、いま画面にある設定（保存前でも）でどう見分けるかを出します。</p>
    <div class="d-flex gap-2 flex-wrap"><input type="text" class="form-control form-control-sm" style="max-width: 28rem" id="mv-try" placeholder="例: 残業予定表（Excel）をメール送付、品質保証課 課長"><button type="button" class="btn btn-sm btn-outline-secondary text-nowrap" id="mv-try-btn">見分ける</button></div>
    <dl class="row small mt-2 mb-0 mv-try-out" id="mv-try-out"></dl>
  </div></div>
  <div class="mv-savebar">
    <button type="button" class="btn btn-primary" id="mv-save"><i class="bi bi-check2"></i> 保存</button>
    <button type="button" class="btn btn-outline-danger" id="mv-reset">既定に戻す</button>
    <span class="small" id="mv-msg" role="status"></span>
  </div>
</div>
<script type="application/json" id="mv-data">{{ vocab_json|safe }}</script>
<script type="application/json" id="mv-choices">{{ choices_json|safe }}</script>
{% endblock %}
{% block scripts %}
<script>
(function () {
  var V = JSON.parse(document.getElementById("mv-data").textContent);
  var CH = JSON.parse(document.getElementById("mv-choices").textContent);
  var msg = document.getElementById("mv-msg");
  var dirty = false;
  function touch() { dirty = true; msg.className = "small text-muted"; msg.textContent = "未保存の変更があります"; }
  function el(tag, cls, text) { var e = document.createElement(tag); if (cls) e.className = cls; if (text !== undefined) e.textContent = text; return e; }
  function select(map, value, onChange) {
    var s = el("select", "form-select form-select-sm");
    Object.keys(map).forEach(function (k) { var o = el("option", null, map[k]); o.value = k; if (k === (value || "")) o.selected = true; s.appendChild(o); });
    s.addEventListener("change", function () { onChange(s.value); touch(); });
    return s;
  }
  function input(value, ph, onInput, max) {
    var i = el("input", "form-control form-control-sm"); i.type = "text"; i.value = value || ""; if (ph) i.placeholder = ph; if (max) i.maxLength = max;
    i.addEventListener("input", function () { onInput(i.value); touch(); });
    return i;
  }
  function cell(tr, child, cls) { var td = el("td", cls); td.appendChild(child); tr.appendChild(td); return td; }
  function render(group) {
    var tbody = document.querySelector('.mv-table[data-group="' + group + '"] tbody');
    tbody.innerHTML = "";
    var rows = V[group];
    rows.forEach(function (r, i) {
      var tr = el("tr", "mv-row");
      var mv = el("div", "mv-move");
      var up = el("button", "btn btn-sm btn-light", "▲"); up.type = "button"; up.disabled = i === 0; up.title = "上へ（優先を上げる）";
      var dn = el("button", "btn btn-sm btn-light", "▼"); dn.type = "button"; dn.disabled = i === rows.length - 1; dn.title = "下へ";
      up.addEventListener("click", function () { rows.splice(i - 1, 0, rows.splice(i, 1)[0]); touch(); render(group); });
      dn.addEventListener("click", function () { rows.splice(i + 1, 0, rows.splice(i, 1)[0]); touch(); render(group); });
      mv.appendChild(up); mv.appendChild(dn);
      cell(tr, mv, "text-nowrap");
      var nameBox = el("div", "d-flex align-items-center gap-1");
      nameBox.appendChild(input(r.label, "名前", function (v) { r.label = v; }, 30));
      if (r.builtin) nameBox.appendChild(el("span", "badge text-bg-light border", "組み込み"));
      cell(tr, nameBox);
      cell(tr, input(r.words, "例: Excel、エクセル、xlsx", function (v) { r.words = v; }), "mv-words");
      cell(tr, input(r.exclude, "", function (v) { r.exclude = v; }));
      if (group !== "methods") {
        var wrap = el("div", "d-flex align-items-center gap-1");
        var icons = {}; CH.icons.forEach(function (k) { icons[k] = k.replace(/^bi-/, ""); });
        var prev = el("i", "bi " + r.icon + " mv-icon");
        wrap.appendChild(select(icons, r.icon, function (v) { r.icon = v; prev.className = "bi " + v + " mv-icon"; }));
        wrap.appendChild(prev);
        cell(tr, wrap);
      }
      if (group === "actors") cell(tr, select(CH.users, r.user, function (v) { r.user = v; }));
      if (group === "media") cell(tr, select(CH.shapes, r.shape, function (v) { r.shape = v; }));
      if (group === "methods") {
        cell(tr, select(CH.kinds, r.kind, function (v) { r.kind = v; }));
        cell(tr, select(CH.directions, r.direction, function (v) { r.direction = v; }));
        var media = {}; V.media.forEach(function (m) { media[m.key] = m.label; });
        cell(tr, select(media, r.medium, function (v) { r.medium = v; }));
      }
      var del = el("button", "btn btn-sm btn-outline-danger", "削除"); del.type = "button";
      del.disabled = !!r.builtin; if (r.builtin) del.title = "組み込みの行は消せません";
      del.addEventListener("click", function () { rows.splice(i, 1); touch(); render(group); });
      cell(tr, del);
      tbody.appendChild(tr);
    });
  }
  function renderAll() {
    ["actors", "media", "methods"].forEach(render);
    ["push", "pull"].forEach(function (k) {
      document.getElementById("mv-" + k + "-label").value = V.handoff[k].label;
      document.getElementById("mv-" + k + "-words").value = V.handoff[k].words;
    });
  }
  ["push", "pull"].forEach(function (k) {
    document.getElementById("mv-" + k + "-label").addEventListener("input", function (e) { V.handoff[k].label = e.target.value; touch(); });
    document.getElementById("mv-" + k + "-words").addEventListener("input", function (e) { V.handoff[k].words = e.target.value; touch(); });
  });
  document.querySelectorAll(".mv-add").forEach(function (b) {
    b.addEventListener("click", function () {
      var g = b.dataset.group;
      var row = { key: "", label: "", words: "", exclude: "" };
      if (g === "actors") { row.icon = "bi-person-fill"; row.user = ""; }
      if (g === "media") { row.icon = "bi-file-earmark"; row.shape = "doc"; }
      if (g === "methods") { row.kind = "channel"; row.direction = ""; row.medium = "unknown"; }
      V[g].push(row);
      touch();
      render(g);
      var inputs = document.querySelectorAll('.mv-table[data-group="' + g + '"] tbody tr:last-child input');
      if (inputs[0]) inputs[0].focus();
    });
  });
  async function post(url, body) {
    var r = await fetch(url, { method: "POST", headers: { "Content-Type": "application/json", Accept: "application/json" }, body: JSON.stringify(body) });
    return r.json();
  }
  document.getElementById("mv-try-btn").addEventListener("click", async function () {
    var t = document.getElementById("mv-try").value.trim();
    var out = document.getElementById("mv-try-out");
    out.innerHTML = "";
    if (!t) return;
    var j = await post("{{ url_for('admin.map_vocab_try') }}", { vocab: V, text: t });
    [["情報の箱なら 媒体", j.medium], ["作業の箱なら だれ", j.who], ["線なら 方法", j.method], ["手順の欄なら", j.role], ["引き継ぎ方", j.handoff]].forEach(function (p) {
      out.appendChild(el("dt", "col-sm-4 fw-normal text-muted", p[0]));
      out.appendChild(el("dd", "col-sm-8 mb-1", p[1]));
    });
  });
  document.getElementById("mv-try").addEventListener("keydown", function (e) { if (e.key === "Enter") { e.preventDefault(); document.getElementById("mv-try-btn").click(); } });
  document.getElementById("mv-save").addEventListener("click", async function () {
    try {
      var j = await post("{{ url_for('admin.map_vocab_save') }}", { vocab: V });
      if (!j.success) throw new Error("failed");
      V = j.vocab; dirty = false; renderAll();
      msg.className = "small text-success"; msg.textContent = "保存しました";
    } catch (e) { msg.className = "small text-danger"; msg.textContent = "保存できませんでした"; }
  });
  document.getElementById("mv-reset").addEventListener("click", async function () {
    if (!confirm("区分と判定の語を既定に戻しますか？（追加した区分も消えます）")) return;
    try {
      var j = await post("{{ url_for('admin.map_vocab_save') }}", { reset: true });
      V = j.vocab; dirty = false; renderAll();
      msg.className = "small text-success"; msg.textContent = "既定に戻しました";
    } catch (e) { msg.className = "small text-danger"; msg.textContent = "戻せませんでした"; }
  });
  window.addEventListener("beforeunload", function (e) { if (dirty) { e.preventDefault(); e.returnValue = ""; } });
  renderAll();
})();
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
      <div class="form-check mb-2">
        <input class="form-check-input" type="checkbox" name="board_show_names" value="1" id="board-show-names" {% if settings.board_show_names != '0' %}checked{% endif %}>
        <label class="form-check-label small" for="board-show-names">活動状況に登録者の氏名を表示する（外すと職場名のみ）</label>
      </div>
      <button class="btn btn-primary btn-sm">保存</button>
    </div></form>

    <div class="card mb-3"><div class="card-body">
      <h5 class="card-title">ログイン</h5>
      <p class="small mb-0">{{ auth_status }}</p>
    </div></div>

    <div class="card mb-3"><div class="card-body">
      <h5 class="card-title">業務マップの区分</h5>
      <p class="small mb-2">業務マップの「作業」「情報」「仕事の流れ」「情報の流れ」に付く区分（だれ・媒体・方法・引き継ぎ方）と、名前から見分けるための語を決めます。</p>
      <a class="btn btn-sm btn-outline-primary" href="{{ url_for('admin.map_vocab_page') }}"><i class="bi bi-diagram-3"></i> 業務マップの区分を設定</a>
    </div></div>

    <div class="card mb-3"><div class="card-body">
      <h5 class="card-title">対象者・月次アンケート</h5>
      <p class="small mb-1">回答した人 {{ answered_count }} 人（業務を1件以上登録したか、月次アンケートに回答した人。対象外は除く）。{{ ldap_status }}</p>
      <div class="d-flex flex-wrap gap-2">
        <a class="btn btn-sm btn-outline-primary" href="{{ url_for('admin.employees') }}">対象者一覧</a>
        <a class="btn btn-sm btn-outline-primary" href="{{ url_for('monthly.survey_admin') }}">月次の設問を編集</a>
        <a class="btn btn-sm btn-outline-secondary" href="{{ url_for('monthly.results') }}">月次集計</a>
      </div>
    </div></div>

    <form method="post" action="{{ url_for('admin.save_ai_settings') }}" class="card mb-3" autocomplete="off"><div class="card-body">
      <h5 class="card-title">AI連携（ChatGPT API）</h5>
      <p class="mb-2">{{ ai_status }}</p>
      <div class="mb-2"><label class="form-label small mb-0">API の URL（Responses API）</label><input type="text" name="openai_api_url" class="form-control form-control-sm" value="{{ settings.openai_api_url }}" placeholder="社内の中継サービスの Responses API の URL"></div>
      <div class="mb-2"><label class="form-label small mb-0">API キー</label><input type="text" name="openai_api_key" class="form-control form-control-sm" value="{{ settings.openai_api_key }}"></div>
      <div class="mb-2"><label class="form-label small mb-0">モデル名</label><input type="text" name="openai_model" class="form-control form-control-sm" value="{{ settings.openai_model }}" placeholder="社内で案内されている GPT-6 系の名前"></div>
      <div class="d-flex flex-wrap gap-2">
        <button class="btn btn-primary btn-sm">保存</button>
        <button class="btn btn-outline-secondary btn-sm" formaction="{{ url_for('admin.test_ai_settings') }}">保存して接続テスト</button>
      </div>
      <p class="small text-muted mt-2 mb-0">空欄のままなら AI は使わず、ルールベースで同じ形式の分析をします（既定の URL やモデル名はありません）。保存すればすぐ反映され、再起動は不要です。temperature / top_p は送りません。</p>
    </div></form>

    <div class="card mb-3"><div class="card-body">
      <h5 class="card-title">データ</h5>
      <p class="small mb-2">回答した人 {{ answered_count }} 人／業務 {{ task_count }} 件（うちデモ {{ demo_count }} 件）</p>
      <div class="d-flex flex-wrap gap-2">
        <form method="post" action="{{ url_for('admin.load_demo') }}" data-confirm="デモデータを投入します（既存のデモデータは入れ替えます）。よろしいですか？">
          <button class="btn btn-sm btn-outline-primary">デモデータを投入</button>
        </form>
        <form method="post" action="{{ url_for('admin.remove_demo') }}" data-confirm="デモデータを削除します。よろしいですか？">
          <button class="btn btn-sm btn-outline-danger" {% if not demo_active %}disabled{% endif %}>デモデータを削除</button>
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
      <p class="small text-muted">職場名は直して保存できます（業務・回答の紐づきはそのまま）。業務・月次の回答・対象者・「自動化を検討」の回答者が使っている職場は削除できません。</p>
      {% for w in workplaces %}
      {% set used = workplace_usage.get(w.id) %}
      <form method="post" action="{{ url_for('admin.update_workplace', workplace_id=w.id) }}" class="row g-1 align-items-center mb-2">
        <div class="col-8"><input type="text" name="name" class="form-control form-control-sm" value="{{ w.name }}" required aria-label="職場名"></div>
        <div class="col-4 d-flex gap-1 justify-content-end">
          <button class="btn btn-sm btn-outline-primary">保存</button>
          <button class="btn btn-sm btn-outline-danger" formaction="{{ url_for('admin.delete_workplace', workplace_id=w.id) }}" {% if used %}disabled title="使われている職場は削除できません"{% endif %} onclick="return confirm('「{{ w.name }}」を削除しますか？')">削除</button>
        </div>
        <div class="col-12 small text-muted">{{ used|join('・') if used else '未使用（削除できます）' }}</div>
      </form>
      {% endfor %}
      <form method="post" action="{{ url_for('admin.add_workplace') }}" class="d-flex gap-2 mt-2">
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
        <li class="nav-item"><a class="nav-link {% if bp == 'survey' %}active{% endif %}" href="{{ url_for('survey.index') }}">自動化を検討</a></li>
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
  {% if demo_active %}
  <div class="alert alert-warning demo-banner py-2 small" role="status"><i class="bi bi-cone-striped"></i> <strong>デモデータを表示中</strong>：架空の人・業務・回答が入っています。数字や一覧は実際の活動ではありません。{% if current_user %}<a href="{{ url_for('admin.index') }}" class="alert-link">管理</a>の「デモデータを削除」で消せます。{% endif %}</div>
  {% endif %}
  {% with messages = get_flashed_messages(with_categories=true) %}
    {% set toasts = [] %}
    {% for category, message in messages %}
      {% if category in ('success', 'info', 'secondary', 'message') %}
        {% set _ = toasts.append((category, message)) %}
      {% else %}
      <div class="alert alert-{{ category }} alert-dismissible fade show" role="alert">
        {{ message }}
        <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="閉じる"></button>
      </div>
      {% endif %}
    {% endfor %}
    {% if toasts %}
    <div class="ui-toasts" role="status" aria-live="polite">
      {% for category, message in toasts %}
      <div class="alert alert-{{ 'secondary' if category == 'message' else category }} alert-dismissible fade show" role="alert">
        <i class="bi {{ 'bi-check-circle-fill' if category == 'success' else 'bi-info-circle-fill' }}"></i>{{ message }}
        <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="閉じる"></button>
      </div>
      {% endfor %}
    </div>
    {% endif %}
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
      <p class="small text-muted mb-0">取り出し元の情報 → 業務 → できるもの → 渡す相手の作業。業務の箱をクリックすると案件詳細を開きます。</p>
      {{ info_legend|safe }}
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
      <button type="button" class="btn btn-sm btn-outline-primary" id="fe-add-work" title="作業の箱を足す"><i class="bi bi-plus-square"></i> 作業</button>
      <button type="button" class="btn btn-sm btn-outline-primary" id="fe-add-info" title="情報の箱を足す"><i class="bi bi-file-earmark-plus"></i> 情報</button>
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
      <button type="button" class="btn btn-sm btn-outline-dark" id="fe-max" title="図を画面いっぱいに表示（Esc で戻る）"><i class="bi bi-arrows-fullscreen"></i> 最大表示</button>
      <a class="btn btn-sm btn-outline-dark" href="{{ url_for('bizmap.index', workplace_id=workplace_id, domain_id=domain_id) }}">閲覧へ戻る</a>
    </div>
  </div>
  <div class="fe-hint">
    操作: ドラッグで移動 ／ 文字をダブルクリックで名前を編集 ／ 🔗 で線を引く（相手をクリック。作業どうしは仕事の流れ、作業と情報は情報の流れ。情報どうしは結べません） ／ 線をクリックで方法・引き継ぎ方・向き・削除 ／ 🎨 で種類・だれ・媒体・色 ／ Shift+ドラッグで範囲選択 ／ 空白をドラッグで画面移動 ／ ホイールで拡大縮小 ／ Delete で選択を削除
  </div>
  <div class="map-legend fe-kind-legend" id="fe-kind-legend"></div>
  <div id="fe-mode" class="fe-mode"></div>
  <div class="fe-container" id="fe-container">
    <div class="fe-canvas" id="fe-canvas">
      <svg class="fe-svg" id="fe-svg">
        <defs>
          <marker id="fe-arrow-info" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse"><path d="M0 0L10 5L0 10z" fill="#5b8def"/></marker>
          <marker id="fe-arrow-work" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="6" markerHeight="6" orient="auto-start-reverse"><path d="M0 0L10 5L0 10z" fill="#495057"/></marker>
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
      <div class="mt-2 fe-type-row" id="fe-type-row">
        <span class="small me-2">種類</span>
        <div class="btn-group btn-group-sm" role="group"><button type="button" class="btn btn-outline-secondary" data-type="work"><i class="bi bi-square"></i> 作業</button><button type="button" class="btn btn-outline-secondary" data-type="info"><i class="bi bi-file-earmark"></i> 情報</button></div>
      </div>
      <div class="mt-2" id="fe-who-row">
        <label class="small mb-0" for="fe-who-input">だれ（職場・役職・部署。氏名は書かない）</label>
        <input type="text" class="form-control form-control-sm mb-1" id="fe-who-input" placeholder="例: 製造1課 課長、生産管理課">
        <select class="form-select form-select-sm" id="fe-actor-select" aria-label="だれの区分"></select>
      </div>
      <div class="mt-2" id="fe-medium-row">
        <label class="small mb-0" for="fe-medium-select">媒体</label>
        <select class="form-select form-select-sm" id="fe-medium-select"></select>
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
      <div class="fe-pop-title">線の設定 <span class="small text-muted fw-normal" id="fe-edge-type"></span></div>
      <div class="mb-2" id="fe-method-row"><label class="small mb-0" for="fe-method-select">方法（渡し方）</label><select class="form-select form-select-sm" id="fe-method-select"></select></div>
      <div class="mb-2" id="fe-handoff-row"><label class="small mb-0" for="fe-handoff-select">引き継ぎ方</label><select class="form-select form-select-sm" id="fe-handoff-select"></select></div>
      <input type="text" class="form-control form-control-sm" id="fe-edge-label" placeholder="メモ（書くとラベルになります）">
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
<script type="application/json" id="fe-kinds">{{ kinds_json|safe }}</script>
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
      <div>{{ legend|safe }}{{ kind_legends.get(key, '')|safe }}</div>
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
<p class="text-muted">登録された業務から、仕事の流れ（作業 → 作業）と情報の流れ（情報 → 作業 → 情報）を1枚の図に自動生成します。案件が増えるほど製造部の業務全体像が描かれていきます。間違いは「この図を編集する」で手で直せ、直した内容は保存され、新しい業務を取り込んで再生成しても残ります。</p>

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
      図に載せた業務 {{ task_count }} 件（関わる人・情報の受け取り元の入力あり {{ flow_info_count }} 件）{% if pending_count %}・月次で挙がった未登録の業務 {{ pending_count }} 件{% endif %}
    </div>
  </div>
</div></form>
{% if truncated %}
<div class="alert alert-warning py-2 small">業務が {{ max_tasks }} 件を超えるため、先頭 {{ max_tasks }} 件だけを描いています。職場や領域で絞り込んでください。</div>
{% endif %}

<div class="bizmap-sections">
  <div id="tab-info">
    {{ flow_panel('info', '業務マップ',
       '取り出し元の情報 → 業務 → できるもの → 渡す相手の作業。箱は作業（四角・だれつき）と情報（媒体の形とアイコン）の2種類、線は仕事の流れ（太い灰）と情報の流れ（細い青）の2種類です。聞き取りで答えた名前で描き、同じ名前は1つの箱にまとめます。職場ごとの枠に業務を置き、点線は名前から推定したつながり（ある業務のできるものを別の業務が使っている、前後の業務）です。灰色の点線の箱は、月次アンケートで挙がったものの、まだ流れを登録していない業務です。保留中の業務は点線の枠に「（保留）」、廃止した業務は薄い色に「（廃止済み）」を付けます。',
       '',
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
              <td class="text-nowrap">{{ kind_icon(r.name, r.hint)|safe }}{% if r.cls == 'system' %}<span class="legend-chip lg-system">{{ r.name }}</span>{% else %}{{ r.name }}{% endif %}</td>
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
          <p class="small text-muted">誰から受け取り、誰に渡しているか。「情報」は情報の受け取り元とアウトプットの利用者から、「関与」は関わる人・部署から数えています。業務マップと同じ決まりで、受け取り元は1つずつに分け、役職は職場名つきにし、「他部署」だけの相手は目的の文の部署名にします。</p>
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
      <h5 class="card-title">業務領域ごとの把握状況（ヒートマップ） <span class="badge text-bg-light border align-middle">全体</span></h5>
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
          <h5 class="card-title">同じアウトプットを複数人が作成（統合候補） <span class="badge text-bg-light border align-middle">全体</span></h5>
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
          <h5 class="card-title">「同じ情報を何度も入力する」業務（入力元別） <span class="badge text-bg-light border align-middle">全体</span></h5>
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
          <h5 class="card-title">業務のつながり（前の業務 → 業務 → 後の業務） <span class="badge text-bg-light border align-middle">全体</span></h5>
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
    'board/_board.html': r'''
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
    <div class="pb-cta-text">あなたのPCの定型・定期業務も、自動化を検討してみませんか</div>
    <div class="pb-cta-actions">
      <a class="pb-cta-btn" href="{{ url_for('survey.index') }}">自動化を検討する</a>
      <a class="pb-cta-link" href="{{ url_for('monthly.form') }}">月次アンケートに答える</a>
    </div>
  </div>
</header>

{% if hl %}
<section class="pb-wins" aria-label="これまでの成果と検討への入口">
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
    'board/_macros.html': r'''
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
      ・ 月間 <strong>{{ task.monthly_minutes|hours }} 時間</strong>（{{ C.FREQUENCY_HINTS.get(task.frequency, task.frequency) if task.frequency == '不定期' else task.frequency }}／{{ task.duration }}／{{ task.headcount }}人{% if task.shift %}／{{ task.shift }}{% endif %}）
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
      <h5 class="card-title">検討内容（登録された業務情報）</h5>
      <table class="table table-sm mb-0">
        {% set aif = [] if task.verified else task.ai_fields_list %}
        {% set rulef = task.ai_extra_dict.get('rule_fields') or [] %}
        {% macro mark(key) %}{% if key in aif %}<span class="badge badge-flag me-1">{{ '推定' if key in rulef else 'AI推定' }}</span>{% endif %}{% endmacro %}
        {% if task.steps %}<tr><th class="w-25">手順（聞き取り）</th><td><ol class="mb-0 ps-3 rq-steps-table">{% for x in task.steps %}<li>{{ x.action }}{% if x.system %} <span class="text-muted small">［{{ x.system }}］</span>{% endif %}{% if x.output %} <span class="small">→ {{ x.output }}</span>{% endif %}{% if x.to %} <span class="small text-muted">（{{ x.to }} へ{% if x.how %}・{{ x.how }}{% endif %}）</span>{% endif %}</li>{% endfor %}</ol></td></tr>{% endif %}
        <tr><th class="w-25">業務領域</th><td>{{ mark('domain') }}{{ task.domain.name if task.domain else '未選択' }}{% if task.domain_guess %} <span class="badge badge-flag">AI推定: {{ task.domain_guess }}</span>{% endif %}</td></tr>
        <tr><th>目的</th><td>{{ mark('purpose') }}{{ task.purpose }}{% if task.purpose_unknown %} <span class="badge text-bg-danger">目的不明</span>{% endif %}</td></tr>
        <tr><th>アウトプット</th><td>{{ mark('output') }}{{ task.output }}</td></tr>
        <tr><th>利用者</th><td>{{ mark('output_users') or mark('output_users_note') }}{{ task.output_users_list|join('・') }}{% if task.output_users_note %}<div class="small text-muted">{{ task.output_users_note }}</div>{% endif %}</td></tr>
        <tr><th>入力情報</th><td>{{ mark('inputs') or mark('inputs_note') }}{{ task.inputs_list|join('・') }}{% if task.inputs_note %}<div class="small text-muted">{{ task.inputs_note }}</div>{% endif %}</td></tr>
        <tr><th>前後の業務</th><td>{{ mark('prev_task') or mark('next_task') }}{{ task.prev_task or '－' }} → <strong>{{ task.name }}</strong> → {{ task.next_task or '－' }}</td></tr>
        <tr><th>関わる人・部署</th><td>{{ mark('actors') }}{{ task.actors_list|join('・') or '－' }}</td></tr>
        <tr><th>情報の受け取り元</th><td>{{ mark('info_from') }}{{ task.info_from or '－' }}</td></tr>
        <tr><th>担当人数</th><td>{{ task.headcount or 1 }}人{% if task.ai_extra_dict.get('headcount_guess') %} <span class="small text-muted">（参考: AI の推定では {{ task.ai_extra_dict.headcount_guess }}人。月間時間には入れていません）</span>{% endif %}</td></tr>
        <tr><th>使用ツール</th><td>{{ mark('tools') }}{{ task.tools_list|join('・') }}</td></tr>
        <tr><th>作業の性質</th><td>{{ mark('nature') }}{{ task.nature_list|join('・') }}</td></tr>
        <tr><th>困りごと</th><td>{{ mark('pains') }}{{ task.pains_list|join('・') }}{% if task.pains_note %}<div class="small text-muted">{{ task.pains_note }}</div>{% endif %}</td></tr>
        <tr><th>改善の希望</th><td>{{ task.wish or '' }}</td></tr>
        <tr><th>自分で改善</th><td>{{ task.self_improve or C.SELF_IMPROVE_UNCHECKED }}</td></tr>
        {% if task.note %}<tr><th>補足</th><td>{{ task.note }}</td></tr>{% endif %}
      </table>
      {% if task.intake %}
      <details class="small mt-2"><summary class="text-muted">聞き取りの記録（{{ task.intake|length }} 件）</summary>
        <div class="rq-log mt-1">{% for m in task.intake %}<div class="{{ 'rq-log-ai' if m.role == 'ai' else 'rq-log-user' }}"><span class="text-muted">{{ 'AI' if m.role == 'ai' else '本人' }}:</span> {{ m.text }}</div>{% endfor %}</div>
      </details>
      {% endif %}
    </div></div>

    <div class="card mb-3"><div class="card-body">
      <div class="d-flex justify-content-between align-items-start flex-wrap gap-2">
        <div>
          <h5 class="mb-1">この業務の流れ</h5>
          <p class="small text-muted mb-0">取り出し元の情報 → 手順 → できるもの → 渡す相手の作業、前後の業務。<a href="{{ url_for('bizmap.index', workplace_id=task.workplace_id) }}">職場全体の流れ図</a></p>
          {{ task_legend|safe }}
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
          {% if task.ai_source == 'openai' %}<span class="badge text-bg-primary">ChatGPT API</span>{% elif task.ai_source == 'claude' %}<span class="badge text-bg-primary">Claude API</span>{% elif task.ai_source %}<span class="badge text-bg-secondary">ルールベース</span>{% endif %}
          {% if task.ai_priority %}<span class="badge text-bg-{{ {'高':'danger','中':'warning','低':'light text-muted'}.get(task.ai_priority, 'light') }}">優先度 {{ task.ai_priority }}</span>{% endif %}
        </h5>
        <form method="post" action="{{ url_for('cases.reanalyze', task_id=task.id) }}">
          <button class="btn btn-sm btn-outline-primary">{% if ai_enabled %}AIで再分析{% else %}再分析（ルール）{% endif %}</button>
        </form>
      </div>
      {% if task.ai_stale_note %}
      <div class="small text-muted mt-2"><span class="badge text-bg-light border">{{ task.ai_stale_note }}</span> 業務情報を直したあとは分析していません。新しくするには「{% if ai_enabled %}AIで再分析{% else %}再分析（ルール）{% endif %}」を押してください。</div>
      {% endif %}
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
      {% set extra = task.ai_extra_dict %}
      {% if extra.get('unknowns') %}
      <div class="small mt-2"><strong>聞き取りで分からなかったこと（担当者に確認）</strong>
        <ul class="mb-0">{% for q in extra.unknowns %}<li>{{ q }}</li>{% endfor %}</ul>
      </div>
      {% endif %}
      {% set plan_note = task.ai_plan_note %}
      {% if extra.get('tobe_steps') or extra.get('saving_minutes_per_month') or extra.get('hearing_questions') %}
      <div class="small mt-2{% if plan_note %} pt-2 border-top{% endif %}">
        {% if plan_note == 'デモ' %}<div class="text-muted mb-1"><span class="badge bg-warning-subtle text-warning-emphasis">デモ</span> デモ用の見本の案です（AI で分析したものではありません）。</div>
        {% elif plan_note %}<div class="text-muted mb-1"><span class="badge text-bg-light border">{{ plan_note }}</span> 前回の AI 分析で出た案です。いまの分析（ルールベース）とは別で、再分析では更新していません。</div>{% endif %}
        {% if extra.get('tobe_steps') %}
        <strong>改善後の流れ（AI案）</strong>
        <ol class="mb-1 ps-3">{% for x in extra.tobe_steps %}<li>{{ x }}</li>{% endfor %}</ol>
        {% endif %}
        {% if extra.get('saving_minutes_per_month') %}<div class="text-success">削減の見込み 約 {{ extra.saving_minutes_per_month }} 分/月</div>{% endif %}
        {% if extra.get('hearing_questions') %}
        <div class="mt-2"><strong>ヒアリングで聞くこと（まずは相談したい）</strong>
          <ul class="mb-0">{% for q in extra.hearing_questions %}<li>{{ q }}</li>{% endfor %}</ul>
        </div>
        {% endif %}
      </div>
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
        <li><a href="{{ url_for('cases.detail', task_id=r.id) }}">{{ r.name }}</a> <span class="badge badge-flag">{{ 'AI' if task.ai_source in ('openai', 'claude') else '推定' }}</span> <span class="small text-muted">{{ r.workplace.name if r.workplace else '' }}／{{ r.monthly_minutes|hours }} h</span></li>
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

      {% set missing = task.missing_required %}
      {% macro need(label) %}{% if label in missing %} <span class="badge text-bg-warning">未入力</span>{% endif %}{% endmacro %}
      <label class="form-label small mb-1 mt-2" id="case-info"><strong>業務情報</strong>（直すと推定の印が外れます。必須項目が空なら <span class="badge text-bg-warning">未入力</span> が付きます）</label>
      <div class="row g-2 mb-2">
        <div class="col-8">
          <label class="form-label small mb-0" for="f-domain_id">業務領域{{ need('業務領域') }}</label>
          <select name="domain_id" id="f-domain_id" class="form-select form-select-sm">
            <option value="">未選択</option>
            {% for d in domains %}<option value="{{ d.id }}" {% if task.domain_id == d.id %}selected{% endif %}>{{ d.name }}</option>{% endfor %}
          </select>
        </div>
        <div class="col-4">
          <label class="form-label small mb-0">担当人数</label>
          <input type="number" name="headcount" class="form-control form-control-sm" min="1" step="1" value="{{ task.headcount or 1 }}">
        </div>
        {% if task.ai_extra_dict.get('headcount_guess') %}<div class="col-12 small text-muted mt-0">AI の推定は {{ task.ai_extra_dict.headcount_guess }}人（参考）。同じ作業をしている人数が分かったら担当人数に入れてください（月間時間に掛けます）。</div>{% endif %}
        <div class="col-12">
          <label class="form-label small mb-0" for="f-purpose">目的（なぜやるのか・誰が何に使うか）{{ need('目的') }}</label>
          <textarea name="purpose" id="f-purpose" class="form-control form-control-sm" rows="2" maxlength="500">{{ task.purpose or '' }}</textarea>
        </div>
        <div class="col-12">
          <label class="form-label small mb-0" for="f-output">アウトプット（最終的にできるもの）{{ need('アウトプット') }}</label>
          <input type="text" name="output" id="f-output" class="form-control form-control-sm" value="{{ task.output or '' }}" maxlength="500">
        </div>
        <div class="col-12" id="f-output_users">
          <div class="form-label small mb-0">利用者{{ need('利用者') }}</div>
          <input type="hidden" name="output_users" value="">
          {% for u in C.OUTPUT_USERS %}<div class="form-check form-check-inline small mb-0"><input class="form-check-input" type="checkbox" name="output_users" value="{{ u }}" id="f-ou-{{ loop.index }}" {% if u in task.output_users_list %}checked{% endif %}><label class="form-check-label" for="f-ou-{{ loop.index }}">{{ u }}</label></div>{% endfor %}
          <input type="text" name="output_users_note" class="form-control form-control-sm mt-1" value="{{ task.output_users_note or '' }}" maxlength="500" placeholder="渡す相手の職場・部署・役職（「、」区切り。氏名は書かない）" aria-label="利用者の補足">
        </div>
        <div class="col-12" id="f-inputs">
          <div class="form-label small mb-0">入力元（情報のもと）{{ need('入力元') }}</div>
          <input type="hidden" name="inputs" value="">
          {% for s in C.INPUT_SOURCES %}<div class="form-check form-check-inline small mb-0"><input class="form-check-input" type="checkbox" name="inputs" value="{{ s }}" id="f-in-{{ loop.index }}" {% if s in task.inputs_list %}checked{% endif %}><label class="form-check-label" for="f-in-{{ loop.index }}">{{ s }}</label></div>{% endfor %}
          <input type="text" name="inputs_note" class="form-control form-control-sm mt-1" value="{{ task.inputs_note or '' }}" maxlength="500" placeholder="具体的なシステム名・ファイル名・帳票名（「、」区切り）" aria-label="入力元の補足">
        </div>
        <div class="col-6">
          <label class="form-label small mb-0" for="f-frequency">頻度{{ need('頻度') }}</label>
          <select name="frequency" id="f-frequency" class="form-select form-select-sm">
            {% if task.frequency not in C.FREQUENCIES %}<option value="">未選択</option>{% endif %}
            {% for v in C.FREQUENCIES %}<option value="{{ v }}" {% if task.frequency == v %}selected{% endif %}>{{ C.FREQUENCY_HINTS.get(v, v) }}</option>{% endfor %}
          </select>
        </div>
        <div class="col-6">
          <label class="form-label small mb-0" for="f-duration">1回の時間{{ need('時間') }}</label>
          <select name="duration" id="f-duration" class="form-select form-select-sm">
            {% if task.duration not in C.DURATIONS %}<option value="">未選択</option>{% endif %}
            {% for v in C.DURATIONS %}<option value="{{ v }}" {% if task.duration == v %}selected{% endif %}>{{ v }}</option>{% endfor %}
          </select>
        </div>
        <div class="col-12 small text-muted mt-0">頻度・1回の時間・担当人数を変えると、保存したときに月間時間を計算し直します。</div>
        <div class="col-6">
          <label class="form-label small mb-0">前の業務</label>
          <input type="text" name="prev_task" class="form-control form-control-sm" value="{{ task.prev_task or '' }}" maxlength="500">
        </div>
        <div class="col-6">
          <label class="form-label small mb-0">後の業務</label>
          <input type="text" name="next_task" class="form-control form-control-sm" value="{{ task.next_task or '' }}" maxlength="500">
        </div>
        <div class="col-12">
          <label class="form-label small mb-0">関わる人・部署（「、」区切り。氏名は書かない）</label>
          <input type="text" name="actors" class="form-control form-control-sm" value="{{ task.actors or '' }}" maxlength="500" placeholder="例：課長、生産管理課">
        </div>
      </div>

      <label class="form-label small mb-1 mt-2"><strong>標準手順</strong>（自動化は最後）</label>
      {% if task.steps_lagging %}
      <div class="alert alert-info py-1 small mb-2">段階は「{{ task.stage_name }}」まで進んでいます。標準手順の {% for no, name, desc in C.STEPS if no in task.steps_lagging %}{{ no }}. {{ name }}{% if not loop.last %}・{% endif %}{% endfor %} にまだチェックがありません。済んでいればチェックして保存してください。</div>
      {% endif %}
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
      <div class="alert alert-warning py-1 small">必須項目（{% for key, label in C.REQUIRED_FIELDS if label in task.missing_required %}<a href="#f-{{ key }}" class="alert-link">{{ label }}</a>{% if not loop.last %}・{% endif %}{% endfor %}）が未入力です。ヒアリングで確かめて、上の「案件管理」の<a href="#case-info" class="alert-link">業務情報</a>に入れて「保存する」を押してください。</div>
      {% endif %}
      {% if task.verified %}
        <div class="small mb-2">確認者: {{ task.verified_by or '－' }}</div>
        <button class="btn btn-outline-secondary btn-sm w-100">確認を解除する</button>
      {% else %}
        <div class="mb-2">
          <input type="text" name="verified_by" class="form-control form-control-sm" placeholder="確認者名（推進メンバー）">
        </div>
        {% if task.domain_guess %}
        <div class="form-check mb-2">
          <input class="form-check-input" type="checkbox" name="adopt_ai_domain" id="adopt" value="1">
          <label class="form-check-label small" for="adopt">業務領域をAI推定（{{ task.domain_guess }}）に{{ '変更する' if task.domain else 'する' }}</label>
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
    <a class="btn btn-primary btn-sm" href="{{ url_for('survey.index') }}">自動化を検討する</a>
  </div>
</div>
<p class="text-muted small">登録された業務はすべて改善案件です。希望が「なくしたい・減らしたい」か、困りごとのある業務を「案件候補」として数えます。</p>

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
      <label class="form-label small mb-0">改善の希望</label>
      <select name="wish" class="form-select form-select-sm">
        <option value="">すべて</option>
        {% for w in C.WISHES %}<option value="{{ w }}" {% if args.get('wish') == w %}selected{% endif %}>{{ w }}</option>{% endfor %}
      </select>
    </div>
    <div class="col-6 col-md-4 col-xl">
      <label class="form-label small mb-0">絞り込み</label>
      <select name="filter" class="form-select form-select-sm">
        <option value="">すべて</option>
        <option value="candidates" {% if args.get('filter') == 'candidates' %}selected{% endif %}>案件候補（なくしたい・減らしたい・困りごとあり）</option>
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
    <th>業務名</th><th>職場／登録者</th><th>領域</th><th class="text-end">月間</th><th>改善の希望</th><th>自分で</th><th>ステータス</th><th>進捗（活動状況）</th><th>手順</th><th>確認</th><th>AI</th>
  </tr></thead>
  <tbody>
  {% for t in tasks %}
  <tr>
    <td><a href="{{ url_for('cases.detail', task_id=t.id) }}">{{ t.name }}</a>{% if t.shift and t.shift != '普通勤務' %} <span class="badge text-bg-dark">{{ t.shift }}</span>{% endif %}{% if t.is_demo %} <span class="badge bg-warning-subtle text-warning-emphasis">デモ</span>{% endif %}<div class="small text-muted">{{ t.output }}</div></td>
    <td class="small">{{ t.workplace.name if t.workplace else '' }}<br>{{ t.respondent.name if t.respondent else '' }}</td>
    <td class="small">{{ t.domain.name if t.domain else '' }}</td>
    <td class="text-end">{{ t.monthly_minutes|hours }} h</td>
    <td class="small">{{ t.wish or '' }}</td>
    <td class="small">{% if t.self_improve in C.SELF_IMPROVE_POSITIVE %}<span class="badge text-bg-success">希望</span>{% else %}{{ t.self_improve or C.SELF_IMPROVE_UNCHECKED }}{% endif %}</td>
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
  まだ業務が登録されていません。<a href="{{ url_for('monthly.form') }}">月次アンケート</a>・<a href="{{ url_for('survey.index') }}">自動化を検討</a>から始めるか、
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
      <div class="stat-sub">自ら改善に取り組みたい人（回答した人 {{ k.respondent_count }} 人中）</div>
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

<h2 class="section-title">現状の把握 <span class="section-sub">自動化を検討（棚卸し）から</span></h2>
<div class="row g-3 mb-4">
  <div class="col-6 col-lg-3">
    <div class="card stat h-100"><div class="card-body">
      <div class="stat-head"><span class="stat-tile tile-teal"><i class="bi bi-list-check"></i></span><span class="stat-label">PCの定型・定期業務の月間時間</span></div>
      <div class="stat-value">{{ '%.1f'|format(k.total_hours) }}<span class="stat-unit">時間/月</span></div>
      <div class="stat-sub">年間換算 {{ '%.0f'|format(k.annual_hours) }} 時間 ・ <a href="{{ url_for('cases.index', filter='candidates') }}">案件候補 {{ k.candidate_count }} 件</a>（なくしたい・減らしたい・困りごとあり）</div>
    </div></div>
  </div>
  <div class="col-6 col-lg-3">
    <div class="card stat h-100"><div class="card-body">
      <div class="stat-head"><span class="stat-tile tile-pink"><i class="bi bi-envelope-paper"></i></span><span class="stat-label">月次アンケート回答率</span></div>
      <div class="stat-value">{% if k.response_rate is not none %}{{ '%.0f'|format(k.response_rate) }}<span class="stat-unit">%</span>{% else %}－{% endif %}</div>
      <div class="stat-sub">{% if k.response_period %}{{ k.response_period }}：回答 {{ k.response_answered }} 人 ／ 対象 {{ k.response_members }} 人{% else %}回答はまだありません{% endif %}</div>
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
      <div class="small text-muted mb-2">最新 {{ monthly_latest.label }}：回答 {{ monthly_latest.n }} 人{% if monthly_latest.member_answered != monthly_latest.n %}（うち対象者 {{ monthly_latest.member_answered }} 人）{% endif %}{% if monthly_latest.response_rate is not none %}・回答率 {{ '%.0f'|format(monthly_latest.response_rate) }}%（{{ monthly_latest.member_answered }}／{{ monthly_latest.members }} 人）{% endif %}{% if monthly_latest.pattern_counts %}・{% for k, v in monthly_latest.pattern_counts.items() %}{{ k }} {{ v }}人{% if not loop.last %}／{% endif %}{% endfor %}{% endif %}</div>
      <div class="row g-2 text-center mb-2">
        <div class="col-4"><div class="kpi-label">PCの定型・定期業務あり</div><div class="mini-value">{{ '%.0f'|format(monthly_latest.has_rate) }}%</div></div>
        <div class="col-4"><div class="kpi-label">先月より減った</div><div class="mini-value text-green">{{ '%.0f'|format(monthly_latest.decreased_rate) }}%</div></div>
        <div class="col-4"><div class="kpi-label">自分で・一緒に取り組みたい</div><div class="mini-value text-blue">{{ '%.0f'|format(monthly_latest.self_rate) }}%</div></div>
      </div>
      <div class="chart-box" style="height: 150px"><canvas id="chart-monthly"></canvas></div>
      {% else %}
      <p class="text-muted small mb-2">毎月同じ設問でPCの定型・定期業務の変化を追います。まだ回答がありません。</p>
      <a class="btn btn-sm btn-outline-secondary" href="{{ url_for('monthly.form') }}">回答フォーム</a>
      {% endif %}
    </div></div>
  </div>
  <div class="col-lg-7">
    <div class="card h-100"><div class="card-body">
      <h5 class="card-title">職場別の回答状況</h5>
      <div class="table-responsive">
      <table class="table table-sm mb-0">
        <thead><tr><th>職場</th><th class="text-end">回答した人</th><th class="text-end">業務</th><th class="text-end">月間時間</th><th class="text-end" title="希望が「なくしたい・減らしたい」か、困りごとのある業務">案件候補</th><th class="text-end">自主改善</th></tr></thead>
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
      <p class="small text-muted mt-2 mb-0">回答した人は、業務を1件以上登録したか、月次アンケートに回答した人（対象外は除く）を、いまの職場（月次アンケートか「自動化を検討」のいちばん新しい回答で選んだ職場）で1人1回数えます。自主改善はそのうち自ら改善に取り組みたい人です。業務・月間時間・案件候補は、業務を登録したときの職場で数えます。</p>
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
          <button class="btn btn-sm btn-outline-primary" {% if not ai_enabled %}disabled title="管理画面の AI連携で URL・API キー・モデル名を設定すると使えます"{% endif %}>AIで提案を生成</button>
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
      <p class="small text-muted mt-2 mb-0">AI連携は未設定です。管理画面の「AI連携」で URL・API キー・モデル名を設定すると、案件登録時のAI分析とここでの提案生成が有効になります。</p>
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
          <td><a href="{{ url_for('cases.detail', task_id=t.id) }}">{{ t.name }}</a>{% if t.is_demo %} <span class="badge bg-warning-subtle text-warning-emphasis">デモ</span>{% endif %}</td>
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
  Chart.defaults.animation.duration = 900;
  Chart.defaults.animation.easing = "easeOutQuart";
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
        { label: "PCの定型・定期業務あり", data: data.monthly.has_rate, borderColor: C.red, backgroundColor: C.red, tension: 0.35, pointRadius: 3, borderWidth: 2.5 },
        { label: "減った", data: data.monthly.decreased_rate, borderColor: C.green, backgroundColor: C.green, tension: 0.35, pointRadius: 3, borderWidth: 2.5 },
        { label: "自分で・一緒に取り組みたい", data: data.monthly.self_rate, borderColor: C.blue, backgroundColor: C.blue, tension: 0.35, pointRadius: 3, borderWidth: 2.5 }
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
      {% if survey is none %}<div class="alert alert-info py-2 small">この月度の設問セットはまだありません。直近の月度（なければ基本設問）を表示しています。「保存する」を押すと作られます。</div>{% endif %}
      <div id="qe-list"></div>
      <button type="button" class="btn btn-outline-primary w-100 mt-2" id="qe-add"><i class="bi bi-plus-lg me-1"></i>設問を追加</button>
      <hr>
      <div class="d-flex flex-wrap justify-content-between align-items-center gap-2">
        <div class="form-check form-switch mb-0">
          <input class="form-check-input" type="checkbox" id="qe-active" {% if is_active %}checked{% endif %}>
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
      <p class="small text-muted mt-2 mb-0">月度の設問セットは、回答フォームが初めて開かれたときに直近の月度からコピーして自動で作られます。この編集画面は開いただけでは作らず、保存したときに作ります。</p>
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
              <select class="form-select qe-type"><option value="choice">選択（1つ）</option><option value="text">記述</option><option value="list">記述（複数）</option></select>
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
    node.querySelector(".qe-options").classList.toggle("d-none", type !== "choice");
    node.querySelector(".qe-placeholder").classList.toggle("d-none", type === "choice");
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
  <div class="card mq-card mq-done" data-celebrate><div class="card-body text-center">
    <div class="mq-check"><i class="bi bi-check-lg"></i></div>
    <h2 class="h4 mb-1">回答を受け付けました</h2>
    <p class="text-muted mb-3">{{ period_label }}のアンケートにご協力ありがとうございました。</p>
    {% if wants_tasks %}
    <div class="mq-next">
      <div class="mq-next-title">次は、なくしたい業務の「流れ」を AI と整理してみませんか</div>
      <div class="mq-next-tasks">{% for n in wants_tasks %}<span>{{ n }}</span>{% endfor %}</div>
      <p class="small text-muted mb-3">あなたのことと業務名はアンケートの答えから入力済みです。AI の短い質問に答えると流れ図ができ、改善チームと自動化を検討できます（1件 2〜3分）。</p>
      <a class="btn btn-primary btn-lg" href="{{ url_for('survey.start', emp=row.emp_id, src='monthly', m=row.id) }}"><i class="bi bi-diagram-3"></i> 自動化を検討する</a>
    </div>
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
    <p class="text-muted mb-3">{{ period_label }}のアンケートはすでに回答いただいています。ご協力ありがとうございました。</p>
    <button type="button" class="btn btn-outline-secondary btn-sm" id="mq-other-id">別の社員IDで回答する</button>
  </div></div>

  <form id="mq-form" class="card mq-card {% if existing %}d-none{% endif %}" method="post" action="{{ url_for('monthly.submit') }}" novalidate>
    <div class="card-body">
      <input type="hidden" name="period" value="{{ period }}">
      <div class="mq-progress" id="mq-progress" aria-live="polite"><div class="mq-progress-bar"><span></span></div><span class="mq-progress-text">あと少しです</span></div>
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
              {% for w in workplaces %}<option value="{{ w.id }}">{{ w.name }}</option>{% endfor %}
            </select>
          </div>
          <div class="col-sm-6">
            <label class="form-label" for="work_pattern">勤務形態</label>
            <select name="work_pattern" id="work_pattern" class="form-select" required>
              <option value="">選択</option>
              {% for w in C.WORK_PATTERNS %}<option value="{{ w }}">{{ w }}</option>{% endfor %}
            </select>
          </div>
        </div>
        <div id="emp-msg" class="form-text"></div>
      </div>

      {% set ns = namespace(main=0, nos={}, subs={}, labels={}) %}
      {% for q in questions %}
      {% set is_sub = q.depends and not q.depends.main %}
      {% if is_sub %}
        {% set grp = q.depends.key ~ '|' ~ ((q.depends.show_if or q.depends.required_if or [])|join('/')) %}
        {% set _ = ns.subs.update({grp: ns.subs.get(grp, 1) + 1}) %}
        {% set label = ns.nos.get(q.depends.key, '?') ~ '-' ~ ns.subs[grp] %}
      {% else %}
        {% set ns.main = ns.main + 1 %}
        {% set _ = ns.nos.update({q.key: ns.main}) %}
        {% set label = ns.main %}
      {% endif %}
      {% set _ = ns.labels.update({q.key: label}) %}
      <div class="mq-q{% if is_sub %} mq-sub{% endif %}{% if q.depends %} d-none{% endif %}" data-key="{{ q.key }}" data-type="{{ q.type }}" data-required="{{ 1 if q.required else 0 }}"{% if q.depends %} data-dep-key="{{ q.depends.key }}" data-dep-show='{{ (q.depends.show_if or q.depends.required_if or [])|tojson }}' data-dep-required="{{ 1 if (q.depends.required or q.depends.required_if) else 0 }}"{% endif %}>
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
        {% elif q.type == 'list' %}
        {% set copy_from = C.MONTHLY_COPY_FROM.get(q.key) %}
        <div class="mq-list-input" data-list-key="{{ q.key }}">
          <div class="mq-list-row"><input type="text" name="ans-{{ q.key }}" class="form-control" data-list-item maxlength="100" placeholder="{{ q.placeholder or '' }}"><button type="button" class="mq-list-del" aria-label="この行を消す">×</button></div>
          <div class="mq-list-row"><input type="text" name="ans-{{ q.key }}" class="form-control" data-list-item maxlength="100" placeholder="2つ目（あれば）"><button type="button" class="mq-list-del" aria-label="この行を消す">×</button></div>
          <div class="mq-list-actions">
            <button type="button" class="btn btn-sm btn-outline-primary mq-list-add"><i class="bi bi-plus-lg"></i> 業務を追加</button>
            {% if copy_from and ns.labels.get(copy_from) %}<button type="button" class="btn btn-sm btn-outline-secondary mq-list-copy" data-from="{{ copy_from }}"><i class="bi bi-arrow-down-short"></i> {{ ns.labels[copy_from] }} で答えた業務名を転記</button>{% endif %}
          </div>
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
    <div class="mq-next d-none" id="mq-next">
      <div class="mq-next-title">次は、なくしたい業務の「流れ」を AI と整理してみませんか</div>
      <div class="mq-next-tasks" id="mq-next-tasks"></div>
      <p class="small text-muted mb-3">あなたのことと業務名はアンケートの答えから入力済みです。AI の短い質問に答えると流れ図ができ、改善チームと自動化を検討できます（1件 2〜3分）。</p>
      <a class="btn btn-primary btn-lg" id="mq-register"><i class="bi bi-diagram-3"></i> 自動化を検討する</a>
    </div>
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
  var wpSelect = document.getElementById("workplace_id");
  var patSelect = document.getElementById("work_pattern");
  var msg = document.getElementById("emp-msg");
  var lookupTimer = null, empFound = null;
  var NOT_FOUND = {{ C.EMP_NOT_FOUND|tojson }}, empProblem = NOT_FOUND;
  function setMsg(text, state) { msg.textContent = text || ""; msg.classList.toggle("mq-emp-ok", state === "ok"); msg.classList.toggle("mq-emp-ng", state === "ng"); }
  async function lookupEmp() {
    var id = empInput.value.trim();
    empFound = null;
    if (!id) { setMsg(""); return; }
    var r = await fetch("{{ url_for('monthly.api_employee', emp_id='') }}" + encodeURIComponent(id) + "?period={{ period }}", { headers: { Accept: "application/json" } });
    var j = await r.json();
    if (empInput.value.trim() !== id) return;
    if (j.answered && !j.excluded) {
      form.classList.add("d-none");
      answered.classList.remove("d-none");
      return;
    }
    answered.classList.add("d-none");
    form.classList.remove("d-none");
    empFound = !!j.found && !j.excluded;
    if (empFound) {
      setMsg("社員IDを確認しました", "ok");
    } else {
      empProblem = j.excluded ? (j.message || {{ C.EMP_EXCLUDED|tojson }}) : NOT_FOUND;
      setMsg(empProblem, "ng");
    }
    updateProgress();
  }
  document.getElementById("mq-other-id").addEventListener("click", function () {
    empInput.value = "";
    empFound = null;
    setMsg("");
    answered.classList.add("d-none");
    form.classList.remove("d-none");
    if (window.history && history.replaceState) history.replaceState(null, "", "{{ url_for('monthly.form', period=period) }}");
    updateProgress();
    empInput.focus();
  });
  var prog = document.getElementById("mq-progress");
  function answeredOf(q) {
    if (q.dataset.type === "choice") return !!q.querySelector("input:checked");
    if (q.dataset.type === "list") return Array.prototype.some.call(q.querySelectorAll("[data-list-item]"), function (i) { return i.value.trim(); });
    var t = q.querySelector("textarea");
    return !!(t && t.value.trim());
  }
  function updateProgress() {
    if (!prog) return;
    var items = [!!(empInput.value.trim() && empFound !== false && wpSelect.value && patSelect.value)];
    form.querySelectorAll(".mq-q[data-key]").forEach(function (q) {
      if (!q.classList.contains("d-none") && q.dataset.required === "1") items.push(answeredOf(q));
    });
    var done = items.filter(Boolean).length, rest = items.length - done;
    prog.style.setProperty("--p", items.length ? done / items.length : 0);
    prog.classList.toggle("is-complete", rest === 0);
    form.classList.toggle("mq-ready", rest === 0);
    prog.querySelector(".mq-progress-text").textContent = rest === 0 ? "すべて答えました。送信できます" : "あと " + rest + " 問";
  }
  function applyDeps() {
    form.querySelectorAll(".mq-q[data-dep-key]").forEach(function (q) {
      var parent = form.querySelector('.mq-q[data-key="' + q.dataset.depKey + '"] input:checked');
      var val = parent ? parent.value : "";
      var shown = JSON.parse(q.dataset.depShow || "[]").indexOf(val) >= 0;
      var required = shown && q.dataset.depRequired === "1";
      q.classList.toggle("d-none", !shown);
      q.dataset.required = required ? "1" : "0";
      if (!shown) {
        q.querySelectorAll('textarea, input[type="text"]').forEach(function (t) { t.value = ""; });
        q.querySelectorAll('input[type="radio"]').forEach(function (i) { i.checked = false; });
      }
      var req = q.querySelector('[data-badge="req"]'), opt = q.querySelector('[data-badge="opt"]');
      if (req) req.classList.toggle("d-none", !required);
      if (opt) opt.classList.toggle("d-none", required || !shown);
      var err = q.querySelector(".mq-err");
      if (err && !shown) err.textContent = "";
    });
  }
  form.addEventListener("change", function (e) { if (e.target.matches('input[type="radio"]')) applyDeps(); updateProgress(); });
  form.addEventListener("input", function () { updateProgress(); });
  function listBox(el) { return el.closest(".mq-list-input"); }
  function listValues(box) { return Array.prototype.map.call(box.querySelectorAll("[data-list-item]"), function (i) { return i.value.trim(); }).filter(Boolean); }
  function addListRow(box, value) {
    var first = box.querySelector("[data-list-item]");
    var row = document.createElement("div");
    row.className = "mq-list-row";
    var inp = document.createElement("input");
    inp.type = "text"; inp.name = first.name; inp.className = "form-control"; inp.maxLength = 100;
    inp.setAttribute("data-list-item", "");
    inp.placeholder = (box.querySelectorAll("[data-list-item]").length + 1) + "つ目（あれば）";
    if (value) inp.value = value;
    var del = document.createElement("button");
    del.type = "button"; del.className = "mq-list-del"; del.setAttribute("aria-label", "この行を消す"); del.textContent = "×";
    row.appendChild(inp); row.appendChild(del);
    box.insertBefore(row, box.querySelector(".mq-list-actions"));
    return inp;
  }
  form.addEventListener("click", function (e) {
    var add = e.target.closest(".mq-list-add");
    if (add) { addListRow(listBox(add)).focus(); return; }
    var del = e.target.closest(".mq-list-del");
    if (del) {
      var box = listBox(del), row = del.closest(".mq-list-row");
      if (box.querySelectorAll(".mq-list-row").length > 1) row.remove(); else row.querySelector("[data-list-item]").value = "";
      return;
    }
    var copy = e.target.closest(".mq-list-copy");
    if (copy) {
      var target = listBox(copy);
      var source = form.querySelector('.mq-list-input[data-list-key="' + copy.dataset.from + '"]');
      var names = source ? listValues(source) : [];
      var err = target.closest(".mq-q").querySelector(".mq-err");
      if (!names.length) { err.textContent = "先に上の設問に業務名を入れてください"; return; }
      err.textContent = "";
      var have = listValues(target);
      names.forEach(function (n) {
        if (have.indexOf(n) >= 0) return;
        var empty = Array.prototype.filter.call(target.querySelectorAll("[data-list-item]"), function (i) { return !i.value.trim(); })[0];
        if (empty) empty.value = n; else addListRow(target, n);
        have.push(n);
      });
    }
  });
  form.addEventListener("keydown", function (e) {
    if (e.key !== "Enter" || e.isComposing || !e.target.matches("[data-list-item]")) return;
    e.preventDefault();
    var box = listBox(e.target);
    var items = Array.prototype.slice.call(box.querySelectorAll("[data-list-item]"));
    var next = items[items.indexOf(e.target) + 1];
    if (next) next.focus(); else addListRow(box).focus();
  });
  applyDeps();
  updateProgress();
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
      else if (q.dataset.type === "list") { val = Array.prototype.map.call(q.querySelectorAll("[data-list-item]"), function (i) { return i.value.trim(); }).filter(Boolean).join("\n"); }
      else { var t = q.querySelector("textarea"); val = t ? t.value.trim() : ""; }
      var err = q.querySelector(".mq-err"); err.textContent = "";
      if (q.dataset.required === "1" && !val) { err.textContent = q.dataset.type === "choice" ? "選択してください" : "入力してください"; if (!firstMissing) firstMissing = q; }
      answers[key] = val;
    });
    var missing = !form.emp_id.value.trim() ? "社員IDを入力してください" : empFound === false ? empProblem : !form.workplace_id.value ? "職場を選んでください" : !form.work_pattern.value ? "勤務形態を選んでください" : "";
    if (missing) {
      errBox.textContent = missing; errBox.classList.remove("d-none");
      document.querySelector(".mq-identity").scrollIntoView({ behavior: "smooth", block: "center" });
      return;
    }
    if (firstMissing) { firstMissing.scrollIntoView({ behavior: "smooth", block: "center" }); return; }
    var payload = { period: form.period.value, emp_id: form.emp_id.value.trim(), workplace_id: form.workplace_id.value, work_pattern: form.work_pattern.value, answers: answers, comment: form.comment.value.trim() };
    submitBtn.disabled = true; submitBtn.textContent = "送信中…";
    try {
      var r = await fetch(form.action, { method: "POST", headers: { "Content-Type": "application/json", Accept: "application/json" }, body: JSON.stringify(payload) });
      var j = await r.json();
      if (j.success) {
        form.classList.add("d-none");
        document.getElementById("mq-success-msg").textContent = j.message;
        if (j.wants) {
          document.getElementById("mq-register").href = j.register_url;
          var tl = document.getElementById("mq-next-tasks");
          tl.innerHTML = "";
          (j.wants_tasks || []).forEach(function (n) { var sp = document.createElement("span"); sp.textContent = n; tl.appendChild(sp); });
          document.getElementById("mq-next").classList.remove("d-none");
        }
        document.getElementById("mq-success").classList.remove("d-none");
        window.scrollTo({ top: 0, behavior: "smooth" });
        if (window.uiCelebrate) window.uiCelebrate();
      } else {
        errBox.textContent = j.message || "送信できませんでした"; errBox.classList.remove("d-none");
        if (j.errors && j.errors.emp_id) { empFound = false; empProblem = j.errors.emp_id; setMsg(empProblem, "ng"); updateProgress(); }
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
    <a class="btn btn-outline-primary btn-sm" href="{{ url_for('admin.employees') }}">対象者一覧</a>
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
          <th>月度</th><th class="text-end">回答者</th><th class="text-end">回答率</th><th class="text-end">PCの定型・定期業務あり</th><th class="text-end">推計 月間時間</th>
          <th class="text-end">減った</th><th class="text-end">増えた</th><th class="text-end">自動化したい</th><th class="text-end">依頼したい</th><th class="text-end">一緒に取り組みたい</th><th class="text-end">自分で取り組みたい</th><th class="text-end">新規業務</th><th class="text-end">コメント</th>
        </tr></thead>
        <tbody>
        {% for r in rows %}
        <tr class="{% if r.period == selected %}table-primary{% endif %}">
          <td><a href="{{ url_for('monthly.results', period=r.period) }}#detail">{{ r.label }}</a></td>
          <td class="text-end">{{ r.n }}{% if r.member_answered != r.n %}<div class="small text-muted">うち対象者 {{ r.member_answered }}</div>{% endif %}</td>
          <td class="text-end">{% if r.response_rate is not none %}{{ '%.0f'|format(r.response_rate) }}%<div class="small text-muted">{{ r.member_answered }}／{{ r.members }}</div>{% else %}－{% endif %}</td>
          <td class="text-end">{{ '%.0f'|format(r.has_rate) }}%</td>
          <td class="text-end">{{ '%.0f'|format(r.hours_month) }} h</td>
          <td class="text-end">{{ '%.0f'|format(r.decreased_rate) }}%</td>
          <td class="text-end">{{ '%.0f'|format(r.increased_rate) }}%</td>
          <td class="text-end">{{ '%.0f'|format(r.wants_rate) }}%</td>
          <td class="text-end">{{ '%.0f'|format(r.ask_rate) }}%</td>
          <td class="text-end">{{ '%.0f'|format(r.together_rate) }}%</td>
          <td class="text-end">{{ '%.0f'|format(r.self_only_rate) }}%</td>
          <td class="text-end">{{ r.new_count }}</td>
          <td class="text-end">{{ r.comment_count }}</td>
        </tr>
        {% endfor %}
        </tbody>
      </table>
      </div>
      <p class="small text-muted mt-2 mb-0">減った・増えた・自動化したいは ① で「ある」と答えた人のうちの割合、依頼したい・一緒に取り組みたい・自分で取り組みたいは「なくしたい業務がある」と答えた人のうちの割合です（その設問に答えた人が分母）。回答率は、その月までに対象者一覧に載った人（初めて回答した月か、管理画面で追加した月から。「対象外」にした人を除く）のうち、その月に回答した人の割合です（下の小さい数字が 回答した対象者／対象者）。「うち対象者」は回答者に対象外の人が含まれるときだけ出ます。推計月間時間 ＝ 1シフトあたり時間の換算値 × {{ C.SHIFTS_PER_MONTH }}（1か月の出勤日数の目安）。</p>
    </div></div>
  </div>
</div>

<h2 class="section-title">設問ごとの推移 <span class="section-sub">選択肢の割合（%）が月度ごとにどう動いたか</span></h2>
<div class="row g-3 mb-2">
  {% for t in trends %}
  <div class="col-lg-6">
    <div class="card h-100"><div class="card-body">
      <h5 class="card-title mb-0">{{ t.q.text }}</h5>
      <div class="small text-muted mb-2">{{ t.label }}{% if t.note %}<br>{{ t.note }}{% endif %}</div>
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

<h2 class="section-title">職場別の推移 <span class="section-sub">PCの定型・定期業務あり率と推計月間時間</span></h2>
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
  <p class="small text-muted mt-2 mb-0">上段がPCの定型・定期業務あり率、下段が推計月間時間と回答者数です。</p>
</div></div>

<h2 class="section-title">勤務形態別の推移 <span class="section-sub">普通勤務と交代勤務の違い</span></h2>
<div class="card mb-2"><div class="card-body">
  <div class="table-responsive">
  <table class="table table-sm trend-table mb-0">
    <thead><tr><th>勤務形態</th>{% for l in pattern_trends.labels %}<th class="text-end">{{ l|replace('年', '/')|replace('月度', '') }}</th>{% endfor %}</tr></thead>
    <tbody>
    {% for row in pattern_trends.rows %}
    <tr>
      <td>{{ row.pattern }}</td>
      {% for c in row.cells %}
      <td class="text-end">{% if c.n %}<div>{{ '%.0f'|format(c.has_rate) }}%</div><div class="cell-sub">{{ '%.0f'|format(c.hours_month) }} h・{{ c.n }}人</div>{% else %}<span class="text-muted">－</span>{% endif %}</td>
      {% endfor %}
    </tr>
    {% else %}
    <tr><td colspan="{{ 1 + pattern_trends.labels|length }}" class="text-muted">回答はまだありません</td></tr>
    {% endfor %}
    </tbody>
  </table>
  </div>
  <p class="small text-muted mt-2 mb-0">勤務形態は本人が初回に選んだものです。「未設定」は勤務形態を聞く前の回答です。</p>
</div></div>
{% endif %}

<h2 class="section-title" id="detail">月度の詳細
  <form method="get" class="d-inline-flex align-items-center ms-1">
    <select name="period" class="form-select form-select-sm" onchange="this.form.submit()" style="width: auto">
      {% for s in surveys %}<option value="{{ s.period }}" {% if s.period == selected %}selected{% endif %}>{{ s.period_label }}</option>{% endfor %}
      {% if not surveys %}<option value="{{ selected }}">{{ selected_label }}</option>{% endif %}
    </select>
  </form>
  <span class="section-sub">{% if summary %}回答 {{ summary.n }} 人{% if summary.member_answered != summary.n %}（うち対象者 {{ summary.member_answered }} 人）{% endif %}{% if summary.response_rate is not none %}・回答率 {{ '%.0f'|format(summary.response_rate) }}%（{{ summary.member_answered }}／{{ summary.members }} 人）{% endif %}{% else %}回答はまだありません{% endif %}</span>
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
          {% for opt in s.options %}
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
  Chart.defaults.plugins.tooltip.backgroundColor = "#1d1d1f";
  Chart.defaults.plugins.tooltip.padding = 10;
  Chart.defaults.plugins.tooltip.cornerRadius = 8;
  Chart.defaults.animation.duration = 900;
  Chart.defaults.animation.easing = "easeOutQuart";
  var line = function (label, data, color) { return { label: label, data: data, borderColor: color, backgroundColor: color, tension: 0.35, pointRadius: 3, borderWidth: 2.5 }; };
  new Chart(document.getElementById("chart-rates"), {
    type: "line",
    data: { labels: d.labels, datasets: [
      line("PCの定型・定期業務あり", d.has_rate, "#ff453a"), line("減った", d.decreased_rate, "#30d158"), line("自動化したい", d.wants_rate, "#ff9f0a"),
      line("改善チームに依頼したい", d.ask_rate, "#bf5af2"), line("一緒に取り組みたい", d.together_rate, "#0a84ff"), line("自分で取り組みたい", d.self_only_rate, "#ff9f0a")
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
<p class="text-muted">「自動化を検討」で選んだ「スキルを身につけながら、一緒にやってみたい」と使える技術・学びたい技術、月次アンケートの改善の進め方から、一人ひとりの意欲をまとめます。</p>

<div class="row g-3 mb-3">
  <div class="col-6 col-lg-3"><div class="card kpi-card h-100"><div class="card-body py-2">
    <div class="kpi-label">回答した人</div>
    <div class="kpi-value fs-3">{{ summary.total }}<span class="fs-6 text-muted"> 人</span></div>
    <div class="kpi-sub">業務を1件以上登録したか、月次アンケートに回答した人（対象外は除く）</div>
  </div></div></div>
  <div class="col-6 col-lg-3"><div class="card kpi-card h-100"><div class="card-body py-2">
    <div class="kpi-label">自分で改善したい</div>
    <div class="kpi-value fs-3">{{ summary.positive }}<span class="fs-6 text-muted"> 人</span></div>
    <div class="kpi-sub">検討での「一緒にやってみたい」のチェックと、直近6か月の月次の進め方（「自分で取り組みたい」「学びながら改善チームと一緒に取り組みたい」）のうち、新しいほうで判断（チェックなしは意思表示なしとして扱う）。伴走している「自主改善」の案件がある人も含む</div>
    {% if summary.positive_by_pattern %}<div class="kpi-sub">{% for k, v in summary.positive_by_pattern.items() %}{{ k }} {{ v }}人{% if not loop.last %}／{% endif %}{% endfor %}</div>{% endif %}
  </div></div></div>
  <div class="col-6 col-lg-3"><div class="card kpi-card h-100"><div class="card-body py-2">
    <div class="kpi-label">直近の月次で「自分で・一緒に取り組みたい」</div>
    <div class="kpi-value fs-3">{{ summary.self_now }}<span class="fs-6 text-muted"> 人</span></div>
    <div class="kpi-sub">{{ summary.period_label }}</div>
  </div></div></div>
  <div class="col-6 col-lg-3"><div class="card kpi-card h-100"><div class="card-body py-2">
    <div class="kpi-label">学びたい技術を挙げた人</div>
    <div class="kpi-value fs-3">{{ summary.learners }}<span class="fs-6 text-muted"> 人</span></div>
    <div class="kpi-sub">勉強会・伴走の対象（下の「学びたい技術」）</div>
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
      {% for opt in C.DX_WISH_OPTIONS %}<option value="{{ opt }}" {% if args.wish == opt %}selected{% endif %}>{{ opt }}</option>{% endfor %}
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
      <th>氏名</th><th>職場</th><th>勤務形態</th><th>意欲（自動化を検討）</th><th>月次：改善の進め方</th>
      <th>使える技術</th><th>学びたい技術</th><th>業務</th>
    </tr></thead>
    <tbody>
    {% for p in people %}
    <tr>
      <td class="text-nowrap">{{ p.name }}{% if p.emp_id %}<div class="small text-muted">{{ p.emp_id }}</div>{% endif %}</td>
      <td class="small text-nowrap">{{ p.workplace.name if p.workplace else '' }}</td>
      <td class="small text-nowrap">{{ p.work_pattern or '－' }}</td>
      <td class="small text-nowrap">
        {% if p.wish %}<span class="badge {{ 'text-bg-success' if p.wish_positive else ('text-bg-secondary' if p.wish_note else 'text-bg-light border') }}">{{ p.wish }}</span>{% if p.wish_note %}<div class="text-muted">{{ p.wish_note }}{% if p.wish_hint %} <span class="badge bg-warning-subtle text-warning-emphasis">{{ p.wish_hint }}</span>{% endif %}</div>{% endif %}{% else %}<span class="text-muted">－</span>{% endif %}
      </td>
      <td class="small text-nowrap">
        {% if p.stance_latest %}
        <span class="badge {{ 'text-bg-secondary' if p.stance_old else ('text-bg-success' if p.stance_latest in C.MONTHLY_KPI['self_positive'] else ('text-bg-primary' if p.stance_latest in C.MONTHLY_KPI['ask_positive'] else 'text-bg-light border')) }}">{{ p.stance_latest }}</span>
        <div class="text-muted">{{ p.stance_label }}{% if p.stance_old %}（{{ C.STANCE_MONTHS }}か月より前）{% endif %}</div>
        {% elif p.monthly %}<span class="text-muted">進め方の回答なし</span>
        {% else %}<span class="text-muted">－</span>{% endif %}
        {% if p.monthly %}<div class="text-muted">自分で／一緒に {{ p.self_positive_count }}回（回答 {{ p.monthly|length }} 回中）</div>{% endif %}
      </td>
      <td class="small dx-wide">{{ p.skills|join('・') or (C.SKILL_NONE if p.skills_none else '－') }}</td>
      <td class="small dx-wide">{{ p.learn|join('・') or (C.SKILL_NONE if p.learn_none else '－') }}</td>
      <td class="small dx-tasks">
        {% if p.task_count %}<div>{{ p.task_count }} 件</div>{% endif %}
        {% for x in p.target_tasks %}<div><a href="{{ url_for('cases.detail', task_id=x.task.id) }}">{{ x.task.name }}</a> <span class="badge {{ 'text-bg-primary' if x.label == C.BUILD_STAGE_SELF else 'text-bg-light border' }}">{{ x.label }}</span></div>{% endfor %}
        {% if not p.task_count %}
          {% if p.monthly_wants %}
          <div class="text-muted">業務の登録なし。{{ p.monthly_wants_label }}の月次で挙げたなくしたい業務:</div>
          <div>{{ p.monthly_wants|join('・') }}</div>
          <a href="{{ url_for('survey.start', emp=p.emp_id, src='monthly', m=p.monthly_wants_id) }}">「自動化を検討」で流れを整理する</a>
          {% else %}<span class="text-muted">－</span>{% endif %}
        {% endif %}
      </td>
    </tr>
    {% else %}
    <tr><td colspan="8" class="text-muted">該当する人はいません</td></tr>
    {% endfor %}
    </tbody>
  </table>
  </div>
</div></div>

<div class="row g-3">
  <div class="col-lg-4">
    <div class="card h-100"><div class="card-body">
      <h5 class="card-title">意欲（検討と直近の月次）</h5>
      {% set maxc = (wish_counts|map(attribute=1)|max) if wish_counts else 1 %}
      {% for opt, n in wish_counts %}
      <div class="d-flex align-items-center small mb-1">
        <div style="width: 15rem">{{ opt }}</div>
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
{% block title %}自動化を検討{% endblock %}
{% block content %}
<div class="rq" id="rq">
  <div class="rq-top">
    <h1 class="rq-title">自動化を検討</h1>
    <p class="rq-lead">AI が短い質問をしながら、あなたの業務の流れを図にしていきます。所要時間は2〜3分。正確でなくて大丈夫、あとで改善チームが確認します。</p>
  </div>

  <section class="card rq-card rq-sec" id="sec-person">
    <div class="card-body">
      <div class="rq-sec-head"><span class="rq-no">1</span><span class="rq-sec-title">あなたのこと</span></div>
      <div id="rq-person"></div>
    </div>
  </section>

  <section class="card rq-card rq-sec" id="sec-task">
    <div class="card-body">
      <div class="rq-sec-head"><span class="rq-no">2</span><span class="rq-sec-title">検討する業務</span><span class="rq-sec-sub" id="rq-registered-count"></span></div>
      <div id="rq-registered" class="rq-registered"></div>

      <div id="rq-from-monthly" class="rq-from d-none"></div>
      <div id="rq-pick">
        <div class="rq-label"><span>どの業務を検討しますか</span></div>
        <div class="rq-suggest d-none" id="rq-suggest"></div>
        <div class="rq-pick-row">
          <input type="text" class="form-control" id="rq-title" placeholder="業務名（分かる範囲で。例：日報の実績を集計表に転記する）" maxlength="200">
          <button type="button" class="btn btn-primary" id="rq-start"><i class="bi bi-chat-dots"></i> 聞き取りを始める</button>
        </div>
        <div class="small text-muted mt-2">業務名が決まっていなくても始められます。AI が聞きながら決めます。</div>
        <div id="rq-dup-note" class="rq-dup d-none"></div>
      </div>

      <div id="rq-intake" class="rq-intake d-none">
        <div class="rq-flowpane" id="rq-flowpane">
          <div class="rq-flowhead">
            <span class="rq-flowtitle" id="rq-flow-title">流れ図</span>
            <span class="rq-flowhead-right"><button type="button" class="btn btn-link btn-sm rq-replay d-none" id="rq-replay"><i class="bi bi-play-fill"></i> 最初から見る</button><span class="rq-stepcount" id="rq-stepcount">手順 0</span></span>
          </div>
          <div class="rq-flowbox" id="rq-flow"><div class="rq-flow-empty">ここに流れができていきます</div></div>
          <ol class="rq-steps" id="rq-steps"></ol>
        </div>
        <div class="rq-chatpane">
          <div class="rq-chathead"><span id="rq-progress">AI が質問します</span><button type="button" class="btn btn-link btn-sm" id="rq-restart">最初からやり直す</button></div>
          <div class="rq-messages" id="rq-messages"></div>
          <div class="rq-chips" id="rq-chips"></div>
          <div class="rq-inputrow" id="rq-inputrow">
            <input type="text" class="form-control" id="rq-input" placeholder="答えを入力" maxlength="1000" autocomplete="off">
            <button type="button" class="btn btn-primary" id="rq-send" aria-label="送る"><i class="bi bi-send"></i></button>
          </div>
          <div class="rq-tools">
            <button type="button" class="btn btn-outline-secondary btn-sm" id="rq-bulk-toggle"><i class="bi bi-pencil"></i> 一気に書く</button>
            <label class="btn btn-outline-secondary btn-sm mb-0"><i class="bi bi-camera"></i> 写真を貼る<input type="file" id="rq-photo" accept="image/*" class="d-none"></label>
          </div>
          <div class="rq-bulk d-none" id="rq-bulk">
            <textarea class="form-control" id="rq-bulk-text" rows="5" placeholder="やっている順に、1行に1手順で。&#10;例：&#10;生産管理システムから前日の実績をCSVで出す&#10;Excelの集計表に貼り付ける&#10;不良件数を手で数えて追記する&#10;課長にメールで送る"></textarea>
            <button type="button" class="btn btn-primary btn-sm mt-2" id="rq-bulk-send">読み取ってもらう</button>
          </div>
          <div class="rq-confirm d-none" id="rq-confirm">
            <div class="rq-confirm-q">この流れで合っていますか？</div>
            <div class="d-flex gap-2 flex-wrap">
              <button type="button" class="btn btn-success" id="rq-confirm-yes"><i class="bi bi-check-lg"></i> はい、合っています</button>
              <button type="button" class="btn btn-outline-secondary" id="rq-confirm-no">直したい（会話で伝える）</button>
            </div>
          </div>
        </div>
      </div>

      <div id="rq-final" class="rq-final d-none">
        <div class="rq-final-title">あと少しです。3つ選んで登録してください。</div>
        <div class="rq-field" data-key="frequency"><div class="rq-label"><span>頻度</span><span class="rq-req">必須</span></div><div class="rq-pills" id="rq-frequency"></div><div class="rq-err"></div></div>
        <div class="rq-field" data-key="duration"><div class="rq-label"><span>1回の時間</span><span class="rq-req">必須</span></div><div class="rq-pills" id="rq-duration"></div><div class="rq-err"></div></div>
        <div class="rq-field d-none" data-key="shift" id="rq-shift-field"><div class="rq-label"><span>この業務はいつやりますか</span><span class="rq-req">必須</span></div><div class="rq-pills" id="rq-shift"></div><div class="rq-err"></div></div>
        <div class="rq-field" data-key="wish"><div class="rq-label"><span>この業務をどうしたいですか</span><span class="rq-req">必須</span></div><div class="rq-pills" id="rq-wish"></div><div class="rq-err"></div></div>
        <div class="form-check rq-self">
          <input class="form-check-input" type="checkbox" id="rq-self">
          <label class="form-check-label" for="rq-self" id="rq-self-label"></label>
          <div class="small text-muted d-none" id="rq-self-note"></div>
        </div>
        <div id="rq-skills" class="rq-skills d-none">
          <div class="rq-field"><div class="rq-label"><span>使える技術</span><span class="rq-opt">任意</span></div><div class="rq-pills" id="rq-skills-have"></div></div>
          <div class="rq-field"><div class="rq-label"><span>学びたい技術</span><span class="rq-opt">任意</span></div><div class="rq-pills" id="rq-skills-learn"></div></div>
        </div>
        <div class="rq-actions">
          <button type="button" class="btn btn-primary btn-lg" id="rq-save-task"><i class="bi bi-plus-lg"></i> この業務を登録</button>
        </div>
        <div id="rq-dup" class="rq-dup d-none"></div>
      </div>

      <div id="rq-result" class="rq-result d-none">
        <div class="rq-check"><i class="bi bi-check-lg"></i></div>
        <div class="fw-semibold" id="rq-result-title"></div>
        <div class="rq-impression" id="rq-impression"></div>
        <div class="small text-muted mt-1">改善チームが内容を確認して連絡します。活動状況にも載ります。</div>
        <div class="d-flex flex-wrap justify-content-center gap-2 mt-3">
          <button type="button" class="btn btn-primary d-none" id="rq-next-monthly"></button>
          <button type="button" class="btn btn-outline-primary" id="rq-another"><i class="bi bi-plus-lg"></i> 続けて別の業務を検討する</button>
        </div>
      </div>
    </div>
  </section>

  <div class="rq-finish">
    <button type="button" class="btn btn-primary btn-lg" id="rq-finish"><i class="bi bi-send"></i> 送信して完了</button>
  </div>

  <section class="card rq-card rq-done d-none" id="rq-done">
    <div class="card-body text-center">
      <div class="rq-check"><i class="bi bi-check-lg"></i></div>
      <h2 class="h4 mb-1">受け付けました</h2>
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
    prefill: {{ prefill|tojson }},
    monthly: {{ monthly|tojson }},
    api: {
      lookup: "{{ url_for('monthly.api_employee', emp_id='') }}",
      respondent: "{{ url_for('survey.api_respondent') }}",
      reset: "{{ url_for('survey.api_reset') }}",
      tasks: "{{ url_for('survey.api_task_create') }}",
      intake: "{{ url_for('survey.api_intake') }}"
    }
  };
</script>
<script src="https://cdn.jsdelivr.net/npm/mermaid@11.4.1/dist/mermaid.min.js"></script>
<script src="{{ url_for('static', filename='survey.js') }}"></script>
{% endblock %}
''',
}


# ===== 静的ファイル =====
STATIC = {
    'app.js': r'''// 共通スクリプト
document.addEventListener("DOMContentLoaded", function () {
  document.querySelectorAll("form[data-confirm]").forEach(function (form) {
    form.addEventListener("submit", function (e) {
      if (!window.confirm(form.getAttribute("data-confirm"))) {
        e.preventDefault();
      }
    });
  });
});

(function () {
  if (window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches) return;
  var targets = document.querySelectorAll(".stat-value, .mini-value, .kpi-value");
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
    setTimeout(function () { if (node.nodeValue !== text) node.nodeValue = text; }, delay + duration + 300);
  });
})();

(function () {
  var sel = document.querySelector(".pb-scope-select");
  if (!sel) return;
  if (window.CSS && CSS.supports && CSS.supports("field-sizing", "content")) return;
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

// 仕上げの動き: ヘッダーの影・波紋・スクロールで現れる・お知らせ・送信中・紙吹雪
(function () {
  var reduce = !!(window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches);
  var root = document.documentElement;
  var header = document.querySelector(".app-header");

  function onScroll() { root.classList.toggle("is-scrolled", window.scrollY > 4); }
  function headerHeight() { if (header) root.style.setProperty("--ui-header-h", header.offsetHeight + "px"); }
  window.addEventListener("scroll", onScroll, { passive: true });
  window.addEventListener("resize", headerHeight);
  onScroll();
  headerHeight();

  document.addEventListener("pointerdown", function (e) {
    if (reduce || e.button !== 0) return;
    var btn = e.target.closest ? e.target.closest(".btn") : null;
    if (!btn || btn.disabled || btn.tagName === "INPUT" || btn.classList.contains("btn-link") || btn.classList.contains("is-busy")) return;
    var r = btn.getBoundingClientRect();
    var size = Math.max(r.width, r.height) * 2.2;
    var wrap = btn.querySelector(":scope > .ui-ripple");
    if (!wrap) {
      wrap = document.createElement("span");
      wrap.className = "ui-ripple";
      wrap.setAttribute("aria-hidden", "true");
      btn.appendChild(wrap);
    }
    var dot = document.createElement("span");
    dot.style.width = dot.style.height = size + "px";
    dot.style.left = (e.clientX - r.left - size / 2) + "px";
    dot.style.top = (e.clientY - r.top - size / 2) + "px";
    wrap.appendChild(dot);
    dot.addEventListener("animationend", function () { dot.remove(); });
  });

  if (!reduce && "IntersectionObserver" in window) {
    var io = new IntersectionObserver(function (entries) {
      var n = 0;
      entries.forEach(function (en) {
        if (!en.isIntersecting) return;
        var el = en.target;
        io.unobserve(el);
        el.style.transitionDelay = Math.min(n++, 5) * 70 + "ms";
        el.classList.add("is-in");
        setTimeout(function () { el.classList.add("is-done"); el.style.transitionDelay = ""; }, 1100);
      });
    }, { rootMargin: "0px 0px -6% 0px", threshold: 0.05 });
    var limit = window.innerHeight * 0.96;
    document.querySelectorAll("main .card, main .section-title, main .pb-row, main .pb-section").forEach(function (el) {
      if (el.closest("#fe-app, .ui-toasts, .ui-reveal")) return;
      if (el.getBoundingClientRect().top < limit) return;
      el.classList.add("ui-reveal");
      io.observe(el);
    });
  }

  document.querySelectorAll(".ui-toasts .alert").forEach(function (a) {
    var close = function () {
      if (a.classList.contains("is-leaving")) return;
      a.classList.add("is-leaving");
      setTimeout(function () { a.remove(); }, 340);
    };
    a.addEventListener("animationend", function (e) { if (e.animationName === "ui-toast-timer") close(); });
    if (reduce) setTimeout(close, 6000);
  });

  document.addEventListener("submit", function (e) {
    if (e.defaultPrevented) return;
    var f = e.target;
    if (!f) return;
    var b = e.submitter || f.querySelector('button[type="submit"], button:not([type])');
    if (!b || !b.classList.contains("btn")) return;
    b.classList.add("is-busy");
    b.setAttribute("aria-busy", "true");
    setTimeout(function () { b.classList.remove("is-busy"); b.removeAttribute("aria-busy"); }, 10000);
  });
  window.addEventListener("pageshow", function () {
    document.querySelectorAll(".btn.is-busy").forEach(function (b) { b.classList.remove("is-busy"); b.removeAttribute("aria-busy"); });
  });

  // 紙吹雪。window.uiCelebrate() で呼ぶ。data-celebrate のある画面は開いたときに1回
  window.uiCelebrate = function () {
    if (reduce || !document.body) return;
    var c = document.createElement("canvas");
    c.className = "ui-confetti";
    c.setAttribute("aria-hidden", "true");
    document.body.appendChild(c);
    var ctx = c.getContext("2d");
    var dpr = window.devicePixelRatio || 1, W = window.innerWidth, H = window.innerHeight;
    c.width = W * dpr; c.height = H * dpr;
    ctx.scale(dpr, dpr);
    var colors = ["#0a84ff", "#30d158", "#ff9f0a", "#bf5af2", "#ff375f", "#5ac8fa", "#ffd60a", "#5e5ce6"];
    var parts = [];
    for (var i = 0; i < 150; i++) {
      var left = i % 2 === 0;
      parts.push({
        x: left ? -12 : W + 12, y: H * (0.5 + Math.random() * 0.25),
        vx: (left ? 1 : -1) * (5 + Math.random() * 8), vy: -(10 + Math.random() * 9),
        w: 6 + Math.random() * 6, h: 9 + Math.random() * 9, r: Math.random() * Math.PI, vr: (Math.random() - 0.5) * 0.32,
        color: colors[i % colors.length]
      });
    }
    var start = null, life = 2800;
    function frame(t) {
      if (start === null) start = t;
      var el = t - start;
      ctx.clearRect(0, 0, W, H);
      parts.forEach(function (p) {
        p.vy += 0.3; p.vx *= 0.99; p.x += p.vx; p.y += p.vy; p.r += p.vr;
        ctx.save();
        ctx.translate(p.x, p.y);
        ctx.rotate(p.r);
        ctx.globalAlpha = Math.max(0, 1 - Math.max(0, el - 1600) / (life - 1600));
        ctx.fillStyle = p.color;
        ctx.fillRect(-p.w / 2, -p.h / 2, p.w, p.h * Math.max(0.25, Math.abs(Math.cos(p.r * 1.7))));
        ctx.restore();
      });
      if (el < life) requestAnimationFrame(frame); else c.remove();
    }
    requestAnimationFrame(frame);
  };
  if (document.querySelector("[data-celebrate]")) setTimeout(window.uiCelebrate, 300);
})();
''',
    'flowedit.js': r'''// 業務マップ 編集エディタ
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
    typeRow: document.getElementById("fe-type-row"),
    whoRow: document.getElementById("fe-who-row"),
    whoInput: document.getElementById("fe-who-input"),
    actorSelect: document.getElementById("fe-actor-select"),
    mediumRow: document.getElementById("fe-medium-row"),
    mediumSelect: document.getElementById("fe-medium-select"),
    edgeType: document.getElementById("fe-edge-type"),
    methodRow: document.getElementById("fe-method-row"),
    methodSelect: document.getElementById("fe-method-select"),
    handoffRow: document.getElementById("fe-handoff-row"),
    handoffSelect: document.getElementById("fe-handoff-select"),
    undo: document.getElementById("fe-undo"),
    redo: document.getElementById("fe-redo")
  };

  var G = { nodes: [], edges: [], suppressed_nodes: [], suppressed_edges: [] };

  var KC = JSON.parse(document.getElementById("fe-kinds").textContent);
  var STATUS_MARKS = KC.status_marks || {};
  function vrow(group, key) { var rows = KC.vocab[group] || []; for (var i = 0; i < rows.length; i++) if (rows[i].key === key) return rows[i]; return null; }
  function typeOf(n) { return n.type === "info" || n.type === "work" ? n.type : (KC.cls_type[n.cls] || "work"); }
  function shapeOf(n) {
    if (typeOf(n) !== "info") return "rect";
    var r = vrow("media", n.medium || "unknown");
    return r && ["doc", "cyl", "stadium"].indexOf(r.shape) !== -1 ? r.shape : "doc";
  }
  function iconOf(n) {
    if (typeOf(n) === "info") {
      var m = n.medium || "unknown", r = vrow("media", m) || vrow("media", "unknown");
      return [r ? r.icon : "bi-question-square", "m-" + m];
    }
    if (n.cls === "done") return ["bi-check-circle-fill", "done"];
    if (n.cls === "task" || STATUS_MARKS[n.cls]) return ["", ""];
    if (n.cls === "pending") return ["bi-hourglass-split", "pending"];
    if (n.who_state === "unconfirmed" || !n.actor) return ["bi-question-circle", "who-unconfirmed"];
    var a = vrow("actors", n.actor);
    return [a ? a.icon : "bi-person-fill", "a-" + n.actor];
  }
  function normText(v) { return String(v || "").replace(/[\s　（）()・,、。.／/\-−―「」『』]/g, "").toLowerCase(); }
  function showWho(n) {
    return typeOf(n) === "work" && (n.cls === "work" || n.cls === "warn") && (n.who_state === "ok" || n.who_state === "unknown") && n.who && normText(n.who) !== normText(n.text);
  }
  function edgeLabel(e) {
    if (e.label) return e.label;
    if (e.type === "work") return (e.handoff === "push" || e.handoff === "pull") && KC.vocab.handoff[e.handoff] ? KC.vocab.handoff[e.handoff].label : "";
    if (e.show_method && e.method) { var r = vrow("methods", e.method); return r ? r.label : ""; }
    return "";
  }
  var SHAPE_PATHS = { doc: "M0 0H100V84C82 70 66 70 50 84S18 98 0 84Z" };
  function shapeBg(shape) {
    var svg = svgEl("svg", { "class": "fe-shape-bg", viewBox: "0 0 100 100", preserveAspectRatio: "none" });
    svg.appendChild(svgEl("path", { d: SHAPE_PATHS[shape] }));
    return svg;
  }
  function glyph(shape) { return '<svg class="ks" viewBox="0 0 18 14" aria-hidden="true">' + (KC.glyphs[shape] || KC.glyphs.rect) + "</svg>"; }
  function escHtml(v) { return String(v).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }
  function renderLegend() {
    var box = document.getElementById("fe-kind-legend");
    if (!box) return;
    var html = '<div class="legend-row lg-grammar"><span class="legend-title">箱と線</span><span class="lg-item">' + glyph("rect") + '作業</span><span class="lg-item">' + glyph("doc") + "情報</span>" +
      '<span class="lg-item"><i class="lg-line lg-line-work"></i>仕事の流れ</span><span class="lg-item"><i class="lg-line lg-line-info"></i>情報の流れ</span>' +
      '<span class="lg-item"><i class="lg-line lg-line-guess"></i>名前から推定</span></div>';
    var roles = Object.keys(KC.cls_labels).filter(function (c) { return G.nodes.some(function (n) { return n.cls === c; }); });
    if (roles.length) html += '<div class="legend-row"><span class="legend-title">色</span>' + roles.map(function (c) { return '<span class="legend-chip lg-' + c + '">' + escHtml(KC.cls_labels[c]) + "</span>"; }).join("") + "</div>";
    var media = (KC.vocab.media || []).filter(function (r) { return G.nodes.some(function (n) { return typeOf(n) === "info" && (n.medium || "unknown") === r.key; }); });
    if (media.length) html += '<div class="kind-legend"><span class="legend-title">媒体</span>' + media.map(function (r) {
      return '<span class="kind-chip">' + glyph(r.shape) + '<i class="bi ' + r.icon + " fi m-" + r.key + '"></i>' + escHtml(r.label) + "</span>";
    }).join("") + "</div>";
    var actors = (KC.vocab.actors || []).filter(function (r) { return G.nodes.some(function (n) { return typeOf(n) === "work" && (n.cls === "work" || n.cls === "warn") && n.actor === r.key && n.who_state !== "unconfirmed"; }); });
    var unconf = G.nodes.some(function (n) { return typeOf(n) === "work" && n.cls === "work" && (n.who_state === "unconfirmed" || !n.actor); });
    if (actors.length || unconf) html += '<div class="kind-legend"><span class="legend-title">だれ</span>' + actors.map(function (r) {
      return '<span class="kind-chip"><i class="bi ' + r.icon + " fi a-" + r.key + '"></i>' + escHtml(r.label) + "</span>";
    }).join("") + (unconf ? '<span class="kind-chip"><i class="bi bi-question-circle fi who-unconfirmed"></i>だれ未確認</span>' : "") + "</div>";
    box.innerHTML = html;
  }
  var classifySeq = {};
  async function classify(list) {
    list = list.filter(Boolean);
    if (!list.length) return;
    // 送った時点の中身を覚えておき、返ってくるまでに変わった箱・後から送り直した箱には当てない
    var sent = {};
    list.forEach(function (n) { classifySeq[n.id] = (classifySeq[n.id] || 0) + 1; sent[n.id] = { seq: classifySeq[n.id], type: typeOf(n), text: n.text, who: n.who || "" }; });
    try {
      var r = await fetch(KC.classify_url, { method: "POST", headers: { "Content-Type": "application/json", Accept: "application/json" },
        body: JSON.stringify({ items: list.map(function (n) { return { id: n.id, type: typeOf(n), text: n.text, who: n.who || "", workplace: n.group_title || "" }; }) }) });
      var j = await r.json();
      var changed = false;
      (j.items || []).forEach(function (it) {
        var n = nodeById(it.id), was = sent[it.id];
        if (!n || !was || classifySeq[n.id] !== was.seq || typeOf(n) !== was.type || n.text !== was.text || (n.who || "") !== was.who) return;
        if (typeOf(n) === "info") { if (!n.medium_edited && n.medium !== it.medium) { n.medium = it.medium; changed = true; } }
        else if (!n.actor_locked) { n.actor = it.actor; n.who_state = it.who_state; if (it.who) n.who = it.who; changed = true; }
        else if (it.who && it.who !== n.who) { n.who = it.who; changed = true; }
      });
      if (!changed) return;
      // 判定の結果は直前の操作の一部として履歴に入れる（元に戻す・やり直しで消えないように）
      if (historyIndex >= 0) history[historyIndex] = snapshot();
      updateButtons();
      updateDirty();
      if (editingId) { drawEdges(); renderLegend(); } else renderAll();
      if (popNodeId) refreshNodePopover();
    } catch (e) { setStatus("種類の判定に失敗しました（通信状態を確認してください）"); }
  }
  var meta = {};
  var view = { x: 40, y: 40, zoom: 1 };
  var history = [], historyIndex = -1, MAX_HISTORY = 60;
  var savedSnapshot = null, dirty = false;
  var selected = [];
  var drag = null, pan = null, selecting = null, linkSource = null, editingId = null;
  var popNodeId = null, popEdgeId = null;

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

  function renderAll() {
    els.nodes.innerHTML = "";
    G.nodes.forEach(function (n) { els.nodes.appendChild(createNodeEl(n)); });
    drawEdges();
    renderLegend();
  }

  function createNodeEl(n) {
    var div = document.createElement("div");
    var type = typeOf(n), shape = shapeOf(n), ic = iconOf(n);
    div.className = "fe-node fe-type-" + type + " fe-cls-" + (n.cls || (type === "info" ? "source" : "work")) + " fe-shape-" + shape + (n.source === "manual" ? " fe-manual" : "");
    div.dataset.id = n.id;
    if (SHAPE_PATHS[shape]) div.appendChild(shapeBg(shape));
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
    btns.appendChild(btn("🎨", "種類・だれ・媒体・色", "fe-btn-color", function (e) { openNodePopover(n.id, e); }));
    btns.appendChild(btn("✕", n.source === "manual" ? "削除" : "非表示にする", "fe-btn-del", function () { deleteNode(n.id); }));
    div.appendChild(btns);

    if (showWho(n)) {
      var whoEl = document.createElement("div");
      whoEl.className = "fe-who";
      whoEl.textContent = n.who;
      div.appendChild(whoEl);
    }
    var label = document.createElement("div");
    label.className = "fe-label";
    if (ic[0]) {
      var icon = document.createElement("i");
      icon.className = "bi " + ic[0] + " fe-icon fi " + ic[1];
      label.appendChild(icon);
    }
    var text = document.createElement("div");
    text.className = "fe-text";
    text.textContent = n.text || "";
    label.appendChild(text);
    if (STATUS_MARKS[n.cls]) {
      var state = document.createElement("span");
      state.className = "fe-state";
      state.textContent = STATUS_MARKS[n.cls];
      label.appendChild(state);
    }
    div.appendChild(label);
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
      var type = e.type === "work" ? "work" : "info";
      var g = svgEl("g", { "class": "fe-edge-group", "data-id": e.id });
      var hit = svgEl("path", { d: c.d, "class": "fe-hit" });
      var path = svgEl("path", { d: c.d, "class": "fe-edge fe-type-" + type + (e.inferred ? " fe-inferred" : "") + (e.source === "manual" ? " fe-src-manual" : "") });
      path.setAttribute("marker-end", "url(#fe-arrow-" + type + ")");
      g.appendChild(hit);
      g.appendChild(path);
      var lab = edgeLabel(e);
      if (lab) {
        var t = svgEl("text", { x: c.mx, y: c.my - 4, "class": "fe-edge-label", "text-anchor": "middle" });
        t.textContent = lab;
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

  function paintSelection() {
    els.nodes.querySelectorAll(".fe-node").forEach(function (el) { el.classList.toggle("fe-selected", selected.indexOf(el.dataset.id) !== -1); });
  }
  function selectOnly(id) {
    selected = id ? [id] : [];
    paintSelection();
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
    paintSelection();
  }

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
        if (typeOf(n) === "info" && !n.medium_edited) classify([n]);
      } else {
        renderAll();
      }
    }
    text.addEventListener("blur", function () { finish(false); }, { once: true });
    text.addEventListener("keydown", function (e) {
      if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); text.blur(); }
      if (e.key === "Escape") { e.preventDefault(); finish(true); }
    });
  }

  function addNode(x, y, type) {
    type = type === "info" ? "info" : "work";
    var n = { id: newId("m_"), text: type === "info" ? "新しい情報" : "新しい作業", type: type, cls: type === "info" ? "source" : "work",
      medium: type === "info" ? "unknown" : "", actor: "", who: "", who_state: type === "work" ? "unconfirmed" : "", group: null, group_title: null, url: null,
      source: "manual", x: x, y: y, color: null, text_edited: false };
    G.nodes.push(n);
    els.nodes.appendChild(createNodeEl(n));
    selectOnly(n.id);
    drawEdges();
    pushHistory();
    setTimeout(function () { startEdit(n.id); }, 50);
    return n;
  }

  function pushManualEdge(from, to, type) {
    G.edges.push({ id: from + "->" + to, from: from, to: to, label: "", type: type, method: "", handoff: "", show_method: false, inferred: false, source: "manual", reversed: false, label_edited: false });
  }

  function addChild(parentId) {
    var p = nodeById(parentId), b = bounds(parentId);
    if (!p || !b) return;
    var childType = typeOf(p) === "info" ? "work" : "info";
    var n = addNode(p.x + b.w + 90, p.y, childType);
    pushManualEdge(parentId, n.id, "info");
    drawEdges();
    pushHistory();
  }

  // 箱と、その箱につながる線を図から外す（自動生成の箱は非表示の一覧に入れる）
  function dropNode(id) {
    var n = nodeById(id);
    if (!n) return false;
    if (n.source !== "manual" && G.suppressed_nodes.indexOf(id) === -1) G.suppressed_nodes.push(id);
    G.nodes = G.nodes.filter(function (m) { return m.id !== id; });
    G.edges = G.edges.filter(function (e) { return e.from !== id && e.to !== id; });
    return true;
  }

  function deleteNode(id) {
    if (!dropNode(id)) return;
    selected = selected.filter(function (s) { return s !== id; });
    closePopovers();
    renderAll();
    pushHistory();
  }

  function deleteSelected() {
    if (!selected.length) return;
    selected.forEach(function (id) { dropNode(id); });
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
    var na = nodeById(from), nb = nodeById(to);
    if (na && nb && typeOf(na) === "info" && typeOf(nb) === "info") {
      setMode("情報どうしは結べません。間に作業を置いてください");
      setTimeout(cancelLink, 1800);
      return;
    }
    var id = from + "->" + to;
    if (edgeById(id)) { setMode("この矢印はすでにあります"); setTimeout(cancelLink, 1200); return; }
    var sup = G.suppressed_edges.indexOf(id);
    if (sup !== -1) G.suppressed_edges.splice(sup, 1);
    var ltype = na && nb && typeOf(na) === "work" && typeOf(nb) === "work" ? "work" : "info";
    pushManualEdge(from, to, ltype);
    cancelLink();
    drawEdges();
    pushHistory();
    setMode("✅ 矢印を作成しました");
    setTimeout(function () { if (!linkSource) setMode(""); }, 1500);
  }

  function placePopover(pop, ev) {
    var r = els.container.getBoundingClientRect();
    pop.classList.add("fe-show");
    var x = Math.min(ev.clientX - r.left + 8, r.width - pop.offsetWidth - 4), y = Math.min(ev.clientY - r.top + 8, r.height - pop.offsetHeight - 4);
    pop.style.left = Math.max(4, x) + "px";
    pop.style.top = Math.max(4, y) + "px";
  }
  function closePopovers() {
    // 入力中の だれ・線のメモ は閉じる前に確定する（change を先に起こす）
    var a = document.activeElement;
    if (a && (a === els.whoInput || a === els.edgeLabel)) a.blur();
    els.nodePop.classList.remove("fe-show");
    els.edgePop.classList.remove("fe-show");
    wbPop.classList.remove("fe-show");
    popNodeId = null; popEdgeId = null;
  }
  function fillSelect(sel, first, rows) {
    sel.innerHTML = "";
    first.concat(rows).forEach(function (o) { var op = document.createElement("option"); op.value = o[0]; op.textContent = o[1]; sel.appendChild(op); });
  }
  fillSelect(els.actorSelect, [["", "自動（名前から）"]], (KC.vocab.actors || []).map(function (r) { return [r.key, r.label]; }));
  fillSelect(els.mediumSelect, [["", "自動（名前から）"]], (KC.vocab.media || []).map(function (r) { return [r.key, r.label]; }));
  fillSelect(els.methodSelect, [["", "自動"], ["-", "なし（書かない）"]], (KC.vocab.methods || []).map(function (r) { return [r.key, r.label]; }));
  fillSelect(els.handoffSelect, [["", "自動"], ["-", "未設定"]], [["push", KC.vocab.handoff.push.label], ["pull", KC.vocab.handoff.pull.label], ["none", "引き継ぎなし（同じ人）"]]);
  function refreshNodePopover() {
    var n = nodeById(popNodeId);
    if (!n) return;
    var type = typeOf(n);
    els.typeRow.style.display = n.source === "manual" ? "flex" : "none";
    els.typeRow.querySelectorAll("[data-type]").forEach(function (b) { b.classList.toggle("active", b.dataset.type === type); });
    els.whoRow.style.display = type === "work" && ["task", "done", "hold", "abolished", "pending"].indexOf(n.cls) === -1 ? "block" : "none";
    if (document.activeElement !== els.whoInput) els.whoInput.value = n.who || "";
    els.actorSelect.value = n.actor_locked ? (n.actor || "") : "";
    els.mediumRow.style.display = type === "info" ? "block" : "none";
    els.mediumSelect.value = n.medium_edited ? (n.medium || "") : "";
  }
  function openNodePopover(id, ev) {
    closePopovers();
    if (!nodeById(id)) return;
    popNodeId = id;
    refreshNodePopover();
    placePopover(els.nodePop, ev);
  }
  function openEdgePopover(id, ev) {
    closePopovers();
    var e = edgeById(id);
    if (!e) return;
    popEdgeId = id;
    var work = e.type === "work";
    els.edgeLabel.value = e.label || "";
    els.edgeType.textContent = work ? "仕事の流れ（作業 → 作業）" : "情報の流れ";
    els.methodRow.style.display = work ? "none" : "block";
    els.handoffRow.style.display = work ? "block" : "none";
    els.methodSelect.value = e.method_edited ? (e.method || "-") : "";
    els.handoffSelect.value = e.handoff_edited ? (e.handoff || "-") : "";
    placePopover(els.edgePop, ev);
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
  els.typeRow.querySelectorAll("[data-type]").forEach(function (b) {
    b.addEventListener("click", function () {
      var n = nodeById(popNodeId);
      var t = b.dataset.type;
      if (!n || n.source !== "manual" || typeOf(n) === t) return;
      n.type = t;
      n.cls = t === "info" ? "source" : "work";
      n.medium = t === "info" ? "unknown" : "";
      n.medium_edited = false;
      n.actor = ""; n.who = ""; n.who_state = t === "work" ? "unconfirmed" : ""; n.who_edited = false; n.actor_locked = false;
      // 線を直すのは、この箱につながる線だけ
      var mine = function (e) { return e.from === n.id || e.to === n.id; };
      G.edges = G.edges.filter(function (e) {
        if (!mine(e)) return true;
        var x = nodeById(e.from), y = nodeById(e.to);
        var bad = x && y && typeOf(x) === "info" && typeOf(y) === "info";
        if (bad && e.source !== "manual" && G.suppressed_edges.indexOf(e.id) < 0) G.suppressed_edges.push(e.id);
        return !bad;
      });
      G.edges.forEach(function (e) { if (!mine(e)) return; var x = nodeById(e.from), y = nodeById(e.to); if (x && y) e.type = typeOf(x) === "work" && typeOf(y) === "work" ? "work" : "info"; });
      renderAll();
      refreshNodePopover();
      pushHistory();
      classify([n]);
    });
  });
  els.whoInput.addEventListener("change", function () {
    var n = nodeById(popNodeId);
    if (!n) return;
    n.who = els.whoInput.value.trim();
    n.who_edited = true;
    pushHistory();
    // だれの区分を手で選んでいても、名前はサーバーで氏名を置き換えてから使う
    if (n.actor_locked) renderAll();
    classify([n]);
  });
  els.actorSelect.addEventListener("change", function () {
    var n = nodeById(popNodeId);
    if (!n) return;
    var v = els.actorSelect.value;
    n.who_edited = true;
    if (v) { n.actor = v; n.actor_locked = true; n.who_state = v === "unknown" ? "unknown" : "ok"; renderAll(); }
    else { n.actor_locked = false; classify([n]); }
    pushHistory();
  });
  els.mediumSelect.addEventListener("change", function () {
    var n = nodeById(popNodeId);
    if (!n) return;
    var v = els.mediumSelect.value;
    if (v) { n.medium = v; n.medium_edited = true; renderAll(); }
    else { n.medium_edited = false; classify([n]); }
    pushHistory();
  });
  els.methodSelect.addEventListener("change", function () {
    var e = edgeById(popEdgeId);
    if (!e) return;
    var v = els.methodSelect.value;
    if (!v) {
      e.method_edited = false;
      if (e.source === "manual") { e.method = ""; e.show_method = false; }
      else { e.method = e.auto_method || ""; e.show_method = !!e.auto_show_method; }
    }
    else if (v === "-") { e.method = ""; e.show_method = false; e.method_edited = true; }
    else { e.method = v; e.show_method = true; e.method_edited = true; }
    drawEdges();
    pushHistory();
  });
  els.handoffSelect.addEventListener("change", function () {
    var e = edgeById(popEdgeId);
    if (!e) return;
    var v = els.handoffSelect.value;
    if (!v) { e.handoff_edited = false; e.handoff = e.source === "manual" ? "" : (e.auto_handoff || ""); }
    else { e.handoff = v === "-" ? "" : v; e.handoff_edited = true; }
    drawEdges();
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

  els.container.addEventListener("mousedown", function (e) {
    if (e.button !== 0) return;
    var empty = e.target === els.container || e.target === els.canvas || e.target === els.nodes || e.target.id === "fe-svg";
    if (!empty) return;
    closePopovers();
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
    var typing = /^(INPUT|TEXTAREA|SELECT)$/.test(e.target.tagName) || e.target.isContentEditable || !!(e.target.closest && e.target.closest(".fe-popover"));
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "s") { e.preventDefault(); save(); return; }
    if (typing) return;
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "z" && !e.shiftKey) { e.preventDefault(); undo(); }
    else if ((e.ctrlKey || e.metaKey) && (e.key.toLowerCase() === "y" || (e.key.toLowerCase() === "z" && e.shiftKey))) { e.preventDefault(); redo(); }
    else if (e.key === "Delete" || e.key === "Backspace") { if (selected.length) { e.preventDefault(); deleteSelected(); } }
    else if (e.key === "Escape") {
      if (!linkSource && !app.querySelector(".fe-popover.fe-show") && app.classList.contains("is-max") && !document.fullscreenElement) setMax(false);
      cancelLink(); closePopovers(); selectOnly(null);
    }
  });

  var maxBtn = document.getElementById("fe-max");
  function setMax(on) {
    if (app.classList.contains("is-max") === on) return;
    app.classList.toggle("is-max", on);
    document.body.classList.toggle("flow-max-open", on);
    maxBtn.innerHTML = on ? '<i class="bi bi-fullscreen-exit"></i> 戻す' : '<i class="bi bi-arrows-fullscreen"></i> 最大表示';
    maxBtn.title = on ? "元の表示に戻す（Esc）" : "図を画面いっぱいに表示（Esc で戻る）";
    if (on) {
      if (app.requestFullscreen) app.requestFullscreen().then(function () { app.dataset.fs = "1"; }).catch(function () {});
    } else {
      if (document.fullscreenElement === app && document.exitFullscreen) document.exitFullscreen().catch(function () {});
      app.dataset.fs = "";
    }
    setTimeout(fitView, 80);
  }
  maxBtn.addEventListener("click", function () { setMax(!app.classList.contains("is-max")); });
  document.addEventListener("fullscreenchange", function () {
    if (document.fullscreenElement === app) setTimeout(fitView, 80);
    else if (app.dataset.fs === "1") setMax(false);
  });

  window.addEventListener("beforeunload", function (e) {
    if (dirty) { e.preventDefault(); e.returnValue = ""; }
  });

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

  var wbPop = document.getElementById("fe-wb-popover");
  var wbList = document.getElementById("fe-wb-list");
  var wbApi = cfg.api.replace(/\/graph$/, "/writeback");
  function writebackUrl() { return wbApi + apiUrl().slice(cfg.api.length); }

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
        cb.type = "checkbox"; cb.checked = c.checked !== false; cb.value = c.id; cb.className = "form-check-input me-1";
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
      var res = await fetch(wbApi, {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ kind: cfg.kind, workplace_id: cfg.workplace || null, domain_id: cfg.domain || null, ids: ids })
      });
      var data = await res.json();
      if (!res.ok || !data.ok) throw new Error(data.error || res.statusText);
      wbPop.classList.remove("fe-show");
      var done = "✅ 業務データに " + data.applied + " 件反映しました";
      if (data.ai_stale) done += "（AI 分析の案件 " + data.ai_stale + " 件は、案件詳細の「再分析」で新しくなります）";
      setMode(done + "。図を再生成します");
      await load({ keepView: true });
      setMode(done);
    } catch (err) {
      setMode("反映に失敗しました: " + err.message);
    }
  }
  document.getElementById("fe-writeback").addEventListener("click", openWriteback);
  document.getElementById("fe-wb-apply").addEventListener("click", applyWriteback);
  document.getElementById("fe-wb-close").addEventListener("click", function () { wbPop.classList.remove("fe-show"); });
  wbPop.addEventListener("mousedown", function (e) { e.stopPropagation(); });

  document.getElementById("fe-save").addEventListener("click", save);
  els.undo.addEventListener("click", undo);
  els.redo.addEventListener("click", redo);
  ["work", "info"].forEach(function (t) {
    document.getElementById("fe-add-" + t).addEventListener("click", function () {
      var r = els.container.getBoundingClientRect();
      var c = toCanvas(r.left + r.width / 2, r.top + r.height / 2);
      addNode(Math.round(c.x - 60), Math.round(c.y - 20), t);
    });
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
    flowchart: { htmlLabels: true, curve: "basis", nodeSpacing: 28, rankSpacing: 48, padding: 8, wrappingWidth: 1000 }
  });

  var scales = {}, before = {};
  var iconFont = document.fonts && document.fonts.load ? document.fonts.load('16px "bootstrap-icons"').catch(function () {}) : Promise.resolve();

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
    await iconFont;
    try {
      var out = await mermaid.render("svg-" + el.id, def);
      el.innerHTML = out.svg;
      if (out.bindFunctions) out.bindFunctions(el);
      var svg = el.querySelector("svg");
      var vb = svg && svg.viewBox && svg.viewBox.baseVal;
      var canvas = el.parentElement;
      if (vb && vb.width && canvas) {
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

  function fitScale(el) {
    var svg = el && el.querySelector("svg");
    var vb = svg && svg.viewBox && svg.viewBox.baseVal;
    var canvas = el && el.parentElement;
    if (!vb || !vb.width || !canvas) return null;
    return Math.max(0.2, Math.min(1.5, (canvas.clientWidth - 32) / vb.width, (canvas.clientHeight - 32) / vb.height));
  }

  function refit(key) {
    var el = document.getElementById("flow-" + key);
    var s = fitScale(el);
    if (s) { scales[key] = s; applyScale(el); }
  }

  function setMax(tb, on) {
    var key = tb.dataset.target;
    var card = tb.closest(".card");
    var btn = tb.querySelector("[data-max]");
    if (!card || card.classList.contains("flow-max") === on) return;
    card.classList.toggle("flow-max", on);
    document.body.classList.toggle("flow-max-open", on);
    if (btn) {
      btn.innerHTML = on ? '<i class="bi bi-fullscreen-exit"></i> 戻す' : '<i class="bi bi-arrows-fullscreen"></i> 最大表示';
      btn.title = on ? "元の表示に戻す（Esc）" : "図を画面いっぱいに表示";
    }
    if (on) {
      before[key] = scales[key] || 1;
      if (card.requestFullscreen) card.requestFullscreen().then(function () { card.dataset.fs = "1"; }).catch(function () {});
      setTimeout(function () { refit(key); }, 80);
    } else {
      if (document.fullscreenElement === card && document.exitFullscreen) document.exitFullscreen().catch(function () {});
      card.dataset.fs = "";
      scales[key] = before[key] || scales[key] || 1;
      applyScale(document.getElementById("flow-" + key));
    }
  }

  document.addEventListener("fullscreenchange", function () {
    document.querySelectorAll(".card.flow-max").forEach(function (card) {
      var tb = card.querySelector(".flow-toolbar");
      if (!tb) return;
      if (document.fullscreenElement === card) setTimeout(function () { refit(tb.dataset.target); }, 80);
      else if (card.dataset.fs === "1") setMax(tb, false);
    });
  });
  document.addEventListener("keydown", function (e) {
    if (e.key !== "Escape") return;
    document.querySelectorAll(".card.flow-max .flow-toolbar").forEach(function (tb) { setMax(tb, false); });
  });
  window.addEventListener("resize", function () {
    document.querySelectorAll(".card.flow-max .flow-toolbar").forEach(function (tb) { refit(tb.dataset.target); });
  });

  document.querySelectorAll(".flow-toolbar").forEach(function (tb) {
    var key = tb.dataset.target;
    var maxBtn = document.createElement("button");
    maxBtn.type = "button";
    maxBtn.className = "btn btn-outline-secondary flow-max-btn";
    maxBtn.dataset.max = "1";
    maxBtn.title = "図を画面いっぱいに表示";
    maxBtn.innerHTML = '<i class="bi bi-arrows-fullscreen"></i> 最大表示';
    tb.appendChild(maxBtn);
    tb.addEventListener("click", function (e) {
      var b = e.target.closest("button");
      if (!b) return;
      var el = document.getElementById("flow-" + key);
      if (b.dataset.max) { setMax(tb, !tb.closest(".card").classList.contains("flow-max")); return; }
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
    'style.css': r'''.step-list .step-done {
  text-decoration: line-through;
  color: #6c757d;
}
.chart-box {
  position: relative;
  height: 260px;
}
.table-sm td, .table-sm th {
  vertical-align: middle;
}
.flow-canvas {
  overflow: auto;
  border: 1px solid #dee2e6;
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
.lg-system { background: #f8f9fa; border-color: #6c757d; }
.lg-warn { background: #f8d7da; border-color: #dc3545; }
.lg-pending { background: #f8f9fa; border-color: #adb5bd; border-style: dashed; color: #6c757d; }
.lg-hold { background: #e7f1ff; border-color: #0d6efd; border-style: dashed; }
.lg-abolished { background: #f1f3f5; border-color: #ced4da; color: #adb5bd; }
.legend-row { display: flex; flex-wrap: wrap; align-items: center; row-gap: 0.25rem; }
.legend-title { font-size: 0.75rem; color: #6c757d; margin-right: 0.4rem; white-space: nowrap; }
.kind-legend { display: flex; flex-wrap: wrap; align-items: center; gap: 0.25rem 0.8rem; margin-top: 0.35rem; font-size: 0.78rem; color: #495057; }
.kind-chip { display: inline-flex; align-items: center; gap: 0.25rem; white-space: nowrap; }
.ks { width: 18px; height: 14px; flex: 0 0 auto; }
.ks rect, .ks path { fill: #fff; stroke: #6c757d; stroke-width: 1.2; }
.kind-chip .fi { margin-right: 0; }
.fi { margin-right: 0.2em; }
.fi.m-excel { color: #1d6f42; }
.fi.m-word { color: #2b579a; }
.fi.m-ppt { color: #c43e1c; }
.fi.m-pdf { color: #d93025; }
.fi.m-csv, .fi.done { color: #198754; }
.fi.m-paper, .fi.m-doc { color: #6c757d; }
.fi.m-photo, .fi.m-board { color: #6f42c1; }
.fi.m-system { color: #495057; }
.fi.m-web { color: #0aa2c0; }
.fi.m-mail { color: #0d6efd; }
.fi.m-chat { color: #5b5fc7; }
.fi.m-verbal { color: #b35c00; }
.fi.m-unknown, .fi.who-unconfirmed, .fi.pending { color: #adb5bd; }
.map-legend { margin-top: 0.15rem; }
.lg-item { display: inline-flex; align-items: center; gap: 0.3rem; font-size: 0.78rem; color: #495057; margin-right: 0.85rem; white-space: nowrap; }
.lg-line { display: inline-block; width: 26px; height: 0; vertical-align: middle; }
.lg-line-work { border-top: 3px solid #495057; }
.lg-line-info { border-top: 1.5px solid #5b8def; }
.lg-line-guess { border-top: 2px dashed #6c757d; }
.lg-work { background: #fff; border-color: #6c757d; }
.lg-source { background: #fff3cd; border-color: #e0a800; }
.lg-output { background: #fff; border-color: #0d6efd; border-style: solid; }
.flow-max-btn { white-space: nowrap; }
.card.flow-max { position: fixed; inset: 0; z-index: 1050; margin: 0; border-radius: 0; background: #fff; }
.card.flow-max > .card-body { display: flex; flex-direction: column; height: 100%; min-height: 0; }
.card.flow-max .flow-canvas { flex: 1 1 auto; max-height: none; min-height: 0; }
body.flow-max-open { overflow: hidden; }
/* ---- 業務マップ 編集エディタ ---- */
.fe-wrap { display: flex; flex-direction: column; height: calc(100vh - 150px); min-height: 520px; }
.fe-wrap.is-max { position: fixed; inset: 0; z-index: 1050; height: auto; min-height: 0; padding: 8px 12px; background: #fff; }
.fe-kind-legend { margin: 0 0 0.3rem; }
.fe-toolbar { display: flex; flex-wrap: wrap; justify-content: space-between; align-items: center; gap: 0.5rem; margin-bottom: 0.4rem; }
.fe-buttons { display: flex; flex-wrap: wrap; align-items: center; gap: 0.3rem; }
.fe-sep { width: 1px; height: 1.4rem; background: #dee2e6; margin: 0 0.2rem; }
.fe-status { font-size: 0.8rem; margin-left: 0.6rem; }
.fe-hint { font-size: 0.78rem; color: #6c757d; margin-bottom: 0.3rem; }
.mv-move { display: flex; gap: 2px; }
.mv-move .btn { padding: 0 0.4rem; line-height: 1.3; font-size: 0.7rem; }
.mv-words { min-width: 16rem; }
.mv-table td:nth-child(2) { min-width: 11rem; }
.mv-table td:nth-child(4) { min-width: 7rem; }
.mv-icon { font-size: 1.15rem; width: 1.3rem; text-align: center; }
.mv-savebar { position: sticky; bottom: 0; z-index: 5; display: flex; gap: 0.5rem; align-items: center; padding: 0.75rem 0; background: linear-gradient(to top, #f5f6f8 70%, rgba(245, 246, 248, 0)); }
.member-wrap { overflow: auto; max-height: 72vh; }
.member-table th, .member-table td { white-space: nowrap; }
.member-table thead th { position: sticky; top: 0; z-index: 1; background: #fff; }
.member-month { font-size: 0.78rem; line-height: 1.15; }
.member-cell { min-width: 2.6rem; }
.member-yes { color: var(--ui-green-text); font-size: 1.05rem; }
.member-before { color: #ced4da; }
.member-table tr.is-excluded td { color: #adb5bd; background: #f8f9fa; }
.member-table tr.is-excluded .member-yes { color: #adb5bd; }
.member-table tfoot th, .member-table tfoot td { background: #f8f9fa; font-size: 0.8rem; line-height: 1.2; }
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
.fe-edge.fe-type-info { stroke: #5b8def; stroke-width: 1.6px; }
.fe-edge.fe-type-work { stroke: #495057; stroke-width: 2.6px; }
.fe-edge.fe-inferred { stroke-dasharray: 6 4; }
.fe-who { font-size: 11px; font-weight: 700; color: #495057; line-height: 1.2; margin-bottom: 2px; }
.fe-cls-work { --fe-fill: #fff; --fe-stroke: #6c757d; color: #343a40; }
.fe-type-row { display: flex; align-items: center; }
.fe-type-row .btn.active { background: #e7f1ff; border-color: #0d6efd; color: #052c65; }
.fe-node.fe-manual { border-style: solid; }
.fe-node.fe-manual::after { content: "手"; position: absolute; left: -7px; top: -8px; font-size: 9px; line-height: 1.4; padding: 0 4px; border-radius: 6px; background: #6f42c1; color: #fff; }
.fe-shape-rect { border-radius: 2px; }
.fe-edge-label { font-size: 11px; fill: #343a40; paint-order: stroke; stroke: #fff; stroke-width: 3px; pointer-events: none; }
.fe-temp-line { fill: none; stroke: #60a5fa; stroke-width: 2px; stroke-dasharray: 4 4; }
#fe-nodes { position: absolute; left: 0; top: 0; }
.fe-node { position: absolute; width: max-content; min-width: 90px; max-width: 150px; padding: 6px 10px; border: 2px solid var(--fe-stroke, #0d6efd); border-radius: 8px; background: var(--fe-fill, #fff);
  font-size: 13px; line-height: 1.3; text-align: center; cursor: move; user-select: none; box-shadow: 0 2px 4px rgba(0,0,0,0.08); isolation: isolate; }
.fe-label { display: flex; align-items: center; justify-content: center; gap: 4px; }
.fe-icon { flex: 0 0 auto; font-size: 14px; }
.fe-shape-stadium { border-radius: 999px; padding-left: 14px; padding-right: 14px; }
.fe-shape-cyl { border-radius: 50% / 11px; padding-top: 20px; padding-bottom: 8px; }
.fe-shape-cyl::before { content: ""; position: absolute; left: -2px; right: -2px; top: -2px; height: 22px; border: 2px solid transparent; border-bottom-color: var(--fe-stroke, #6c757d); border-radius: 50%; pointer-events: none; }
.fe-shape-doc { background: transparent; border-color: transparent; box-shadow: none; padding-bottom: 14px; }
.fe-shape-bg { position: absolute; left: -2px; top: -2px; width: calc(100% + 4px); height: calc(100% + 4px); overflow: visible; pointer-events: none; z-index: -1; filter: drop-shadow(0 2px 2px rgba(0,0,0,0.08)); }
.fe-shape-bg path { fill: var(--fe-fill, #fff); stroke: var(--fe-stroke, #0d6efd); stroke-width: 2px; vector-effect: non-scaling-stroke; }
.fe-cls-pending .fe-shape-bg path { stroke-dasharray: 6 3; }
.fe-node:hover { box-shadow: 0 4px 10px rgba(0,0,0,0.15); z-index: 5; }
.fe-node.fe-selected { outline: 3px solid rgba(25, 135, 84, 0.45); outline-offset: 1px; z-index: 6; }
.fe-node.fe-link-source { outline: 3px solid rgba(246, 173, 85, 0.8); outline-offset: 1px; }
.fe-node .fe-text { white-space: pre-wrap; word-break: break-word; min-width: 40px; }
.fe-node .fe-text.fe-editing { outline: 1px solid #0d6efd; background: #fff; cursor: text; user-select: text; }
.fe-node .fe-sub { font-size: 10px; color: #6c757d; margin-top: 2px; }
.fe-node .fe-edited { position: absolute; right: 3px; bottom: 1px; font-size: 9px; color: #6f42c1; }
.fe-btns { display: none; position: absolute; top: -26px; left: 50%; transform: translateX(-50%); gap: 3px; white-space: nowrap; z-index: 7; }
.fe-node:hover .fe-btns, .fe-node.fe-selected .fe-btns { display: flex; }
.fe-btn { width: 22px; height: 22px; border: 1px solid #ced4da; border-radius: 50%; background: #fff; font-size: 11px; line-height: 1; padding: 0; cursor: pointer; box-shadow: 0 1px 2px rgba(0,0,0,0.15); }
.fe-btn:hover { transform: scale(1.1); }
.fe-btn-del { color: #dc3545; }
.fe-cls-task { --fe-fill: #e7f1ff; --fe-stroke: #0d6efd; color: #052c65; }
.fe-cls-done { --fe-fill: #d1e7dd; --fe-stroke: #198754; color: #0f5132; }
.fe-cls-source { --fe-fill: #fff3cd; --fe-stroke: #e0a800; color: #664d03; }
.fe-cls-output { --fe-fill: #fff; --fe-stroke: #0d6efd; color: #052c65; }
.fe-cls-dupout { --fe-fill: #ffe5d0; --fe-stroke: #fd7e14; }
.fe-cls-warn { --fe-fill: #f8d7da; --fe-stroke: #dc3545; color: #842029; }
.fe-cls-pending { --fe-fill: #f8f9fa; --fe-stroke: #adb5bd; border-style: dashed; color: #6c757d; }
.fe-cls-hold { --fe-fill: #e7f1ff; --fe-stroke: #0d6efd; border-style: dashed; color: #052c65; }
.fe-cls-abolished { --fe-fill: #f1f3f5; --fe-stroke: #ced4da; color: #adb5bd; }
.fe-node .fe-state { font-size: 11px; white-space: nowrap; }
.fe-node[data-color="white"] { --fe-fill: #fff; }
.fe-node[data-color="red"] { --fe-fill: #fee2e2; }
.fe-node[data-color="orange"] { --fe-fill: #fed7aa; }
.fe-node[data-color="yellow"] { --fe-fill: #fef3c7; }
.fe-node[data-color="green"] { --fe-fill: #d1fae5; }
.fe-node[data-color="blue"] { --fe-fill: #dbeafe; }
.fe-node[data-color="purple"] { --fe-fill: #e9d5ff; }
.fe-node[data-color="pink"] { --fe-fill: #fce7f3; }
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
/* ---- 進捗掲示板 ---- */
.pb-head { margin: 0.25rem 0 0.75rem; display: flex; flex-wrap: wrap; justify-content: space-between; align-items: flex-start; gap: 0.75rem 1.5rem; }
.pb-head-main { flex: 1 1 420px; min-width: 0; }
.pb-link { color: var(--ui-accent); text-decoration: none; font-size: 0.9rem; white-space: nowrap; }
.pb-link:hover { text-decoration: underline; }
.pb-scope-form { display: inline-block; max-width: 100%; }
.pb-h1 { font-weight: 700; margin: 0; white-space: nowrap; display: flex; align-items: baseline; gap: 0.1em; flex-wrap: wrap; }
.pb-scope-select { font: inherit; font-weight: 700; letter-spacing: inherit; color: var(--ui-accent); background: transparent; border: 0; border-bottom: 3px solid rgba(0, 113, 227, 0.3); border-radius: 0; padding: 0 1.15em 0 0; margin: 0; appearance: none; -webkit-appearance: none; cursor: pointer; max-width: 80vw; background-image: url("data:image/svg+xml;charset=utf-8,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 16 16'%3E%3Cpath fill='none' stroke='%230071e3' stroke-width='2.2' stroke-linecap='round' stroke-linejoin='round' d='M3 6l5 5 5-5'/%3E%3C/svg%3E"); background-repeat: no-repeat; background-position: right 0.05em center; background-size: 0.6em; field-sizing: content; }
.pb-scope-select:hover { border-bottom-color: var(--ui-accent); }
.pb-scope-select:focus { outline: none; border-bottom-color: var(--ui-accent); }
.pb-section { display: flex; align-items: baseline; gap: 0.4rem; }
.pb-section .pb-link { margin-left: auto; font-weight: 400; letter-spacing: 0; }
.pb-row { display: grid; grid-template-columns: 170px minmax(220px, 1fr) minmax(460px, 2.2fr); gap: 1.25rem; align-items: center;
  background: #fff; padding: 1rem 1.25rem; margin-bottom: 0.6rem; }
.pb-wp { font-weight: 600; font-size: 0.95rem; }
.pb-name { color: #6c757d; font-size: 0.9rem; }
.pb-title { font-weight: 600; line-height: 1.35; }
.pb-title a { color: inherit; text-decoration: none; }
.pb-title a:hover { text-decoration: underline; }
.pb-note { color: #6c757d; font-size: 0.88rem; margin-top: 0.2rem; }
.pb-pill { display: inline-block; vertical-align: middle; font-size: 0.72rem; font-weight: 600; padding: 0.1rem 0.5rem; border-radius: 999px; background: #fff3cd; color: #997404; margin-left: 0.4rem; }
.pb-empty { color: #adb5bd; font-size: 0.9rem; padding: 0.5rem 0; }
/* 順番待ち */
.pb-waiting { list-style: none; margin: 0 0 1rem; padding: 0; background: #fff; }
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
/* スマホ: 順番待ちの行は縦に積む */
@media (max-width: 575.98px) {
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
.app-nav .nav-link { color: var(--ui-text); font-weight: 500; border-radius: 999px; white-space: nowrap; }
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
.card { border: 0; border-radius: var(--ui-radius); background: var(--ui-surface); }
.card > .card-body { padding: 1.4rem 1.5rem; }
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
.kpi-card .kpi-value { font-size: 2.1rem; font-weight: 700; line-height: 1.1; letter-spacing: -0.03em; font-variant-numeric: tabular-nums; color: var(--ui-text); }
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
.form-control, .form-select { border-radius: var(--ui-radius-sm); border: 1px solid transparent; background-color: var(--ui-fill); color: var(--ui-text); padding: 0.55rem 0.85rem; }
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
.suggestion-kind { font-size: 0.75rem; font-weight: 600; }

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

/* ============================================================
   仕上げ: ページの切り替え・スクロールで現れる・押した手応え・お知らせ・送信中・紙吹雪
   ============================================================ */
/* 背景にうっすら色を差す（上のほうだけ） */
body {
  background:
    radial-gradient(1100px 520px at 0% -10%, rgba(10, 132, 255, 0.07), transparent 62%),
    radial-gradient(900px 480px at 100% -14%, rgba(191, 90, 242, 0.06), transparent 60%),
    var(--ui-bg);
  background-repeat: no-repeat;
}
::selection { background: rgba(0, 113, 227, 0.18); }
a:focus-visible, .nav-link:focus-visible, summary:focus-visible, [tabindex]:focus-visible, .nav-admin:focus-visible { outline: 3px solid rgba(0, 113, 227, 0.35); outline-offset: 2px; border-radius: 8px; }
.section-title::before { content: ""; align-self: center; width: 4px; height: 1.05em; border-radius: 4px; background: linear-gradient(180deg, var(--c-blue), var(--c-indigo)); flex: 0 0 auto; }
.card { box-shadow: var(--ui-shadow), inset 0 0 0 1px rgba(0, 0, 0, 0.025); }
.brand-mark { transition: transform 0.35s cubic-bezier(0.3, 1.4, 0.4, 1); }
.navbar-brand:hover .brand-mark { transform: rotate(-8deg) scale(1.08); }

/* ページの切り替え（対応ブラウザだけ）。ヘッダーは動かさず、いまのページの印が次のページへ滑って移る */
@view-transition { navigation: auto; }
@keyframes ui-vt-out { to { opacity: 0; transform: translateY(-4px); } }
@keyframes ui-vt-in { from { opacity: 0; transform: translateY(10px); } }
::view-transition-old(root) { animation: ui-vt-out 0.16s ease both; }
::view-transition-new(root) { animation: ui-vt-in 0.34s cubic-bezier(0.2, 0.7, 0.2, 1) both; }
.app-header { view-transition-name: app-header; }
.app-nav .nav-link.active { view-transition-name: nav-current; }
::view-transition-group(nav-current) { animation-duration: 0.42s; animation-timing-function: cubic-bezier(0.3, 1.25, 0.4, 1); }

/* スクロールして見えてきたカードだけ、ふわっと出す（JS が画面の外のものに印を付ける） */
.ui-reveal { animation: none !important; opacity: 0; transform: translateY(18px); }
.ui-reveal.is-in { opacity: 1; transform: none; transition: opacity 0.6s cubic-bezier(0.2, 0.7, 0.2, 1), transform 0.6s cubic-bezier(0.2, 0.7, 0.2, 1); }
.ui-reveal.is-done { transition: background-color 0.18s ease, box-shadow 0.22s ease, transform 0.18s ease; }

/* ヘッダー: スクロールしたら薄い影 */
.app-header { transition: box-shadow 0.25s ease; }
html.is-scrolled .app-header { box-shadow: 0 1px 0 rgba(0, 0, 0, 0.02), 0 8px 24px rgba(0, 0, 0, 0.06); }

/* 押した手応え（押した所から広がる波紋） */
.btn { position: relative; }
.ui-ripple { position: absolute; inset: 0; border-radius: inherit; overflow: hidden; pointer-events: none; }
.ui-ripple > span { position: absolute; border-radius: 50%; background: currentColor; opacity: 0.2; transform: scale(0); animation: ui-ripple 0.6s cubic-bezier(0.2, 0.7, 0.2, 1) forwards; }
@keyframes ui-ripple { to { transform: scale(1); opacity: 0; } }

/* 送信中のボタン（くるくる） */
@keyframes ui-spin { to { transform: rotate(360deg); } }
.btn.is-busy { pointer-events: none; opacity: 0.85; }
.btn.is-busy::after { content: ""; display: inline-block; width: 0.95em; height: 0.95em; margin-left: 0.5em; vertical-align: -0.12em; border: 2px solid currentColor; border-right-color: transparent; border-radius: 50%; animation: ui-spin 0.7s linear infinite; }

/* お知らせ（保存しました など）: 右上に出て数秒で消える。触れている間は止まる。注意・エラーは本文に残す */
.ui-toasts { position: fixed; top: calc(var(--ui-header-h, 56px) + 12px); right: 16px; z-index: 1080; display: flex; flex-direction: column; gap: 0.5rem; width: min(420px, calc(100vw - 32px)); pointer-events: none; }
.ui-toasts .alert { position: relative; pointer-events: auto; margin: 0; padding-right: 2.75rem; background: rgba(255, 255, 255, 0.92); -webkit-backdrop-filter: saturate(180%) blur(16px); backdrop-filter: saturate(180%) blur(16px); color: var(--ui-text); box-shadow: 0 1px 2px rgba(0, 0, 0, 0.06), 0 12px 32px rgba(0, 0, 0, 0.14); animation: ui-toast-in 0.5s cubic-bezier(0.2, 0.9, 0.3, 1.15) both; overflow: hidden; }
.ui-toasts .alert > .bi { margin-right: 0.45rem; }
.ui-toasts .alert-success > .bi { color: var(--ui-green); }
.ui-toasts .alert-info > .bi, .ui-toasts .alert-secondary > .bi { color: var(--ui-accent); }
.ui-toasts .alert::after { content: ""; position: absolute; left: 0; right: 0; bottom: 0; height: 3px; background: currentColor; opacity: 0.18; transform-origin: left center; animation: ui-toast-timer 4.5s linear 0.4s forwards; }
.ui-toasts .alert-success::after { background: var(--ui-green); opacity: 0.45; }
.ui-toasts .alert:hover::after { animation-play-state: paused; }
.ui-toasts .alert.is-leaving { animation: ui-toast-out 0.32s ease forwards; }
@keyframes ui-toast-in { from { opacity: 0; transform: translateX(28px) scale(0.98); } }
@keyframes ui-toast-out { to { opacity: 0; transform: translateX(28px); } }
@keyframes ui-toast-timer { from { transform: scaleX(1); } to { transform: scaleX(0); } }
main > .alert-warning, main > .alert-danger { animation: ui-nudge 0.5s cubic-bezier(0.3, 1.4, 0.4, 1) both; }
@keyframes ui-nudge { 0% { opacity: 0; transform: translateX(0); } 30% { opacity: 1; transform: translateX(-6px); } 60% { transform: translateX(4px); } 100% { transform: none; } }

/* 紙吹雪（回答・登録が終わったとき） */
.ui-confetti { position: fixed; inset: 0; width: 100vw; height: 100vh; pointer-events: none; z-index: 1090; }

/* 図を描いている間は、光が流れる下地を出す */
@keyframes ui-shimmer { from { background-position: 200% 0; } to { background-position: -200% 0; } }
.flow-inner > .text-muted.small:only-child { display: flex; align-items: center; justify-content: center; min-height: 160px; border-radius: 12px; background: linear-gradient(90deg, #f2f2f5 0%, #fbfbfd 40%, #f2f2f5 80%); background-size: 200% 100%; animation: ui-shimmer 1.6s linear infinite; }

/* 最大表示は、ふわっと広がる */
@keyframes ui-max-in { from { opacity: 0.35; transform: scale(0.985); } }
.card.flow-max, .fe-wrap.is-max { animation: ui-max-in 0.28s cubic-bezier(0.2, 0.7, 0.2, 1) both; }

/* 開閉（詳細の記録など）を滑らかに */
:root { interpolate-size: allow-keywords; }
details::details-content { block-size: 0; overflow-y: clip; transition: block-size 0.35s cubic-bezier(0.2, 0.7, 0.2, 1), content-visibility 0.35s allow-discrete; }
details[open]::details-content { block-size: auto; }
summary { transition: color 0.15s ease; }

/* 選んだ選択肢は、ぽんと弾む */
@keyframes ui-pick { 0% { transform: scale(1); } 40% { transform: scale(1.07); } 100% { transform: scale(1); } }
.mq-option input:checked + span { animation: ui-pick 0.32s cubic-bezier(0.3, 1.4, 0.4, 1); }

/* 月次アンケート: 進み具合（あと何問か） */
.mq-progress { position: sticky; top: calc(var(--ui-header-h, 56px) + 10px); z-index: 4; display: flex; align-items: center; gap: 0.8rem; margin: 0 0 0.75rem; padding: 0.5rem 0.95rem; border-radius: 999px; background: rgba(255, 255, 255, 0.9); -webkit-backdrop-filter: saturate(180%) blur(14px); backdrop-filter: saturate(180%) blur(14px); box-shadow: 0 1px 2px rgba(0, 0, 0, 0.05), 0 6px 18px rgba(0, 0, 0, 0.07); }
.mq-progress-bar { flex: 1 1 auto; height: 6px; border-radius: 999px; background: var(--ui-fill); overflow: hidden; }
.mq-progress-bar > span { display: block; height: 100%; border-radius: inherit; background: linear-gradient(90deg, var(--c-blue), var(--c-indigo)); transform-origin: left center; transform: scaleX(var(--p, 0)); transition: transform 0.5s cubic-bezier(0.2, 0.7, 0.2, 1), background 0.3s ease; }
.mq-progress-text { flex: 0 0 auto; font-size: 0.82rem; font-weight: 600; color: var(--ui-text-2); font-variant-numeric: tabular-nums; }
.mq-progress.is-complete .mq-progress-bar > span { background: linear-gradient(90deg, #28b24a, var(--c-green)); }
.mq-progress.is-complete .mq-progress-text { color: var(--ui-green-text); }
@keyframes ui-glow { 0%, 100% { box-shadow: 0 0 0 0 rgba(0, 113, 227, 0); } 50% { box-shadow: 0 0 0 6px rgba(0, 113, 227, 0.16); } }
.mq-ready #mq-submit { animation: ui-glow 1.8s ease-in-out 2; }

@media (prefers-reduced-motion: reduce) {
  ::view-transition-group(*), ::view-transition-old(*), ::view-transition-new(*) { animation: none !important; }
  .ui-reveal { opacity: 1 !important; transform: none !important; }
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
.mq-emp-ok { color: var(--ui-green-text); font-weight: 600; }
.mq-emp-ng { color: #dc3545; }
.mq-q { padding: 1.1rem 0; border-bottom: 1px solid var(--ui-line); }
.mq-q:last-of-type { border-bottom: 0; }
.mq-q-head { display: flex; align-items: baseline; gap: 0.6rem; margin-bottom: 0.7rem; }
.mq-no { display: inline-flex; align-items: center; justify-content: center; min-width: 1.7rem; height: 1.7rem; background: var(--ui-text); color: #fff; font-size: 0.8rem; font-weight: 700; flex: 0 0 auto; align-self: center; }
.mq-text { font-weight: 600; font-size: 1.02rem; line-height: 1.4; }
.mq-optional { font-size: 0.72rem; color: var(--ui-text-3); border: 1px solid var(--ui-line); border-radius: 4px; padding: 0 0.35rem; align-self: center; white-space: nowrap; }
.mq-required { font-size: 0.72rem; color: var(--ui-red-text); background: var(--ui-red-soft); border-radius: 4px; padding: 0 0.35rem; align-self: center; white-space: nowrap; font-weight: 700; }
.mq-options { display: flex; flex-wrap: wrap; gap: 0.5rem; }
.mq-list-input { display: flex; flex-direction: column; gap: 0.45rem; }
.mq-list-row { display: flex; align-items: center; gap: 0.35rem; }
.mq-list-row .form-control { flex: 1; }
.mq-list-del { border: 0; background: none; color: var(--ui-text-3); font-size: 1.25rem; line-height: 1; padding: 0 0.35rem; cursor: pointer; }
.mq-list-del:hover { color: var(--ui-red-text); }
.mq-list-actions { display: flex; flex-wrap: wrap; gap: 0.4rem; }
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
/* ---- 自動化を検討（1画面） ---- */
.rq { max-width: 1080px; margin: 0 auto; }
.rq-top { display: flex; flex-wrap: wrap; align-items: center; justify-content: space-between; gap: 0.75rem 1.5rem; margin-bottom: 1rem; }
.rq-title { font-size: 1.6rem; margin: 0; }
.rq-no { display: inline-flex; align-items: center; justify-content: center; width: 1.5rem; height: 1.5rem; border-radius: 50%; background: #fff; color: var(--ui-text-2); font-size: 0.78rem; font-weight: 700; }
.rq-opt { font-size: 0.66rem; border: 1px solid currentColor; border-radius: 4px; padding: 0 0.3rem; opacity: 0.7; }
.rq-card .card-body { padding: 1.5rem 1.75rem; }
.rq-field { margin-bottom: 1.15rem; }
.rq-label { display: flex; align-items: center; gap: 0.45rem; font-weight: 600; margin-bottom: 0.45rem; font-size: 0.95rem; }
.rq-req { font-size: 0.66rem; color: var(--ui-red-text); background: var(--ui-red-soft); border-radius: 4px; padding: 0.05rem 0.35rem; font-weight: 700; }
.rq-pills { display: flex; flex-wrap: wrap; gap: 0.45rem; }
.rq-pill { border: 1.5px solid transparent; background: var(--ui-fill); color: var(--ui-text); border-radius: 999px; padding: 0.45rem 0.95rem; font-size: 0.92rem; font-weight: 500; cursor: pointer; transition: all 0.15s ease; }
.rq-pill:hover { background: #e3e3e8; }
.rq-pill:active { transform: scale(0.97); }
.rq-pill.is-on { background: var(--ui-accent); color: #fff; box-shadow: 0 2px 8px rgba(0, 113, 227, 0.3); }
.rq-pill:focus-visible { outline: none; box-shadow: 0 0 0 4px rgba(0, 113, 227, 0.2); }
.rq-field.is-invalid .form-control, .rq-field.is-invalid .form-select { border-color: var(--c-red); background: #fff; }
.rq-field.is-invalid .rq-pill:not(.is-on) { border-color: rgba(255, 69, 58, 0.45); }
.rq-err { color: var(--ui-red-text); font-size: 0.8rem; margin-top: 0.35rem; }
.rq-err:empty { display: none; }
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
.rq-finish { display: flex; justify-content: center; margin: 1.5rem 0; }
.rq-finish .btn { min-width: 260px; }
.rq-done { animation: ui-fade-up 0.4s cubic-bezier(0.2, 0.7, 0.2, 1) both; }
.rq-id-row { display: flex; align-items: flex-end; gap: 1rem; flex-wrap: wrap; }
.rq-id-row .rq-field { margin-bottom: 0.5rem; min-width: 220px; }
.rq-identity { font-weight: 600; font-size: 0.9rem; color: var(--ui-green-text); padding-bottom: 0.9rem; }
.rq-identity.is-ng { color: #dc3545; font-weight: 500; }
.rq-identity:empty { display: none; }
/* ---- ヘッダー1段: 管理者タブは藍色で区別 ---- */
.app-header { z-index: 1030; }
.app-nav .nav-link { font-size: 0.88rem; padding: 0.38rem 0.6rem; }
.nav-sep { width: 1px; height: 1.25rem; background: #d6d6dc; margin: 0 0.4rem 0 0.5rem; align-self: center; }
.nav-lock { align-self: center; color: #4f46e5; font-size: 0.8rem; margin-right: 0.1rem; display: inline-flex; align-items: center; gap: 0.3rem; }
.app-nav .nav-link.nav-link-admin { color: #4338ca; }
.app-nav .nav-link.nav-link-admin:hover { background: rgba(79, 70, 229, 0.08); color: #312e81; }
.app-nav .nav-link.nav-link-admin.active { background: #4f46e5; color: #fff; }
/* 右上の管理者メニュー（目立たせない） */
.nav-admin { background: none; border: 0; padding: 0.4rem 0.5rem; font-size: 0.78rem; color: var(--ui-text-3); text-decoration: none; white-space: nowrap; border-radius: 999px; margin-left: 0.5rem; }
.nav-admin:hover { color: var(--ui-text-2); background: rgba(0, 0, 0, 0.04); text-decoration: none; }
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

/* ---- 進捗掲示板: 成果と検討への導線 ---- */
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
.lg-dupout { background: #fff3cd; border-color: #fd7e14; border-style: dashed; border-width: 2px; }
.fe-cls-dupout { background: #fff3cd; border-color: #fd7e14; border-style: dashed; color: #664d03; transform: skewX(-6deg); }
.fe-cls-dupout .fe-text { transform: skewX(6deg); }

/* ---- 自動化を検討: 業務の選択・月次アンケートからの続き・AI との聞き取り ---- */
.rq-lead { color: var(--ui-text-2); margin: 0.25rem 0 0; max-width: 46rem; }
.rq-sec-sub { margin-left: auto; font-size: 0.85rem; color: var(--ui-text-2); }
.rq-registered { display: flex; flex-direction: column; gap: 0.25rem; margin-bottom: 0.75rem; }
.rq-reg { display: flex; align-items: center; gap: 0.5rem; padding: 0.35rem 0.6rem; background: var(--ui-fill); border-radius: var(--ui-radius-sm); }
.rq-reg-name { font-weight: 600; }
.rq-reg-meta { color: var(--ui-text-2); font-size: 0.85rem; margin-left: auto; white-space: nowrap; }
.rq-suggest { display: flex; flex-wrap: wrap; gap: 0.4rem; align-items: center; margin-bottom: 0.6rem; }
.rq-suggest-label { font-size: 0.85rem; color: var(--ui-text-2); }
.rq-suggest-row { display: flex; flex-wrap: wrap; gap: 0.4rem; align-items: center; width: 100%; }
.rq-from { background: var(--ui-accent-soft); border-radius: var(--ui-radius-sm); padding: 0.7rem 0.9rem; margin-bottom: 0.9rem; }
.rq-from-title { font-weight: 700; color: var(--ui-accent-text); margin-bottom: 0.15rem; }
.rq-from-list { display: flex; flex-wrap: wrap; gap: 0.35rem; margin-top: 0.45rem; }
.rq-from-item { background: var(--ui-surface); border-radius: 999px; padding: 0.15rem 0.7rem; font-size: 0.85rem; }
.rq-from-item.is-done { color: var(--ui-green-text); }
.rq-from-item.is-now { font-weight: 700; color: var(--ui-accent-text); box-shadow: 0 0 0 1.5px var(--ui-accent); }
.mq-next { margin-top: 1.1rem; padding-top: 1rem; border-top: 1px solid var(--ui-line); }
.mq-next-title { font-weight: 700; margin-bottom: 0.45rem; }
.mq-next-tasks { display: flex; flex-wrap: wrap; justify-content: center; gap: 0.35rem; margin-bottom: 0.6rem; }
.mq-next-tasks span { background: var(--ui-fill); border-radius: 999px; padding: 0.2rem 0.75rem; font-size: 0.88rem; }
.rq-pick-row { display: flex; gap: 0.5rem; }
.rq-pick-row .btn { white-space: nowrap; }
.rq-intake { display: grid; grid-template-columns: 1.1fr 1fr; gap: 1rem; align-items: start; }
.rq-flowpane { position: sticky; top: 72px; background: var(--ui-fill); border-radius: var(--ui-radius); padding: 0.75rem; min-height: 16rem; transition: box-shadow 0.4s; }
.rq-flowpane.is-done { box-shadow: 0 0 0 3px var(--ui-green-soft), 0 0 0 5px var(--ui-green); }
.rq-flowhead { display: flex; align-items: center; justify-content: space-between; gap: 0.5rem; margin-bottom: 0.5rem; }
.rq-flowtitle { font-weight: 700; }
.rq-flowhead-right { display: flex; align-items: center; gap: 0.35rem; }
.rq-replay { padding: 0 0.25rem; font-size: 0.8rem; text-decoration: none; }
.rq-note { margin-top: 0.5rem; max-width: 440px; align-self: center; font-size: 0.78rem; color: var(--ui-accent-text); background: var(--ui-accent-soft); border-radius: 999px; padding: 0.15rem 0.7rem; cursor: pointer; }
.rq-note.is-new { animation: rq-slide 0.35s ease-out; }
.rq-stepcount { font-size: 0.8rem; font-weight: 600; color: var(--ui-accent-text); background: var(--ui-accent-soft); border-radius: 999px; padding: 0.15rem 0.6rem; display: inline-block; }
.rq-stepcount.is-bump { animation: rq-bump 0.5s ease-out; }
.rq-flowbox { background: var(--ui-surface); border-radius: var(--ui-radius-sm); padding: 0.5rem; overflow: auto; min-height: 9rem; display: flex; align-items: center; justify-content: center; }
.rq-flowbox svg { max-width: 100%; height: auto; }
.rq-flow-empty { color: var(--ui-text-3); font-size: 0.9rem; }
.rq-steps { margin: 0.6rem 0 0; padding-left: 1.3rem; font-size: 0.88rem; }
.rq-steps li { margin: 0.15rem 0; }
.rq-steps li.is-new { animation: rq-slide 0.5s ease-out; }
.rq-node-new { animation: rq-pop 0.6s ease-out; transform-box: fill-box; transform-origin: center; }
.rq-chatpane { display: flex; flex-direction: column; gap: 0.5rem; min-height: 16rem; }
.rq-chathead { display: flex; justify-content: space-between; align-items: center; font-size: 0.85rem; color: var(--ui-text-2); }
.rq-messages { display: flex; flex-direction: column; gap: 0.45rem; max-height: 22rem; overflow-y: auto; padding: 0.25rem 0.1rem; }
.rq-msg { max-width: 88%; padding: 0.5rem 0.75rem; border-radius: 14px; font-size: 0.95rem; line-height: 1.45; white-space: pre-wrap; word-break: break-word; }
.rq-msg.is-ai { background: var(--ui-fill); color: var(--ui-text); align-self: flex-start; border-bottom-left-radius: 4px; }
.rq-msg.is-user { background: var(--ui-accent); color: #fff; align-self: flex-end; border-bottom-right-radius: 4px; }
.rq-msg.is-new { animation: rq-slide 0.35s ease-out; }
.rq-typing { display: flex; gap: 4px; align-items: center; padding: 0.7rem 0.9rem; }
.rq-typing span { width: 7px; height: 7px; border-radius: 50%; background: var(--ui-text-3); animation: rq-dot 1.2s infinite; }
.rq-typing span:nth-child(2) { animation-delay: 0.2s; }
.rq-typing span:nth-child(3) { animation-delay: 0.4s; }
.rq-chips { display: flex; flex-wrap: wrap; gap: 0.4rem; align-items: center; }
.rq-chips-label { width: 100%; font-size: 0.75rem; color: var(--ui-text-3); }
.rq-chip { animation: rq-slide 0.3s ease-out; background: var(--ui-surface); border-color: var(--ui-accent); color: var(--ui-accent-text); font-weight: 600; }
.rq-chip:hover { background: var(--ui-accent-soft); }
.rq-chip .bi { font-size: 0.85em; margin-right: 0.25rem; }
.rq-inputrow { display: flex; gap: 0.4rem; }
.rq-tools { display: flex; gap: 0.4rem; flex-wrap: wrap; }
.rq-confirm { background: var(--ui-green-soft); border-radius: var(--ui-radius-sm); padding: 0.75rem; animation: rq-slide 0.4s ease-out; }
.rq-confirm-q { font-weight: 700; margin-bottom: 0.5rem; }
.rq-final { margin-top: 1rem; padding-top: 1rem; border-top: 1px solid var(--ui-line); }
.rq-dup { background: var(--ui-yellow-soft); color: var(--ui-yellow-text); border-radius: var(--ui-radius-sm); padding: 0.6rem 0.8rem; margin: 0.6rem 0; font-size: 0.92rem; }
.rq-dup .btn { margin: 0.4rem 0.4rem 0 0; }
.rq-final-title { font-weight: 700; margin-bottom: 0.5rem; }
.rq-self { margin: 0.5rem 0; }
.rq-result { text-align: center; padding: 1rem 0 0.5rem; }
.rq-impression { margin-top: 0.4rem; color: var(--ui-accent-text); font-weight: 600; }
.rq-log { display: flex; flex-direction: column; gap: 0.2rem; }
.rq-log-ai { color: var(--ui-text-2); }
.rq-log-user { padding-left: 1rem; }
.rq-steps-table li { margin: 0.1rem 0; }
@keyframes rq-pop { 0% { transform: scale(0.6); opacity: 0; } 60% { transform: scale(1.08); opacity: 1; } 100% { transform: scale(1); } }
@keyframes rq-slide { from { transform: translateY(6px); opacity: 0; } to { transform: translateY(0); opacity: 1; } }
@keyframes rq-bump { 0% { transform: scale(1); } 40% { transform: scale(1.3); } 100% { transform: scale(1); } }
@keyframes rq-dot { 0%, 80%, 100% { transform: translateY(0); opacity: 0.4; } 40% { transform: translateY(-4px); opacity: 1; } }
@media (max-width: 767.98px) {
  .rq-intake { grid-template-columns: 1fr; }
  .rq-flowpane { position: static; min-height: 0; }
  .rq-messages { max-height: 16rem; }
  .rq-pick-row { flex-direction: column; }
}

''',
    'survey.js': r'''// 自動化を検討: AI と会話しながら手順を聞き取り、流れ図がその場でできていく
(function () {
  var O = JSON.parse(document.getElementById("rq-options").textContent);
  var S = JSON.parse(document.getElementById("rq-state").textContent);
  var CFG = window.RQ_CONFIG || { api: {} };
  var toastTimer = null, timers = {}, renderSeq = 0;
  if (window.mermaid) mermaid.initialize({ startOnLoad: false, securityLevel: "loose", theme: "neutral", fontFamily: "inherit", flowchart: { htmlLabels: true, curve: "basis", nodeSpacing: 22, rankSpacing: 36, padding: 6, wrappingWidth: 1000 } });

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
    toastTimer = setTimeout(function () { t.classList.remove("is-show"); }, 2600);
  }
  function debounce(key, fn, ms) { clearTimeout(timers[key]); timers[key] = setTimeout(fn, ms || 700); }
  function pills(container, options, single, onChange) {
    container.innerHTML = "";
    options.forEach(function (o) {
      var b = el("button", "rq-pill", (O.frequency_hints && O.frequency_hints[o]) || o);
      b.type = "button";
      b.dataset.value = o;
      b.addEventListener("click", function () {
        if (single) container.querySelectorAll(".rq-pill").forEach(function (p) { if (p !== b) p.classList.remove("is-on"); });
        b.classList.toggle("is-on");
        var f = container.closest(".rq-field");
        if (f) clearError(f);
        if (onChange) onChange();
      });
      container.appendChild(b);
    });
  }
  function pillValue(c) { var on = c.querySelector(".rq-pill.is-on"); return on ? on.dataset.value : ""; }
  function pillValues(c) { return Array.prototype.map.call(c.querySelectorAll(".rq-pill.is-on"), function (p) { return p.dataset.value; }); }
  function setPill(c, v) { c.querySelectorAll(".rq-pill").forEach(function (p) { p.classList.toggle("is-on", !!v && p.dataset.value === v); }); }
  function fieldError(key, msg) {
    var f = document.querySelector('#rq-final .rq-field[data-key="' + key + '"]');
    if (!f) return;
    f.classList.add("is-invalid");
    f.querySelector(".rq-err").textContent = msg;
  }

  // ---------- 入力欄（あなたのこと） ----------
  function renderField(f, onChange) {
    var wrap = el("div", "rq-field");
    wrap.dataset.key = f.key;
    var label = el("div", "rq-label");
    label.appendChild(el("span", null, f.label));
    if (f.required) label.appendChild(el("span", "rq-req", "必須"));
    wrap.appendChild(label);
    if (f.type === "select") {
      var sel = el("select", "form-select");
      sel.name = f.key;
      var first = el("option", null, "選択");
      first.value = "";
      sel.appendChild(first);
      var opts = f.options === "work_patterns" ? O.work_patterns.map(function (w) { return { value: w, label: w }; }) : O.workplaces.map(function (w) { return { value: String(w.id), label: w.name }; });
      opts.forEach(function (o) { var op = el("option", null, o.label); op.value = o.value; sel.appendChild(op); });
      wrap.appendChild(sel);
    } else {
      var input = el("input", "form-control");
      input.type = "text";
      input.name = f.key;
      if (f.placeholder) input.placeholder = f.placeholder;
      if (f.lookup) input.autocomplete = "off";
      wrap.appendChild(input);
    }
    wrap.appendChild(el("div", "rq-err"));
    wrap.addEventListener("input", function () { clearError(wrap); if (onChange) onChange(); });
    wrap.addEventListener("change", function () { if (onChange) onChange(); });
    return wrap;
  }
  function clearError(wrap) { wrap.classList.remove("is-invalid"); var e = wrap.querySelector(".rq-err"); if (e) e.textContent = ""; }
  function getValues(container) {
    var v = {};
    container.querySelectorAll(".rq-field").forEach(function (w) {
      var inp = w.querySelector("input, select");
      v[w.dataset.key] = inp ? inp.value.trim() : "";
    });
    return v;
  }
  function showErrors(container, errors) {
    var first = null;
    Object.keys(errors || {}).forEach(function (k) {
      var w = container.querySelector('.rq-field[data-key="' + k + '"]');
      if (!w) return;
      w.classList.add("is-invalid");
      w.querySelector(".rq-err").textContent = errors[k];
      if (!first) first = w;
    });
    if (first) first.scrollIntoView({ behavior: "smooth", block: "center" });
    return !!first;
  }

  // ---------- 1. あなたのこと（社員IDは確認だけ。職場・勤務形態は選ぶ） ----------
  var personBox = $("rq-person");
  function personField(key) { return O.person.filter(function (f) { return f.key === key; })[0]; }
  function savePersonSoon() { debounce("person", function () { savePerson(false); }, 300); }
  var idRow = el("div", "rq-id-row");
  var idWrap = renderField(personField("emp_id"), null);
  idRow.appendChild(idWrap);
  idRow.appendChild(renderField(personField("workplace_id"), savePersonSoon));
  idRow.appendChild(renderField(personField("work_pattern"), savePersonSoon));
  var identity = el("div", "rq-identity");
  idRow.appendChild(identity);
  personBox.appendChild(idRow);
  var empInput = idWrap.querySelector("input");
  var found = null, idProblem = "", confirmedId = "";
  var NOT_FOUND = "この社員IDは見つかりません。IDを確かめてください";
  function normId(v) { return (v || "").normalize("NFKC").trim().toLowerCase(); }
  function setIdentity(text, ng) { identity.textContent = text || ""; identity.classList.toggle("is-ng", !!ng); }
  function fillSelects(workplaceId, pattern) {
    var ws = personBox.querySelector('select[name="workplace_id"]');
    if (ws && workplaceId) ws.value = String(workplaceId);
    var ps = personBox.querySelector('select[name="work_pattern"]');
    if (ps && pattern) ps.value = pattern;
  }
  function clearSelects() { personBox.querySelectorAll("select").forEach(function (s) { s.value = ""; }); }
  function forgetPerson() {
    confirmedId = "";
    clearSelects();
    if (!S.respondent) return;
    S.respondent = null;
    takeTasks({});
    api("POST", CFG.api.reset);
  }
  function respondentLost(j) {
    if (!j || !j.errors || j.errors._ !== "respondent") return false;
    forgetPerson();
    if (j.errors.emp_id) { found = false; idProblem = j.errors.emp_id; setIdentity(j.errors.emp_id, true); }
    toast(j.message || "「あなたのこと」を入れ直してください");
    $("sec-person").scrollIntoView({ behavior: "smooth", block: "start" });
    return true;
  }
  async function lookupEmp() {
    var id = empInput.value.trim();
    found = null;
    idProblem = "";
    if (!id) { setIdentity(""); if (S.respondent || confirmedId) forgetPerson(); return; }
    var j = await api("GET", CFG.api.lookup + encodeURIComponent(id));
    if (empInput.value.trim() !== id || !j._ok) return;
    var key = normId(j.emp_id || id);
    if (j.found && !j.excluded) {
      if ((S.respondent && normId(S.respondent.emp_id) !== key) || (confirmedId && confirmedId !== key)) forgetPerson();
      confirmedId = key;
      found = j;
      setIdentity("社員IDを確認しました");
      var v = getValues(personBox);
      if (v.workplace_id && v.work_pattern) debounce("person", function () { savePerson(false); }, 200);
    } else {
      found = false;
      idProblem = j.excluded ? (j.message || NOT_FOUND) : NOT_FOUND;
      if (S.respondent || confirmedId) forgetPerson();
      setIdentity(idProblem, true);
    }
  }
  empInput.addEventListener("input", function () { $("rq-done").classList.add("d-none"); debounce("lookup", lookupEmp, 500); });
  empInput.addEventListener("change", lookupEmp);
  var personSaving = null;
  function savePerson(showErr) {
    if (personSaving) return personSaving;
    var v = getValues(personBox);
    var errors = {};
    if (!v.emp_id) errors.emp_id = "入力してください";
    else if (found === false) errors.emp_id = idProblem || "見つかりません";
    if (!v.workplace_id) errors.workplace_id = "選んでください";
    if (!v.work_pattern) errors.work_pattern = "選んでください";
    if (Object.keys(errors).length) {
      if (showErr) showErrors(personBox, errors);
      else if (found && !errors.emp_id) showErrors(personBox, errors);
      return Promise.resolve(false);
    }
    personSaving = api("POST", CFG.api.respondent, { emp_id: v.emp_id, workplace_id: v.workplace_id, work_pattern: v.work_pattern }).then(function (j) {
      personSaving = null;
      if (!j._ok) {
        if (j.errors && j.errors.emp_id && v.emp_id) { found = false; idProblem = j.errors.emp_id; setIdentity(j.errors.emp_id, true); }
        showErrors(personBox, j.errors || {});
        return false;
      }
      if (normId(empInput.value) !== normId(j.respondent.emp_id)) { api("POST", CFG.api.reset); return false; }
      var changed = !S.respondent || S.respondent.id !== j.respondent.id;
      S.respondent = j.respondent;
      S.stance = j.stance || null;
      confirmedId = normId(j.respondent.emp_id);
      if (changed) {
        takeTasks(j);
        if (j.resumed && S.tasks.length) toast("前回の続きから登録できます");
      }
      setIdentity("社員IDを確認しました");
      return true;
    });
    return personSaving;
  }
  function empNow() { return empInput.value.trim(); }

  // ---------- 2. 検討する業務（AI の聞き取り） ----------
  var pickBox = $("rq-pick"), intakeBox = $("rq-intake"), finalBox = $("rq-final"), resultBox = $("rq-result");
  var msgBox = $("rq-messages"), chipBox = $("rq-chips"), input = $("rq-input"), flowBox = $("rq-flow"), stepsBox = $("rq-steps");
  var conv = null, busy = false;

  function takeTasks(j) {
    S.tasks = j.tasks || [];
    S.suggested = j.suggested || []; S.suggested_wants = j.suggested_wants || [];
    S.stance = j.stance || null;
    renderRegistered();
    renderSuggest();
  }
  function renderRegistered() {
    var box = $("rq-registered");
    box.innerHTML = "";
    $("rq-registered-count").textContent = S.tasks.length ? "登録済み " + S.tasks.length + " 件" : "";
    S.tasks.forEach(function (t) {
      var row = el("div", "rq-reg");
      row.appendChild(el("span", "rq-reg-name", t.name));
      row.appendChild(el("span", "rq-reg-meta", "月 " + t.hours + " h・" + (t.wish || "") + (t.stage ? "・" + t.stage : "")));
      if (t.can_delete) {
        var del = el("button", "btn btn-link btn-sm text-danger", "取り消す");
        del.type = "button";
        del.addEventListener("click", async function () {
          if (!confirm("「" + t.name + "」の登録を取り消しますか？")) return;
          var j = await api("DELETE", CFG.api.tasks + "/" + t.id, { emp_id: empNow() });
          if (j._ok) takeTasks(j);
          else if (!respondentLost(j)) toast(j.message || "取り消せませんでした");
        });
        row.appendChild(del);
      }
      box.appendChild(row);
    });
  }
  function renderSuggest() {
    renderMonthlyBanner();
    var box = $("rq-suggest");
    box.innerHTML = "";
    var names = S.suggested || [], wants = S.suggested_wants || [];
    box.classList.toggle("d-none", !names.length);
    if (!names.length) return;
    var others = names.filter(function (n) { return wants.indexOf(n) < 0; });
    [[wants, "なくしたい・自動化したいと答えた業務:"], [others, wants.length ? "アンケートで挙げたほかの業務:" : "アンケートで挙げた業務から選ぶ:"]].forEach(function (g) {
      if (!g[0].length) return;
      var row = el("div", "rq-suggest-row");
      row.appendChild(el("span", "rq-suggest-label", g[1]));
      g[0].forEach(function (n) {
        var b = el("button", "rq-pill", n);
        b.type = "button";
        b.addEventListener("click", function () { $("rq-title").value = n; startIntake(false, wants.indexOf(n) >= 0 ? "monthly" : "", n); });
        row.appendChild(b);
      });
      box.appendChild(row);
    });
  }
  function doneNames() {
    var out = [];
    (S.tasks || []).forEach(function (t) { out.push(normName(t.name)); if (t.monthly_name) out.push(normName(t.monthly_name)); });
    return out;
  }
  function nextMonthlyTask() {
    if (!CFG.monthly) return "";
    var done = doneNames();
    return (CFG.monthly.wants_tasks || []).filter(function (n) { return done.indexOf(normName(n)) < 0; })[0] || "";
  }
  function renderMonthlyBanner() {
    var box = $("rq-from-monthly");
    if (!CFG.monthly || !box) return;
    box.innerHTML = "";
    box.appendChild(el("div", "rq-from-title", CFG.monthly.period_label + "の月次アンケートの続きです"));
    box.appendChild(el("div", "small", "「なくしたい・自動化したい」と答えた業務を、1つずつ AI と整理します。あなたのことと業務名はアンケートの答えから入力済みです。"));
    var list = el("div", "rq-from-list");
    var done = doneNames();
    var current = conv ? normName(conv.monthlyName || conv.title) : "";
    CFG.monthly.wants_tasks.forEach(function (n) {
      var k = normName(n), st = done.indexOf(k) >= 0 ? "is-done" : (k === current ? "is-now" : "");
      list.appendChild(el("span", "rq-from-item " + st, (st === "is-done" ? "✓ " : st === "is-now" ? "▶ " : "") + n));
    });
    box.appendChild(list);
    box.classList.remove("d-none");
  }
  async function startFromMonthly() {
    renderMonthlyBanner();
    var next = nextMonthlyTask();
    if (!next) return;
    $("rq-title").value = next;
    if (!S.respondent) {
      fillSelects(CFG.monthly.workplace_id, CFG.monthly.work_pattern);
      await lookupEmp();
      var ok = await savePerson(false);
      if (!ok) { toast("「あなたのこと」の空欄を選ぶと始められます"); return; }
    }
    startIntake(false, "monthly", next);
  }
  function addMsg(role, text, animate) {
    var m = el("div", "rq-msg " + (role === "ai" ? "is-ai" : "is-user") + (animate ? " is-new" : ""), text);
    msgBox.appendChild(m);
    msgBox.scrollTop = msgBox.scrollHeight;
    return m;
  }
  function typing(on) {
    var t = msgBox.querySelector(".rq-typing");
    if (on && !t) {
      t = el("div", "rq-msg is-ai rq-typing");
      t.innerHTML = "<span></span><span></span><span></span>";
      msgBox.appendChild(t);
      msgBox.scrollTop = msgBox.scrollHeight;
    }
    if (!on && t) t.remove();
  }
  function renderChips(chips) {
    chipBox.innerHTML = "";
    if (!(chips || []).length) return;
    chipBox.appendChild(el("div", "rq-chips-label", "選択肢をタップするか、自由に入力してください"));
    chips.forEach(function (c) {
      var b = el("button", "rq-pill rq-chip");
      b.type = "button";
      b.appendChild(el("i", "bi bi-hand-index-thumb"));
      b.appendChild(document.createTextNode((O.frequency_hints && O.frequency_hints[c]) || c));
      b.addEventListener("click", function () { send(c); });
      chipBox.appendChild(b);
    });
  }
  var CIRCLED = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳";
  function circled(n) { return n >= 1 && n <= CIRCLED.length ? CIRCLED.charAt(n - 1) : "(" + n + ")"; }
  function sleep(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }
  function renderSteps(steps, upto) {
    stepsBox.innerHTML = "";
    steps.slice(0, upto).forEach(function (x, i) {
      stepsBox.appendChild(el("li", i >= conv.shownSteps ? "is-new" : null, x.action + (x.system ? "［" + x.system + "］" : "") + (x.to ? " → " + x.to + (x.how ? "（" + x.how + "）" : "") : "")));
    });
    var sc = $("rq-stepcount");
    sc.textContent = "手順 " + upto;
    if (upto > conv.shownSteps) { sc.classList.remove("is-bump"); void sc.offsetWidth; sc.classList.add("is-bump"); }
  }
  var iconFont = document.fonts && document.fonts.load ? document.fonts.load('16px "bootstrap-icons"').catch(function () {}) : Promise.resolve();
  async function renderDef(def) {
    if (!def) { flowBox.innerHTML = '<div class="rq-flow-empty">ここに流れができていきます</div>'; conv.lastNodes = []; return; }
    if (!window.mermaid) return;
    await iconFont;
    try {
      var out = await mermaid.render("rq-flow-svg-" + (++renderSeq), def);
      flowBox.innerHTML = out.svg;
      var svg = flowBox.querySelector("svg");
      if (svg) { svg.style.maxWidth = "100%"; svg.style.height = "auto"; svg.removeAttribute("height"); }
      var ids = [];
      flowBox.querySelectorAll("g.node").forEach(function (g) {
        var id = (g.id || "").replace(/^flowchart-/, "").replace(/-\d+$/, "");
        ids.push(id);
        if (conv.lastNodes.indexOf(id) < 0) g.classList.add("rq-node-new");
      });
      conv.lastNodes = ids;
    } catch (e) {
      flowBox.innerHTML = '<div class="rq-flow-empty">図を描けませんでした</div>';
    }
  }
  async function playFrames(frames, finalDef, state) {
    var steps = (state && state.steps) || [];
    var base = conv.shownSteps;
    for (var i = 0; i < frames.length; i++) {
      renderSteps(steps, Math.min(steps.length, base + i + 1));
      await renderDef(frames[i]);
      if (i < frames.length - 1) await sleep(550);
    }
    renderSteps(steps, steps.length);
    if (!frames.length || frames[frames.length - 1] !== finalDef) await renderDef(steps.length ? finalDef : "");
    conv.shownSteps = steps.length;
    $("rq-replay").classList.toggle("d-none", !conv.frames.length);
  }
  function addNote(text) {
    var n = el("div", "rq-note is-new", text);
    n.addEventListener("click", function () { $("rq-flowpane").scrollIntoView({ behavior: "smooth", block: "center" }); });
    msgBox.appendChild(n);
    msgBox.scrollTop = msgBox.scrollHeight;
  }
  function setFlowDone(done) {
    $("rq-confirm").classList.toggle("d-none", !done);
    $("rq-inputrow").classList.toggle("d-none", done);
    $("rq-flowpane").classList.toggle("is-done", done);
  }
  function setBusy(on) {
    busy = on;
    input.disabled = on;
    $("rq-send").disabled = on;
    $("rq-bulk-send").disabled = on;
  }
  async function turn(userText, image) {
    if (busy || !conv) return;
    setBusy(true);
    renderChips([]);
    if (userText !== null) { conv.messages.push({ role: "user", text: userText }); addMsg("user", userText); }
    typing(true);
    var j = await api("POST", CFG.api.intake, { emp_id: empNow(), title: conv.title, messages: conv.messages, state: conv.state, image: image || null, shown: conv.shownSteps, origin: conv.origin });
    typing(false);
    setBusy(false);
    if (!j._ok) {
      if (userText !== null) conv.messages.pop();
      if (!respondentLost(j)) toast(j.message || "AI と通信できませんでした。もう一度送ってください");
      return;
    }
    conv.state = j.state || conv.state;
    conv.messages.push({ role: "ai", text: j.reply });
    addMsg("ai", j.reply, true);
    renderChips(j.chips);
    if (conv.state.title) $("rq-flow-title").textContent = conv.state.title;
    var before = conv.shownSteps;
    conv.frames = conv.frames.concat(j.frames || []);
    conv.finalDef = j.diagram || "";
    await playFrames(j.frames || [], conv.finalDef, conv.state);
    if (conv.shownSteps > before) addNote("＋ 手順" + circled(before + 1) + (conv.shownSteps > before + 1 ? "〜" + circled(conv.shownSteps) : "") + " を流れ図に追加しました");
    if (j.done) addNote("✓ 流れができました。右上の「最初から見る」で育つ様子を再生できます");
    $("rq-progress").textContent = j.done ? "流れができました！" : (j.remaining > 0 ? "あと " + j.remaining + " 問くらい" : "もう少しです");
    conv.done = !!j.done;
    setFlowDone(conv.done);
    if (!j.done) input.focus();
    else $("rq-confirm").scrollIntoView({ behavior: "smooth", block: "nearest" });
  }
  function send(text) {
    text = (text || input.value).trim();
    if (!text) return;
    input.value = "";
    turn(text, null);
  }
  function shrinkImage(file) {
    return new Promise(function (resolve, reject) {
      var reader = new FileReader();
      reader.onload = function () {
        var img = new Image();
        img.onload = function () {
          var max = 1600, w = img.width, h = img.height;
          var k = Math.min(1, max / Math.max(w, h));
          var c = document.createElement("canvas");
          c.width = Math.round(w * k); c.height = Math.round(h * k);
          c.getContext("2d").drawImage(img, 0, 0, c.width, c.height);
          resolve(c.toDataURL("image/jpeg", 0.85));
        };
        img.onerror = reject;
        img.src = reader.result;
      };
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
  }
  function normName(v) { return (v || "").normalize("NFKC").replace(/\s+/g, "").toLowerCase(); }
  function ownTask(name) { var n = normName(name); return n ? (S.tasks || []).filter(function (t) { return normName(t.name) === n; })[0] : null; }
  function showDupNote(task, origin, monthlyName) {
    var box = $("rq-dup-note");
    box.innerHTML = "";
    box.appendChild(el("div", null, "同じ名前の業務「" + task.name + "」を登録済みです（月 " + task.hours + " h）。このまま始めると別の案件になります。"));
    var go = el("button", "btn btn-sm btn-outline-dark", "別の業務として続ける");
    go.type = "button";
    go.addEventListener("click", function () { box.classList.add("d-none"); startIntake(true, origin, monthlyName); });
    var stop = el("button", "btn btn-sm btn-outline-secondary", "やめる（下の一覧で取り消せます）");
    stop.type = "button";
    stop.addEventListener("click", function () { box.classList.add("d-none"); });
    box.appendChild(go);
    box.appendChild(stop);
    box.classList.remove("d-none");
  }
  async function startIntake(skipDup, origin, monthlyName) {
    if (!S.respondent) {
      var ok = await savePerson(true);
      if (!ok) { toast("先に「あなたのこと」を入力してください"); return; }
    }
    var title = $("rq-title").value.trim();
    if (!skipDup) {
      var dupTask = ownTask(title);
      if (dupTask) { showDupNote(dupTask, origin, monthlyName); return; }
    }
    $("rq-dup-note").classList.add("d-none");
    conv = { title: title, messages: [], state: { title: title, steps: [], sources: [], outputs: [], recipients: [], purpose: "", frequency: null, duration: null, shift: null, pains: [], unknowns: [] }, done: false, lastNodes: [], shownSteps: 0, frames: [], finalDef: "", allowDup: skipDup ? normName(title) : "", origin: origin || "", monthlyName: monthlyName || "" };
    renderMonthlyBanner();
    $("rq-replay").classList.add("d-none");
    msgBox.innerHTML = "";
    stepsBox.innerHTML = "";
    flowBox.innerHTML = '<div class="rq-flow-empty">ここに流れができていきます</div>';
    $("rq-flow-title").textContent = title || "流れ図";
    $("rq-stepcount").textContent = "手順 0";
    $("rq-progress").textContent = "AI が質問します";
    $("rq-bulk").classList.add("d-none");
    pickBox.classList.add("d-none");
    resultBox.classList.add("d-none");
    finalBox.classList.add("d-none");
    intakeBox.classList.remove("d-none");
    setFlowDone(false);
    intakeBox.scrollIntoView({ behavior: "smooth", block: "start" });
    await turn(null, null);
  }
  function backToPick(quiet) {
    conv = null;
    intakeBox.classList.add("d-none");
    finalBox.classList.add("d-none");
    resultBox.classList.add("d-none");
    pickBox.classList.remove("d-none");
    $("rq-title").value = "";
    renderSuggest();
    if (!quiet) pickBox.scrollIntoView({ behavior: "smooth", block: "start" });
  }
  function monthlySource(title) {
    var k = normName(title);
    if (!k) return "";
    return ((CFG.monthly && CFG.monthly.wants_tasks) || []).concat(S.suggested || []).filter(function (n) { return normName(n) === k; })[0] || "";
  }
  $("rq-start").addEventListener("click", function () { startIntake(false, "", monthlySource($("rq-title").value)); });
  $("rq-title").addEventListener("keydown", function (e) { if (e.key === "Enter") { e.preventDefault(); startIntake(false, "", monthlySource($("rq-title").value)); } });
  $("rq-send").addEventListener("click", function () { send(); });
  input.addEventListener("keydown", function (e) { if (e.key === "Enter" && !e.isComposing) { e.preventDefault(); send(); } });
  $("rq-restart").addEventListener("click", function () { if (!conv || confirm("聞き取りを最初からやり直しますか？")) backToPick(); });
  $("rq-bulk-toggle").addEventListener("click", function () { $("rq-bulk").classList.toggle("d-none"); if (!$("rq-bulk").classList.contains("d-none")) $("rq-bulk-text").focus(); });
  $("rq-bulk-send").addEventListener("click", function () {
    var text = $("rq-bulk-text").value.trim();
    if (!text) { toast("手順を書いてください"); return; }
    $("rq-bulk-text").value = "";
    $("rq-bulk").classList.add("d-none");
    turn(text, null);
  });
  $("rq-photo").addEventListener("change", async function () {
    var file = this.files && this.files[0];
    this.value = "";
    if (!file || !conv) return;
    try {
      var data = await shrinkImage(file);
      turn("手順書（写真）を送りました。読み取ってください", data);
    } catch (e) {
      toast("写真を読み込めませんでした");
    }
  });
  $("rq-confirm-no").addEventListener("click", function () {
    conv.done = false;
    setFlowDone(false);
    $("rq-progress").textContent = "直したいところを教えてください";
    input.placeholder = "例：2番は違う。先に課長に確認する";
    input.focus();
  });
  $("rq-confirm-yes").addEventListener("click", function () {
    setPill($("rq-frequency"), conv.state.frequency);
    setPill($("rq-duration"), conv.state.duration);
    setPill($("rq-wish"), "");
    var shiftWork = !!(S.respondent && S.respondent.work_pattern === "交代勤務");
    $("rq-shift-field").classList.toggle("d-none", !shiftWork);
    setPill($("rq-shift"), conv.state.shift);
    var st = S.stance;
    var preset = !!(st && st.self_improve);
    $("rq-self").checked = preset;
    $("rq-skills").classList.toggle("d-none", !preset);
    fillSkills();
    $("rq-self-note").textContent = preset ? "月次アンケート（" + st.period_label + "）で「" + st.value + "」と答えたので、チェックを入れています（外せます）" : "";
    $("rq-self-note").classList.toggle("d-none", !preset);
    finalBox.classList.remove("d-none");
    finalBox.scrollIntoView({ behavior: "smooth", block: "start" });
  });
  pills($("rq-frequency"), O.frequencies, true);
  pills($("rq-duration"), O.durations, true);
  pills($("rq-wish"), O.wishes, true);
  pills($("rq-shift"), O.shift_options, true);
  pills($("rq-skills-have"), O.skills, false);
  pills($("rq-skills-learn"), O.skills, false);
  [$("rq-skills-have"), $("rq-skills-learn")].forEach(function (box) {
    box.addEventListener("click", function (e) {
      var b = e.target.closest(".rq-pill");
      if (!b || !b.classList.contains("is-on")) return;
      box.querySelectorAll(".rq-pill.is-on").forEach(function (p) {
        if (p !== b && (b.dataset.value === O.skill_none || p.dataset.value === O.skill_none)) p.classList.remove("is-on");
      });
    });
  });
  function fillSkills() {
    var on = function (box, vals) { box.querySelectorAll(".rq-pill").forEach(function (p) { p.classList.toggle("is-on", (vals || []).indexOf(p.dataset.value) >= 0); }); };
    on($("rq-skills-have"), S.respondent && S.respondent.c2_skills);
    on($("rq-skills-learn"), S.respondent && S.respondent.c3_learn);
  }
  $("rq-self-label").textContent = O.self_improve_yes;
  $("rq-self").addEventListener("change", function () { $("rq-skills").classList.toggle("d-none", !this.checked); });
  $("rq-save-task").addEventListener("click", async function () {
    if (!conv) return;
    var btn = this;
    var body = {
      name: conv.state.title || conv.title, state: conv.state, messages: conv.messages,
      frequency: pillValue($("rq-frequency")), duration: pillValue($("rq-duration")), wish: pillValue($("rq-wish")),
      force: !!(conv.allowDup && normName(conv.state.title || conv.title) === conv.allowDup),
      self_improve: $("rq-self").checked, skills: pillValues($("rq-skills-have")), learn: pillValues($("rq-skills-learn")),
      shift: pillValue($("rq-shift")), monthly_name: conv.monthlyName || ""
    };
    var missing = false;
    if (S.respondent && S.respondent.work_pattern === "交代勤務" && !body.shift) { fieldError("shift", "選んでください"); missing = true; }
    if (!body.frequency) { fieldError("frequency", "選んでください"); missing = true; }
    if (!body.duration) { fieldError("duration", "選んでください"); missing = true; }
    if (!body.wish) { fieldError("wish", "選んでください"); missing = true; }
    if (missing) return;
    await registerTask(body, btn);
  });
  async function registerTask(body, btn) {
    btn.disabled = true;
    btn.textContent = "登録中…";
    var j = await api("POST", CFG.api.tasks, Object.assign({}, body, { emp_id: empNow() }));
    btn.disabled = false;
    btn.innerHTML = '<i class="bi bi-plus-lg"></i> この業務を登録';
    $("rq-dup").classList.add("d-none");
    if (!j._ok && j.duplicate) {
      var box = $("rq-dup");
      box.innerHTML = "";
      box.appendChild(el("div", null, "同じ名前の業務「" + j.duplicate.name + "」を " + j.duplicate.created_at + " に登録済みです（月 " + j.duplicate.hours + " h）。どうしますか？" + (j.duplicate.can_delete ? "" : "（前の登録は改善チームが対応を始めているため取り消せません）")));
      var replace = el("button", "btn btn-sm btn-primary", "前の登録を取り消して、この内容で登録");
      replace.type = "button";
      replace.addEventListener("click", async function () {
        var d = await api("DELETE", CFG.api.tasks + "/" + j.duplicate.id, { emp_id: empNow() });
        if (!d._ok) { if (!respondentLost(d)) toast(d.message || "取り消せませんでした"); return; }
        S.tasks = d.tasks || S.tasks;
        renderRegistered();
        registerTask(Object.assign({}, body, { force: true }), btn);
      });
      var both = el("button", "btn btn-sm btn-outline-dark", "別の業務として登録");
      both.type = "button";
      both.addEventListener("click", function () { registerTask(Object.assign({}, body, { force: true }), btn); });
      var cancel = el("button", "btn btn-sm btn-outline-secondary", "やめる");
      cancel.type = "button";
      cancel.addEventListener("click", function () { box.classList.add("d-none"); });
      if (j.duplicate.can_delete) box.appendChild(replace);
      box.appendChild(both);
      box.appendChild(cancel);
      box.classList.remove("d-none");
      box.scrollIntoView({ behavior: "smooth", block: "center" });
      return;
    }
    if (!j._ok) {
      if (respondentLost(j)) return;
      Object.keys(j.errors || {}).forEach(function (k) { fieldError(k, j.errors[k]); });
      toast(j.message || "登録できませんでした");
      return;
    }
    if (j.respondent) S.respondent = j.respondent;
    S.stance = j.stance || null;
    S.tasks = j.tasks || S.tasks;
    S.suggested = j.suggested || []; S.suggested_wants = j.suggested_wants || [];
    renderRegistered();
    intakeBox.classList.add("d-none");
    finalBox.classList.add("d-none");
    $("rq-result-title").textContent = "「" + j.task.name + "」を登録しました（月 約 " + j.task.hours + " 時間）";
    $("rq-impression").textContent = j.first_impression || "";
    resultBox.classList.remove("d-none");
    resultBox.scrollIntoView({ behavior: "smooth", block: "center" });
    conv = null;
    renderMonthlyBanner();
    var nextTask = nextMonthlyTask(), nb = $("rq-next-monthly");
    nb.classList.toggle("d-none", !nextTask);
    nb.innerHTML = "";
    if (nextTask) {
      nb.appendChild(el("i", "bi bi-arrow-right-circle"));
      nb.appendChild(document.createTextNode(" 次の業務「" + nextTask + "」を検討する"));
      nb.dataset.task = nextTask;
    }
  }
  $("rq-another").addEventListener("click", function () { backToPick(); });
  $("rq-next-monthly").addEventListener("click", function () { $("rq-title").value = this.dataset.task; startIntake(false, "monthly", this.dataset.task); });
  $("rq-replay").addEventListener("click", async function () {
    if (!conv || busy || !conv.frames.length) return;
    var btn = this;
    btn.disabled = true;
    conv.shownSteps = 0;
    conv.lastNodes = [];
    await renderDef("");
    await sleep(300);
    await playFrames(conv.frames.slice(), conv.finalDef, conv.state);
    btn.disabled = false;
  });

  // ---------- 送信して完了 ----------
  $("rq-finish").addEventListener("click", async function () {
    if (!S.respondent) {
      var ok = await savePerson(true);
      if (!ok) { toast("「あなたのこと」を入力してください"); return; }
    }
    if (!S.tasks.length) {
      toast("業務を1件以上登録してください");
      $("sec-task").scrollIntoView({ behavior: "smooth", block: "start" });
      return;
    }
    var total = S.tasks.reduce(function (a, t) { return a + (t.monthly_minutes || 0); }, 0);
    var sum = "ありがとうございました。登録した業務 " + S.tasks.length + " 件（月 約 " + (total / 60).toFixed(1) + " 時間）は改善案件として確認し、活動状況に載せます。次の人は上の「あなたのこと」から始められます。";
    await api("POST", CFG.api.reset);
    resetPage();
    $("rq-done-sum").textContent = sum;
    var done = $("rq-done");
    done.classList.remove("d-none");
    done.scrollIntoView({ behavior: "smooth", block: "center" });
    toast("送信しました");
    if (window.uiCelebrate) window.uiCelebrate();
  });
  function resetPage() {
    S.respondent = null;
    takeTasks({});
    found = null; idProblem = ""; confirmedId = "";
    empInput.value = "";
    clearSelects();
    personBox.querySelectorAll(".rq-field").forEach(clearError);
    setIdentity("");
    CFG.monthly = null; CFG.emp = ""; CFG.prefill = "";
    var from = $("rq-from-monthly");
    from.innerHTML = "";
    from.classList.add("d-none");
    $("rq-dup-note").classList.add("d-none");
    $("rq-dup").classList.add("d-none");
    backToPick(true);
    if (window.history && history.replaceState) history.replaceState(null, "", location.pathname);
  }

  // ---------- 初期表示 ----------
  renderRegistered();
  renderSuggest();
  if (CFG.prefill) $("rq-title").value = CFG.prefill;
  if (S.respondent) {
    empInput.value = S.respondent.emp_id || "";
    confirmedId = normId(S.respondent.emp_id);
    found = { workplace_id: S.respondent.workplace_id };
    fillSelects(S.respondent.workplace_id, S.respondent.work_pattern);
    setIdentity("社員IDを確認しました");
  } else if (CFG.emp) {
    empInput.value = CFG.emp;
    if (!CFG.monthly) lookupEmp();
  }
  if (CFG.monthly) startFromMonthly();
})();
''',
}


# ===== アプリ生成 =====
def _persistent_secret(instance_path):
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

    public_blueprints = {"board", "survey", "auth"}
    public_endpoints = {"static", "main.index", "monthly.form", "monthly.submit", "monthly.done", "monthly.api_employee"}

    @app.before_request
    def require_login():
        if request.endpoint is None:
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
        except Exception:
            return "<h1>エラーが発生しました</h1><p>しばらくしてからもう一度お試しください。</p>", 500

    @app.context_processor
    def inject_globals():
        return {
            "settings": Setting.as_dict(),
            "C": C,
            "kind_icon": kind_icon_html,
            "current_user": session.get("user"),
            "ai_enabled": ai_enabled(),
            "demo_active": demo_loaded(),
        }

    @app.template_filter("jpdate")
    def jpdate(value, fmt="%Y/%m/%d"):
        if not value:
            return ""
        return value.strftime(fmt)

    @app.template_filter("hours")
    def hours(minutes):
        try:
            return "{:,.1f}".format((minutes or 0) / 60.0)
        except (TypeError, ValueError):
            return "0.0"

    with app.app_context():
        db.create_all()
        ensure_schema()
        init_id_seq()
        seed_masters()

    return app

# ===== 起動 =====
if __name__ == "__main__":
    port = int(conf.get("PORT", 5000) or 5000)
    if len(sys.argv) > 1 and sys.argv[1] == "serve":
        from waitress import serve

        threads = int(conf.get("THREADS", 8) or 8)
        print(f"PC定型業務ゼロ化プロジェクト: http://0.0.0.0:{port}  (waitress, threads={threads})")
        serve(create_app(), host="0.0.0.0", port=port, threads=threads)
    else:
        create_app().run(host="0.0.0.0", port=port, debug=bool(conf.get("DEBUG", False)))
