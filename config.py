"""設定。すべてここに書く（環境変数や .env は使わない）。変えたらアプリを再起動する。空文字は「その機能を使わない」"""

# ---- 起動（python app.py / python app.py serve のとき。flask run / waitress-serve では --port で指定） ----
PORT = 5000
THREADS = 8
DEBUG = False
SECRET_KEY = ""  # 空なら instance/secret_key.txt を自動生成

# ---- 管理者ログイン ----
ADMIN_ID = "admin"
ADMIN_PASSWORD = "adminpass"

# ---- ユーザーID照会: 社内 LDAP API（HTTP/JSON）。例 "http://ldap-api.local/employees/{id}"。{id} がなければ ?LDAP_API_ID_PARAM=ユーザーID を付けて GET ----
LDAP_API_URL = ""
LDAP_API_TOKEN = ""
LDAP_API_ID_PARAM = "employee_id"
LDAP_API_FIELD_ID = "employee_id"
LDAP_API_FIELD_NAME = "display_name"
LDAP_API_FIELD_DEPT = "department"
LDAP_API_FIELD_MAIL = "mail"
LDAP_API_FIELD_EXISTS = "exists"
LDAP_API_TIMEOUT = 5

# ---- ユーザーID照会: LDAP サーバーへ直接（pip install ldap3 が必要。API があれば不要） ----
LDAP_SERVER = ""
LDAP_BIND_DN = ""
LDAP_BIND_PASSWORD = ""
LDAP_BASE_DN = ""
LDAP_ID_ATTR = "employeeID"
LDAP_NAME_ATTR = "displayName"
LDAP_DEPT_ATTR = "department"
LDAP_MAIL_ATTR = "mail"

# ---- AI連携: 社内 Azure OpenAI の ChatGPT API（Responses API）。URL・API キー・モデル名は管理画面で設定する。temperature / top_p は送らない ----
OPENAI_REASONING_EFFORT = "low"  # "" なら送らない
OPENAI_MAX_OUTPUT_TOKENS = 4000
OPENAI_API_TIMEOUT = 90
