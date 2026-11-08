"""PC定型業務ゼロ化プロジェクト 設定ファイル

設定はすべてこのファイルに書く（環境変数や .env は使わない）。変えたらアプリを再起動する。
空文字のままの項目は「その機能を使わない」という意味になる。
"""

# ---- 起動 ----
PORT = 5000              # 待ち受けポート
THREADS = 8              # serve.py（waitress）の同時処理数
DEBUG = False            # run.py（開発用サーバー）でデバッグ表示をするか
SECRET_KEY = ""          # セッション用の秘密鍵。空なら instance/secret_key.txt を自動生成して使う

# ---- ログイン（推進メンバー用の管理者アカウント。アンケート回答と活動状況はログイン不要） ----
ADMIN_ID = "admin"
ADMIN_PASSWORD = "adminpass"

# ---- 社員ID照会: 社内 LDAP API サービス（HTTP/JSON） ----
# 例: "http://ldap-api.local/employees/{id}"  {id} は社員IDに置き換わる。{id} がなければ ?LDAP_API_ID_PARAM=社員ID を付けて GET する
LDAP_API_URL = ""
LDAP_API_TOKEN = ""                    # Bearer トークンが必要なときだけ
LDAP_API_ID_PARAM = "employee_id"
# 応答 JSON の項目名（入れ子は "data.name" のように書く）
LDAP_API_FIELD_ID = "employee_id"
LDAP_API_FIELD_NAME = "display_name"   # 氏名。この項目になければ display_name / displayName / name / cn を順に見る
LDAP_API_FIELD_DEPT = "department"     # 所属（返らなければ空のまま。職場は本人が初回に選ぶ）
LDAP_API_FIELD_MAIL = "mail"
LDAP_API_FIELD_EXISTS = "exists"       # この項目が false なら「いない」。HTTP 404 も「いない」。氏名がなくても「いる」と判定する
LDAP_API_TIMEOUT = 5                   # 秒

# ---- 社員ID照会: LDAP サーバーへの直接問い合わせ（pip install ldap3 が必要。API があれば不要） ----
LDAP_SERVER = ""
LDAP_BIND_DN = ""
LDAP_BIND_PASSWORD = ""
LDAP_BASE_DN = ""
LDAP_ID_ATTR = "employeeID"
LDAP_NAME_ATTR = "displayName"
LDAP_DEPT_ATTR = "department"
LDAP_MAIL_ATTR = "mail"

# ---- AI連携（Claude API。空ならルールベースで動く） ----
ANTHROPIC_API_KEY = ""
ANTHROPIC_MODEL = "claude-opus-5-5"
