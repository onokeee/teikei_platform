"""社員ID照会（社内 LDAP API サービス / LDAP サーバー）。このファイルは config.py と標準ライブラリだけに依存する

照会の順序:
  1. 社内 LDAP API サービス（HTTP/JSON）   … config.py の LDAP_API_URL が設定されているとき
  2. LDAP サーバーへの直接問い合わせ（ldap3） … config.py の LDAP_SERVER が設定され、ldap3 が入っているとき

API が「そのIDが存在するか」しか返さない場合（氏名・所属なし）でも使える。氏名は LDAP_API_FIELD_NAME の項目、
なければ display_name / displayName / name / cn を順に見る。見つかった人の控え（employees テーブル）や
足りない氏名・職場の補いはアプリ側（app.py の lookup）で行う。

使い方:
    import ldap_auth
    info = ldap_auth.lookup_directory("E001")   # {"emp_id", "name", "department", "email"} か None
    ldap_auth.status_text()                    # 管理画面に出す接続状態
"""
import json
import logging
import urllib.error
import urllib.parse
import urllib.request

import config

log = logging.getLogger(__name__)


def _cfg(name, default=None):
    """config.py の値（なければ default）"""
    value = getattr(config, name, default)
    return default if value is None else value


NAME_FIELDS = ("display_name", "displayName", "name", "cn", "full_name", "fullName")


def api_config():
    return {
        "url": str(_cfg("LDAP_API_URL", "") or "").strip(),
        "token": str(_cfg("LDAP_API_TOKEN", "") or "").strip(),
        "id_param": str(_cfg("LDAP_API_ID_PARAM", "employee_id") or "employee_id"),
        "field_id": str(_cfg("LDAP_API_FIELD_ID", "employee_id") or "employee_id"),
        "field_name": str(_cfg("LDAP_API_FIELD_NAME", "name") or "name"),
        "field_dept": str(_cfg("LDAP_API_FIELD_DEPT", "department") or "department"),
        "field_mail": str(_cfg("LDAP_API_FIELD_MAIL", "mail") or "mail"),
        "field_exists": str(_cfg("LDAP_API_FIELD_EXISTS", "exists") or "exists"),
        "timeout": float(_cfg("LDAP_API_TIMEOUT", 5) or 5),
    }


def ldap_config():
    return {
        "server": str(_cfg("LDAP_SERVER", "") or "").strip(),
        "bind_dn": str(_cfg("LDAP_BIND_DN", "") or ""),
        "password": str(_cfg("LDAP_BIND_PASSWORD", "") or ""),
        "base_dn": str(_cfg("LDAP_BASE_DN", "") or ""),
        "id_attr": str(_cfg("LDAP_ID_ATTR", "employeeID") or "employeeID"),
        "name_attr": str(_cfg("LDAP_NAME_ATTR", "displayName") or "displayName"),
        "dept_attr": str(_cfg("LDAP_DEPT_ATTR", "department") or "department"),
        "mail_attr": str(_cfg("LDAP_MAIL_ATTR", "mail") or "mail"),
    }


def _dig(data, path):
    """'a.b.c' のようなキー指定で辞書をたどる"""
    cur = data
    for key in str(path).split("."):
        if isinstance(cur, dict) and key in cur:
            cur = cur[key]
        else:
            return None
    return cur


def _lookup_api(emp_id):
    cfg = api_config()
    url = cfg["url"]
    if "{id}" in url:
        url = url.replace("{id}", urllib.parse.quote(emp_id))
    else:
        sep = "&" if "?" in url else "?"
        url = f"{url}{sep}{urllib.parse.urlencode({cfg['id_param']: emp_id})}"
    req = urllib.request.Request(url, headers={"Accept": "application/json"})
    if cfg["token"]:
        req.add_header("Authorization", f"Bearer {cfg['token']}")
    try:
        with urllib.request.urlopen(req, timeout=cfg["timeout"]) as resp:
            data = json.loads(resp.read().decode("utf-8") or "null")
    except urllib.error.HTTPError as e:
        if e.code == 404:  # いない
            return None
        log.warning("LDAP API の照会に失敗しました: %s", e)
        return None
    except (urllib.error.URLError, ValueError, OSError) as e:
        log.warning("LDAP API の照会に失敗しました: %s", e)
        return None
    exists = None
    if isinstance(data, dict):
        exists = _dig(data, cfg["field_exists"])
        if exists is False or (isinstance(exists, str) and exists.strip().lower() in ("false", "0", "no")):
            return None
        for key in ("data", "result", "employee", "item"):
            if isinstance(data.get(key), (dict, list)):
                data = data[key]
                break
    if isinstance(data, list):
        data = data[0] if data else None
    if data is True:  # true だけ返す API
        data, exists = {}, True
    if not isinstance(data, dict) or (not data and exists is not True):  # null・{}・[] は「いない」
        return None
    name = _dig(data, cfg["field_name"])
    for key in NAME_FIELDS:
        if name:
            break
        name = _dig(data, key)
    return {  # 氏名・所属が返らなくても「存在する」ことは分かる
        "emp_id": str(_dig(data, cfg["field_id"]) or emp_id),
        "name": str(name or ""),
        "department": str(_dig(data, cfg["field_dept"]) or ""),
        "email": str(_dig(data, cfg["field_mail"]) or ""),
    }


def _lookup_ldap3(emp_id):
    cfg = ldap_config()
    try:
        import ldap3
    except ImportError:
        log.warning("ldap3 が未インストールのため LDAP 直接照会をスキップします")
        return None
    try:
        server = ldap3.Server(cfg["server"], get_info=ldap3.NONE, connect_timeout=5)
        conn = ldap3.Connection(server, user=cfg["bind_dn"] or None, password=cfg["password"] or None, auto_bind=True, receive_timeout=5)
        safe_id = ldap3.utils.conv.escape_filter_chars(emp_id)
        conn.search(cfg["base_dn"], f"(&(objectClass=person)({cfg['id_attr']}={safe_id}))", attributes=[cfg["name_attr"], cfg["dept_attr"], cfg["mail_attr"]])
        if not conn.entries:
            return None
        entry = conn.entries[0]

        def attr(name):
            try:
                value = entry[name].value
            except (KeyError, ldap3.core.exceptions.LDAPKeyError):
                return ""
            if isinstance(value, list):
                value = value[0] if value else ""
            return str(value or "")

        name = attr(cfg["name_attr"])
        if not name:
            return None
        return {"emp_id": emp_id, "name": name, "department": attr(cfg["dept_attr"]), "email": attr(cfg["mail_attr"])}
    except Exception as e:  # ldap3 の例外は種類が多いためまとめて扱う
        log.warning("LDAP 直接照会に失敗しました: %s", e)
        return None


def status_text():
    if api_config()["url"]:
        return f"社内 LDAP API 連携 有効（{api_config()['url']}）。照会結果は社員マスタにも保存します"
    if ldap_config()["server"]:
        try:
            import ldap3  # noqa: F401

            return f"LDAP サーバー直接照会 有効（{ldap_config()['server']}）"
        except ImportError:
            return "LDAP_SERVER が設定されていますが ldap3 が未インストールです（pip install ldap3）"
    return "LDAP 未設定（社員マスタのみで照会。管理 → 社員マスタ から CSV 取込できます）"


def lookup_directory(emp_id):
    """社員IDを API / LDAP に問い合わせる。見つからなければ None（データベースは触らない）"""
    emp_id = (emp_id or "").strip()
    if not emp_id:
        return None
    info = None
    if api_config()["url"]:
        info = _lookup_api(emp_id)
    if info is None and ldap_config()["server"]:
        info = _lookup_ldap3(emp_id)
    return info
