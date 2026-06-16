import telebot
from telebot import types
import json
import os
import re
import time
import threading
import datetime
import requests
import phonenumbers
import random
import csv
import io
import tempfile
import openpyxl
import xlrd
from bs4 import BeautifulSoup
from phonenumbers import region_code_for_number, geocoder

_PID_FILE = "/tmp/ar_otp_bot.pid"
_my_pid = os.getpid()
if os.path.exists(_PID_FILE):
    try:
        _old_pid = int(open(_PID_FILE).read().strip())
        if _old_pid != _my_pid:
            try:
                os.kill(_old_pid, 9)
                time.sleep(5)
                print(f"[START] Killed old instance PID {_old_pid}")
            except ProcessLookupError:
                pass
    except Exception:
        pass
open(_PID_FILE, "w").write(str(_my_pid))

API_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "").strip()
# <<SYNC:SUPER_ADMIN_IDS:START>>
SUPER_ADMIN_IDS = [
    6664150885,
    8523774444,
]
# <<SYNC:SUPER_ADMIN_IDS:END>>
ADMIN_IDS = list(SUPER_ADMIN_IDS)
CHANNEL_2 = ""

# ── Panel 1 (Mahofuza) ───────────────────────────────────────────────────────
P1_BASE_URL = "http://91.232.105.47/ints"
P1_LOGIN_PAGE = P1_BASE_URL + "/login"
P1_SIGNIN_URL = P1_BASE_URL + "/signin"
P1_CDR_PAGE = P1_BASE_URL + "/agent/SMSCDRStats"
P1_CDR_DATA_URL = P1_BASE_URL + "/agent/res/data_smscdr.php"
P1_USER_NAME = "Mahofuza"
P1_PASSWORD = "Mahofuza"

# ── Panel 2 (Sagardas50 / XISORA) ────────────────────────────────────────────
P2_BASE_URL = "http://94.23.31.29/sms"
P2_SIGNIN_URL = P2_BASE_URL + "/signmein"
P2_REPORTS_PAGE = P2_BASE_URL + "/client/Reports"
P2_DATA_URL = P2_BASE_URL + "/client/ajax/dt_reports.php"
P2_USER_NAME = "Sagardas50"
P2_PASSWORD = "Sagardas50"

# ── Panel 3 (Rabbi1_FD) ───────────────────────────────────────────────────────
P3_BASE_URL = "http://168.119.13.175/ints"
P3_LOGIN_PAGE = P3_BASE_URL + "/login"
P3_SIGNIN_URL = P3_BASE_URL + "/signin"
P3_CDR_PAGE = P3_BASE_URL + "/agent/SMSCDRStats"
P3_CDR_DATA_URL = P3_BASE_URL + "/agent/res/data_smscdr.php"
P3_USER_NAME = "Rabbi1_FD"
P3_PASSWORD = "Rabbi12"

# ── Panel 4 (Rabbi12) ─────────────────────────────────────────────────────────
P4_BASE_URL = "http://144.217.71.192/ints"
P4_LOGIN_PAGE = P4_BASE_URL + "/login"
P4_SIGNIN_URL = P4_BASE_URL + "/signin"
P4_CDR_PAGE = P4_BASE_URL + "/agent/SMSCDRStats"
P4_CDR_DATA_URL = P4_BASE_URL + "/agent/res/data_smscdr.php"
P4_USER_NAME = "Rabbi12"
P4_PASSWORD = "Rabbi12"

# ── Panel 5 (Rabbi12_v2 / 51.75.144.178) ─────────────────────────────────────
P5_BASE_URL = "http://51.75.144.178/ints"
P5_LOGIN_PAGE = P5_BASE_URL + "/login"
P5_SIGNIN_URL = P5_BASE_URL + "/signin"
P5_CDR_PAGE = P5_BASE_URL + "/agent/SMSCDRStats"
P5_CDR_DATA_URL = P5_BASE_URL + "/agent/res/data_smscdr.php"
P5_USER_NAME = "Rabbi12"
P5_PASSWORD = "Rabbi12@"

# ── Panel 6 (Sagardas50 / TrueSMS.net — SMSRanges) ───────────────────────────
P6_BASE_URL = "https://truesms.net"
P6_LOGIN_PAGE = P6_BASE_URL + "/login"
P6_SIGNIN_URL = P6_BASE_URL + "/signin"
P6_CDR_PAGE = P6_BASE_URL + "/agent/SMSRanges"
P6_CDR_DATA_URL = P6_BASE_URL + "/agent/res/data_smsranges.php"
P6_USER_NAME = "Sagardas50"
P6_PASSWORD = "Sagardas50"

# ── Builtin extra panels (hardcoded, always started automatically) ─────────────
# <<SYNC:_BUILTIN_PANELS:START>>
_BUILTIN_PANELS = [
    {'id': 'bp1', 'host': '139.99.69.196', 'base_url': 'http://139.99.69.196/ints', 'url_hint': 'http://139.99.69.196/ints/agent/SMSCDRStats', 'username': 'Mahofuza12', 'password': 'Mahofuza12', 'engine': 'ints_smscdr', 'data_path': '/agent/res/data_smscdr.php', 'admin_id': None},
    {'id': 'bp2', 'host': '139.99.9.4', 'base_url': 'http://139.99.9.4/ints', 'url_hint': 'http://139.99.9.4/ints/agent/SMSCDRStats', 'username': 'Rabbi12', 'password': 'Rabbi12', 'engine': 'ints_smscdr', 'data_path': '/agent/res/data_smscdr.php', 'admin_id': None},
    {'id': 'bp3', 'host': '54.36.173.235', 'base_url': 'http://54.36.173.235/ints', 'url_hint': 'http://54.36.173.235/ints/agent/SMSCDRStats', 'username': 'Rabbi12', 'password': 'Rabbi@', 'engine': 'ints_smscdr', 'data_path': '/agent/res/data_smscdr.php', 'admin_id': None},
    {'id': 'bp4', 'host': '54.39.104.241', 'base_url': 'http://54.39.104.241/ints', 'url_hint': 'http://54.39.104.241/ints/agent/SMSCDRStats', 'username': 'Rabbi5', 'password': 'Rabbi5', 'engine': 'ints_smsranges', 'data_path': '/agent/res/data_smsranges.php', 'admin_id': None},
    {'id': 'bp5', 'host': '213.32.24.208', 'base_url': 'http://213.32.24.208/ints', 'url_hint': 'http://213.32.24.208/ints/agent/SMSCDRStats', 'username': 'mahofuza', 'password': 'mahofuza@', 'engine': 'ints_smscdr', 'data_path': '/agent/res/data_smscdr.php', 'admin_id': None},
    {'id': 'bp6', 'host': '15.235.182.3', 'base_url': 'http://15.235.182.3/konekta', 'url_hint': 'http://15.235.182.3/konekta/agent/SMSCDRReports', 'username': 'Rabbi200', 'password': 'Rabbi200', 'engine': 'ints_smscdr', 'data_path': '/agent/res/data_smscdr.php', 'admin_id': None},
    {'id': 'bp7', 'host': 'nexor-iprn.com', 'base_url': 'https://nexor-iprn.com', 'url_hint': 'https://nexor-iprn.com/agent/SMSCDRStats', 'username': 'Rabbi12', 'password': 'Rabbi12@', 'engine': 'ints_smscdr', 'data_path': '/agent/res/data_smscdr.php', 'admin_id': None},
    {'id': 'bp8', 'host': '51.77.52.79', 'base_url': 'http://51.77.52.79/ints', 'url_hint': 'http://51.77.52.79/ints/agent/SMSCDRStats', 'username': 'Rabbi12', 'password': 'Rabbi12', 'engine': 'ints_smscdr', 'data_path': '/agent/res/data_smscdr.php', 'admin_id': None},
    {'id': 'bp9', 'host': '51.210.208.26', 'base_url': 'http://51.210.208.26/ints', 'url_hint': 'http://51.210.208.26/ints/agent/SMSCDRStats', 'username': 'Dasbabu50_FD', 'password': 'Dasbabu50_FD', 'engine': 'ints_smscdr', 'data_path': '/agent/res/data_smscdr.php', 'admin_id': None},
    {'id': 'bp10', 'host': 'ivasms.com', 'base_url': 'https://ivasms.com', 'url_hint': 'https://ivasms.com/portal/sms/received', 'username': 'mdrashub2@gmail.com', 'password': 'Rabbi+nnn', 'engine': 'iva_sms', 'data_path': '/portal/sms/received', 'admin_id': None},
    {'id': 'bp11', 'host': '139.99.68.231', 'base_url': 'http://139.99.68.231/ints', 'url_hint': 'http://139.99.68.231/ints/agent/SMSCDRStats', 'username': 'Rabbi12', 'password': 'Rabbi12', 'engine': 'ints_smscdr', 'data_path': '/agent/res/data_smscdr.php', 'admin_id': None},
    {'id': 'bp12', 'host': '51.75.144.178', 'base_url': 'http://51.75.144.178/ints', 'url_hint': 'http://51.75.144.178/ints/agent/SMSCDRStats', 'username': 'Rabbi12', 'password': 'Rabbi12@', 'engine': 'ints_smscdr', 'data_path': '/agent/res/data_smscdr.php', 'admin_id': None},
    {'id': 'd20591', 'host': '54.39.104.241', 'base_url': 'http://54.39.104.241/ints', 'url_hint': 'http://54.39.104.241/ints/client/SMSCDRStats', 'username': 'Atik9898', 'password': 'Atik9898', 'engine': 'ints_smscdr', 'data_path': '/agent/res/data_smscdr.php', 'admin_id': 6664150885},
    {'id': 'd20984', 'host': '54.39.104.241', 'base_url': 'http://54.39.104.241/ints', 'url_hint': 'http://54.39.104.241/ints/agent/SMSCDRStats', 'username': 'Rabbi5', 'password': 'Rabbi5', 'engine': 'ints_smscdr', 'data_path': '/agent/res/data_smscdr.php', 'admin_id': 6664150885},
    {'id': 'pscall1', 'host': 'pscall.net', 'base_url': 'http://pscall.net/restapi', 'url_hint': 'http://pscall.net/restapi/smsreport', 'username': 'api:pscall.net', 'password': '', 'api_key': 'SFNSQz1SS2NygIF6QlBR', 'api_key_param': 'key', 'engine': 'api_key', 'data_path': '/smsreport', 'admin_id': None},
    {'id': 'd34527', 'host': '151.80.19.204', 'base_url': 'http://151.80.19.204/ints', 'url_hint': 'http://151.80.19.204/ints/agent/SMSCDRStats', 'username': 'Atik9898', 'password': 'Atik9898', 'engine': 'ints_smscdr', 'data_path': '/agent/res/data_smscdr.php', 'admin_id': 6664150885},
    {'id': 'd39020', 'host': '139.99.69.196', 'base_url': 'http://139.99.69.196/ints', 'url_hint': 'http://139.99.69.196/ints/agent/SMSCDRStats', 'username': 'Mahofuza12', 'password': 'Mahofuza12', 'engine': 'ints_smscdr', 'data_path': '/agent/res/data_smscdr.php', 'admin_id': 6664150885},
    {'id': 'd42653', 'host': '54.39.104.241', 'base_url': 'http://54.39.104.241/ints', 'url_hint': 'http://54.39.104.241/ints/agent/SMSCDRStats', 'username': 'Rabbi5', 'password': 'Rabbi5', 'engine': 'ints_smscdr', 'data_path': '/agent/res/data_smscdr.php', 'admin_id': 6664150885},
    {'id': 'd6180', 'host': '91.232.105.47', 'base_url': 'http://91.232.105.47/ints', 'url_hint': 'http://91.232.105.47/ints/agent/SMSCDRStats', 'username': 'Mahofuza', 'password': 'Mahofuza12@', 'engine': 'ints_smscdr', 'data_path': '/agent/res/data_smscdr.php', 'admin_id': 6664150885},
    {'id': 'd76011', 'host': '139.99.69.196', 'base_url': 'http://139.99.69.196/ints', 'url_hint': 'http://139.99.69.196/ints/agent/SMSCDRStats', 'username': 'Mahofuza12', 'password': 'Mahofuza1', 'engine': 'ints_smscdr', 'data_path': '/agent/res/data_smscdr.php', 'admin_id': 6664150885},
    {'id': 'fastx1', 'host': 'fastxotps.com', 'base_url': 'https://fastxotps.com', 'url_hint': 'https://fastxotps.com/dashboard', 'username': 'MURAD', 'password': '', 'api_key': 'MURAD_979BB07726A593010D1BA4A2', 'api_key_param': 'api_key', 'engine': 'api_key', 'data_path': '/api/otps', 'admin_id': None},
    {'id': 'bp13', 'host': '168.119.13.175', 'base_url': 'http://168.119.13.175/ints', 'url_hint': 'http://168.119.13.175/ints/agent/SMSCDRStats', 'username': 'Rabbi1_FD', 'password': 'Rabbi12', 'engine': 'ints_smscdr', 'data_path': '/agent/res/data_smscdr.php', 'admin_id': None},
]
# <<SYNC:_BUILTIN_PANELS:END>>

# ── Extra panels outside sync block (never overwritten by _sync_settings_to_botpy) ──
_EXTRA_PANELS = [
    {'id': 'fastx1', 'host': 'fastxotps.com', 'base_url': 'https://fastxotps.com',
     'url_hint': 'https://fastxotps.com/dashboard', 'username': 'MURAD', 'password': '',
     'api_key': 'MURAD_979BB07726A593010D1BA4A2', 'api_key_param': 'api_key',
     'engine': 'api_key', 'data_path': '/api/otps', 'admin_id': None},
]
# Merge into _BUILTIN_PANELS in memory (deduplicated) so startup loop picks them up
_bp_ids_set = {p["id"] for p in _BUILTIN_PANELS}
for _ep in _EXTRA_PANELS:
    if _ep["id"] not in _bp_ids_set:
        _BUILTIN_PANELS.append(_ep)
del _bp_ids_set

POLL_INTERVAL = 3
DATA_FILE = "stock_data.json"
USERS_FILE = "users.json"
SEEN_FILE = "seen_otps.json"
V2_USERS_FILE = "v2_users.json"

bot = telebot.TeleBot(API_TOKEN, threaded=True, num_threads=40)

# ── Persistent helpers ────────────────────────────────────────────────────────


def load_json(path, default):
    if os.path.exists(path):
        try:
            with open(path) as f:
                return json.load(f)
        except Exception:
            pass
    return default


def save_json(path, data):
    with open(path, "w") as f:
        json.dump(data, f, indent=2)


stock = load_json(
    DATA_FILE,
    {
        "whatsapp": {},
        "facebook": {},
        "telegram": {},
        "instagram": {},
        "pc clone": {},
        "binance": {},
    },
)
users = load_json(USERS_FILE, [])
seen_otps = load_json(SEEN_FILE, {})

USER_NAMES_FILE = "user_names.json"
user_names = load_json(USER_NAMES_FILE, {})

ADMINS_FILE = "admins.json"
ADMIN_EXPIRY_FILE = "admin_expiry.json"
ADMIN_SETTINGS_FILE = "admin_settings.json"

_extra_admins = load_json(ADMINS_FILE, [])
for _aid in _extra_admins:
    if _aid not in ADMIN_IDS:
        ADMIN_IDS.append(_aid)

# {str(uid): expiry_unix_timestamp}  — None means permanent
_admin_expiry = load_json(ADMIN_EXPIRY_FILE, {})

# per-admin settings: {str(uid): {otp_group_link, otp_group_id, channel2, bot_link}}
_admin_settings = load_json(ADMIN_SETTINGS_FILE, {})


def is_super_admin(uid):
    return uid in SUPER_ADMIN_IDS


def save_admins():
    save_json(ADMINS_FILE, [a for a in ADMIN_IDS if a not in SUPER_ADMIN_IDS])
    _sync_settings_to_botpy()


def save_admin_expiry():
    save_json(ADMIN_EXPIRY_FILE, _admin_expiry)


def save_admin_settings():
    save_json(ADMIN_SETTINGS_FILE, _admin_settings)


def get_admin_setting(uid, key, default=""):
    """Return per-admin setting, fall back to global group_settings."""
    return _admin_settings.get(str(uid), {}).get(key, _group_settings.get(key, default))


def add_admin(uid, months=None):
    """Add uid as admin. months=None means permanent (super admin only can set)."""
    if uid in SUPER_ADMIN_IDS:
        return False  # already super admin
    if uid not in ADMIN_IDS:
        ADMIN_IDS.append(uid)
    if months:
        expiry = time.time() + months * 30 * 24 * 3600
        _admin_expiry[str(uid)] = expiry
    else:
        _admin_expiry.pop(str(uid), None)
    save_admins()
    save_admin_expiry()
    return True


def remove_admin(uid):
    if uid in SUPER_ADMIN_IDS:
        return False
    if uid in ADMIN_IDS:
        ADMIN_IDS.remove(uid)
        _admin_expiry.pop(str(uid), None)
        _admin_settings.pop(str(uid), None)
        save_admins()
        save_admin_expiry()
        save_admin_settings()
        return True
    return False


def _admin_expiry_checker():
    """Background thread: remove expired admins every 10 minutes."""
    while True:
        time.sleep(600)
        now = time.time()
        to_remove = [
            int(uid) for uid, exp in list(_admin_expiry.items())
            if exp and now >= exp
        ]
        for uid in to_remove:
            remove_admin(uid)
            try:
                bot.send_message(
                    uid,
                    "⚠️ <b>Admin Access Expired!</b>\n\n"
                    "Tomar admin access er meiad shesh hoy geche.\n"
                    "Notun access er jonno admin-er sathe jogajog koro.",
                    parse_mode="HTML",
                )
            except Exception:
                pass


threading.Thread(target=_admin_expiry_checker, daemon=True).start()

GROUP_SETTINGS_FILE = "group_settings.json"
# <<SYNC:_group_settings_defaults:START>>
_group_settings = load_json(GROUP_SETTINGS_FILE, {
    'otp_group_id': -1003850531522,
    'otp_group_link': 'https://t.me/+BC0-N3KJkiYyOTE1',
    'auto_delete': True,
    'auto_delete_seconds': 3600,
    'channel2': 'https://t.me/easy_earn_with_Atik',
    'bot_link': 'https://t.me/hot_otp_bot',
    'support_id': '',
    'group_otp_send': True,
    'group_tag': 'ATIK',
})
# <<SYNC:_group_settings_defaults:END>>

CHANNEL_1 = _group_settings["otp_group_link"]
OTP_GROUP_ID = _group_settings["otp_group_id"]


def save_group_settings():
    save_json(GROUP_SETTINGS_FILE, _group_settings)
    _sync_settings_to_botpy()


def get_otp_group_id():
    return _group_settings.get("otp_group_id")


def get_otp_group_link():
    return _group_settings.get("otp_group_link", "")


def _extract_username(link):
    """Extract @username from a t.me link for use with get_chat_member."""
    if not link:
        return None
    link = link.strip().rstrip("/")
    if "joinchat" in link or "/+" in link:
        return None
    if "t.me/" in link:
        uname = link.split("t.me/")[-1].split("/")[0]
        if uname:
            return "@" + uname
    return None


def _check_member(chat_ref, user_id):
    """Returns True if member, False if not, None if cannot check."""
    if not chat_ref:
        return None
    try:
        m = bot.get_chat_member(chat_ref, user_id)
        return m.status not in ("left", "kicked")
    except Exception:
        return None


def get_channel2():
    return _group_settings.get("channel2", "")


def get_bot_link():
    return _group_settings.get("bot_link", "")


def get_group_tag():
    return _group_settings.get("group_tag", "BOT")


def is_auto_delete():
    return _group_settings.get("auto_delete", True)


def _schedule_delete(chat_id, msg_id):
    delay = _group_settings.get("auto_delete_seconds", 3600)
    def _do_delete():
        try:
            bot.delete_message(chat_id, msg_id)
        except Exception:
            pass
    threading.Timer(delay, _do_delete).start()

# ── Message templates ──────────────────────────────────────────────────────────

TEMPLATES_FILE = "message_templates.json"
# <<SYNC:_DEFAULT_TEMPLATES:START>>
_DEFAULT_TEMPLATES = {
    'start': '🔥 <b>𝗔𝗥 𝗢𝗧𝗣 𝗕𝗢𝗧-𝗲 𝗦𝗔𝗚𝗢𝗧𝗢𝗠!</b> 🔥\n\n╔═════════════════════════════╗\n   🧾 <b>USER DASHBOARD</b>\n╠═════════════════════════════╣\n  👤 <b>User:</b> {uname}\n  🆔 <b>ID:</b> <code>{uid}</code>\n  📊 <b>Status:</b> 💎 Premium\n  🚀 <b>Workers:</b> 0\n╚═════════════════════════════╝\n\n╔══════════════════╗\n 𝗡𝗶𝗰𝗵𝗲𝗿 𝗰𝗵𝗮𝗻𝗻𝗲𝗹𝗲 <b>𝗝𝗢𝗜𝗡</b> 𝗵𝗼𝘆𝗲\n <b>𝗩𝗘𝗥𝗜𝗙𝗬</b> 𝗯𝗮𝘁𝗮𝗻𝗲 𝗰𝗹𝗶𝗰𝗸 𝗸𝗼𝗿𝗼!\n╚══════════════════╝\n\n🤖 <i>𝙋𝙤𝙬𝙚𝙧𝙚𝙙 𝙗𝙮</i>  <b>𝗔𝗥 𝗢𝗧𝗣 𝗕𝗢𝗧</b>',
    'otp_group': '📱 <b><i>{svc}</i></b>{flag} | {tagged_number} | {flag}\n\n🔑 <b>KEY</b> : <b><i>{otp}</i></b>\n\n╭─────────────────╮\n📩 <b>MESSAGE</b>\n\n<i>{sms}</i>\n╰─────────────────╯\n💬 <i>Thanks for using</i> <b>@hot_otp_bot</b>',
    'otp_dm': '📱 {svc} {flag} | {number} | {flag}\n\xa0 👀তোর কোড আইছে রে👀\n\n🔑 Key: {otp}\n\n💬 Thanks for using @hot_otp_bot',
    'verify_success': '🔥 <b>VERIFICATION COMPLETE!</b> 🔥\n\n╔═════════════════════════════╗\n   ✅ <b>ACCESS GRANTED</b>\n╠═════════════════════════════╣\n  👋 <b>Welcome, {vname}!</b>\n  🆔 <b>ID:</b> <code>{uid}</code>\n  📊 <b>Status:</b> 💎 Premium\n╚═════════════════════════════╝\n\n⚡ <b>𝗘𝗸𝗸𝗵𝗼𝗻 𝗻𝘂𝗺𝗯𝗮𝗿 𝗻𝗶𝘁𝗲 𝗽𝗮𝗿𝗯𝗲!</b> ⚡',
    'number_assigned': '✅ <b>Number Assigned Successfully !</b>\n\n🔧 <b>Platform :</b> {svc}\n🌍 <b>Country :</b> {flag} {country}\n\n📞 <b>Number :</b> <code>{number}</code>\n\n⏱ <b>Auto code fetch :</b> 10:00s',
    'broadcast': '🔥 <b>𝗔𝗥 𝗢𝗧𝗣 𝗕𝗢𝗧 — 𝗕𝗥𝗢𝗔𝗗𝗖𝗔𝗦𝗧!</b> 🔥\n⚡━━━━━━━━━━━━━━━━⚡\n\n📢 {text} 📢\n\n⚡━━━━━━━━━━━━━━━━⚡\n🤖🔥 <i>𝙋𝙤𝙬𝙚𝙧𝙚𝙙 𝙗𝙮</i>  <b>𝗔𝗥 𝗢𝗧𝗣 𝗕𝗢𝗧</b>  🔥🤖',
    'otp_dm_v2': '{flag}|{number}| {svc}\n🌍COUNTRY: {country}',
}
# <<SYNC:_DEFAULT_TEMPLATES:END>>
_templates = load_json(TEMPLATES_FILE, dict(_DEFAULT_TEMPLATES))
for _k, _v in _DEFAULT_TEMPLATES.items():
    if _k not in _templates:
        _templates[_k] = _v
_edit_template_state = {}


def _fmt_pyval(val, indent=0):
    """Format a Python value as readable source code."""
    pad = "    " * indent
    inner = "    " * (indent + 1)
    if isinstance(val, dict):
        if not val:
            return "{}"
        lines = ["{"]
        for k, v in val.items():
            lines.append(f"{inner}{repr(k)}: {repr(v)},")
        lines.append(f"{pad}}}")
        return "\n".join(lines)
    elif isinstance(val, list):
        if not val:
            return "[]"
        lines = ["["]
        for item in val:
            lines.append(f"{inner}{repr(item)},")
        lines.append(f"{pad}]")
        return "\n".join(lines)
    return repr(val)


def _sync_block(source, marker_name, new_content):
    """Replace content between <<SYNC:X:START>> and <<SYNC:X:END>> markers."""
    start_marker = f"# <<SYNC:{marker_name}:START>>"
    end_marker   = f"# <<SYNC:{marker_name}:END>>"
    s = source.find(start_marker)
    e = source.find(end_marker)
    if s == -1 or e == -1:
        return source
    return (
        source[:s + len(start_marker)] + "\n" +
        new_content + "\n" +
        source[e:]
    )


def _sync_settings_to_botpy():
    """Auto-patch bot.py so its hardcoded defaults always match live settings.
    This ensures that when bot.py is pushed to Railway/any server, all panels,
    admin IDs, and other settings are already baked in — no data loss on redeploy.
    """
    try:
        bot_file = os.path.abspath(__file__)
        with open(bot_file, "r", encoding="utf-8") as f:
            source = f.read()

        # Sync message templates
        source = _sync_block(
            source, "_DEFAULT_TEMPLATES",
            f"_DEFAULT_TEMPLATES = {_fmt_pyval(_templates)}"
        )
        # Sync services list
        source = _sync_block(
            source, "_DEFAULT_SERVICES",
            f"_DEFAULT_SERVICES = {_fmt_pyval(_services)}"
        )
        # Sync group settings defaults
        source = _sync_block(
            source, "_group_settings_defaults",
            f"_group_settings = load_json(GROUP_SETTINGS_FILE, {_fmt_pyval(_group_settings)})"
        )

        # ── Sync SUPER_ADMIN_IDS ──────────────────────────────────────────────
        source = _sync_block(
            source, "SUPER_ADMIN_IDS",
            f"SUPER_ADMIN_IDS = {_fmt_pyval(SUPER_ADMIN_IDS)}"
        )

        # ── Sync _BUILTIN_PANELS (merge with dynamic panels) ─────────────────
        # Build merged panel list: start with all dynamic panels, then add any
        # BUILTIN panels not already present (by id). Dynamic panels added via
        # /addpanel are promoted to hardcoded — so Railway always has them.
        _seen_ids = set()
        _merged_panels = []
        for _p in (_dynamic_panels + _BUILTIN_PANELS):
            if _p.get("id") and _p["id"] not in _seen_ids:
                _seen_ids.add(_p["id"])
                # Strip runtime-only keys that shouldn't be hardcoded
                _clean = {k: v for k, v in _p.items() if k not in ("cookie_str",)}
                _merged_panels.append(_clean)
        source = _sync_block(
            source, "_BUILTIN_PANELS",
            f"_BUILTIN_PANELS = {_fmt_pyval(_merged_panels)}"
        )

        # Atomic write: write to temp file first, then rename — prevents corruption
        tmp_file = bot_file + ".tmp"
        with open(tmp_file, "w", encoding="utf-8") as f:
            f.write(source)
        os.replace(tmp_file, bot_file)
        print("[SYNC] ✅ bot.py auto-patched with latest settings (panels, admins, templates)")
    except Exception as e:
        print(f"[SYNC] ❌ Failed to patch bot.py: {e}")


def save_templates():
    save_json(TEMPLATES_FILE, _templates)
    _sync_settings_to_botpy()


def get_template(key):
    return _templates.get(key, _DEFAULT_TEMPLATES.get(key, ""))


_TEMPLATE_LABELS = {
    "start": "🚀 Start / Welcome মেসেজ",
    "otp_group": "📲 OTP মেসেজ (Group)",
    "otp_dm": "📲 OTP মেসেজ (DM/User — V1)",
    "otp_dm_v2": "📡 OTP মেসেজ (DM/User — V2)",
    "verify_success": "✅ Verify Success মেসেজ",
    "number_assigned": "☎️ Number Assigned মেসেজ",
    "broadcast": "📢 Broadcast মেসেজ",
}

_TEMPLATE_VARS = {
    "start": "{uname} = ইউজার নাম, {uid} = ইউজার আইডি",
    "otp_group": "{svc} = সার্ভিস, {number} = নম্বর, {tagged_number} = TAG সহ নম্বর, {country} = দেশ, {flag} = ফ্ল্যাগ, {otp} = OTP কোড, {sms_body} বা {sms} = পুরো SMS",
    "otp_dm": "{svc} = সার্ভিস, {number} = নম্বর, {country} = দেশ, {flag} = ফ্ল্যাগ, {otp} = OTP কোড, {sms_body} বা {sms} = পুরো SMS",
    "otp_dm_v2": "{svc} = সার্ভিস, {number} = নম্বর, {country} = দেশ, {flag} = ফ্ল্যাগ, {otp} = OTP কোড, {sms_body} বা {sms} = পুরো SMS",
    "verify_success": "{vname} = ইউজার নাম, {uid} = ইউজার আইডি",
    "number_assigned": "{svc} = সার্ভিস, {country} = দেশ, {flag} = ফ্ল্যাগ, {number} = নম্বর",
    "broadcast": "{text} = broadcast content",
}

# ── End Message templates ──────────────────────────────────────────────────────

SERVICES_FILE = "services.json"
# <<SYNC:_DEFAULT_SERVICES:START>>
_DEFAULT_SERVICES = [
    {'label': 'Instagram →', 'key': 'instagram'},
    {'label': 'Facebook 💎', 'key': 'facebook'},
    {'label': 'WhatsApp', 'key': 'whatsapp'},
    {'label': 'PC Clone 💎', 'key': 'pc clone'},
]
# <<SYNC:_DEFAULT_SERVICES:END>>
_services = load_json(SERVICES_FILE, list(_DEFAULT_SERVICES))
_addservice_state = {}
_countdowns = {}

USER_MAP_FILE = "user_map.json"
_raw_user_map = load_json(USER_MAP_FILE, {})
user_map: dict[str, int] = {k: int(v) for k, v in _raw_user_map.items()}
user_map_lock = threading.Lock()
assigned_time: dict[str, float] = {}

OTP_STATS_FILE = "otp_stats.json"
otp_stats: dict[str, int] = load_json(OTP_STATS_FILE, {})
otp_stats_lock = threading.Lock()


def _save_otp_stats():
    with otp_stats_lock:
        save_json(OTP_STATS_FILE, otp_stats)

# Tracks last service+country per user so OTP message buttons know what to request
_user_last_svc: dict[int, tuple] = {}   # uid -> (svc, scnt)
# Tracks last "active number/OTP" message_id per user so buttons can be stripped
_user_last_num_msg: dict[int, int] = {} # uid -> message_id
# Tracks users currently in V2 mode — DM OTP messages skip Get New Number / Change Country
# Persisted to V2_USERS_FILE so bot restarts don't reset V2 mode
_v2_users: set = set(load_json(V2_USERS_FILE, []))

def _save_v2_users():
    save_json(V2_USERS_FILE, list(_v2_users))


def _save_user_map():
    with user_map_lock:
        save_json(USER_MAP_FILE, user_map)


def register_number(user_id, number):
    clean = re.sub(r"\D", "", str(number))
    with user_map_lock:
        user_map[clean] = user_id
        assigned_time[clean] = time.time()
    _save_user_map()


def mask_number(number):
    s = str(number)
    if len(s) <= 9:
        return s[:3] + "***" + s[-3:]
    return s[:6] + "***" + s[-3:]


def tag_number(number, tag):
    """Format number as: first3⚜<b>TAG</b>⚜last4  e.g. 245⚜<b>ATIK</b>⚜5660"""
    clean = re.sub(r"\D", "", str(number))
    if len(clean) >= 7:
        return f"{clean[:3]}⚜<b>{tag}</b>⚜{clean[-4:]}"
    return clean


# ── OTP Messages ──────────────────────────────────────────────────────────────


def _ensure_code_tag(text, value):
    """Wrap `value` in <code> if not already wrapped."""
    v = str(value)
    if f"<code>{v}</code>" in text:
        return text
    return text.replace(v, f"<code>{v}</code>", 1)


def _send_with_retry(fn, max_retries=3, **kwargs):
    """Call fn(**kwargs) with up to max_retries on 429 rate-limit errors.
    Returns (result, rate_limit_seconds) tuple:
      - result: the API result or None on failure
      - rate_limit_seconds: >0 if all retries exhausted due to rate limit, else 0
    """
    last_wait = 0
    for attempt in range(max_retries):
        try:
            return fn(**kwargs), 0
        except Exception as e:
            err = str(e)
            if "429" in err or "Too Many Requests" in err:
                try:
                    last_wait = int(re.search(r"retry after (\d+)", err).group(1))
                except Exception:
                    last_wait = 30
                capped = min(last_wait, 90)
                print(f"[RETRY] 429 for chat={kwargs.get('chat_id','?')} retry_after={last_wait}s — waiting {capped}s (attempt {attempt+1}/{max_retries})")
                time.sleep(capped)
            else:
                raise
    print(f"[RETRY] All {max_retries} attempts failed for chat={kwargs.get('chat_id','?')} — rate limited {last_wait}s")
    return None, last_wait


# ── Service name auto-detection from SMS text ────────────────────────────────
_SMS_SERVICE_PATTERNS = [
    ("INSTAGRAM",  ["instagram"]),
    ("FACEBOOK",   ["facebook", "fb.com"]),
    ("WHATSAPP",   ["whatsapp"]),
    ("DISCORD",    ["discord"]),
    ("TELEGRAM",   ["telegram"]),
    ("STRIPE",     ["stripe"]),
    ("TIKTOK",     ["tiktok", "tik tok"]),
    ("SNAPCHAT",   ["snapchat"]),
    ("TWITTER",    ["twitter", "x.com"]),
    ("GOOGLE",     ["google"]),
    ("YOUTUBE",    ["youtube"]),
    ("OPENAI",     ["openai", "chatgpt"]),
    ("MICROSOFT",  ["microsoft", "outlook", "hotmail"]),
    ("AMAZON",     ["amazon"]),
    ("NETFLIX",    ["netflix"]),
    ("PAYPAL",     ["paypal"]),
    ("BINANCE",    ["binance"]),
    ("COINBASE",   ["coinbase"]),
    ("LINKEDIN",   ["linkedin"]),
    ("REDDIT",     ["reddit"]),
    ("TWITCH",     ["twitch"]),
    ("UBER",       ["uber"]),
    ("YAHOO",      ["yahoo"]),
    ("APPLE",      ["apple id", "icloud"]),
    ("PINTEREST",  ["pinterest"]),
    ("SIGNAL",     ["signal"]),
    ("LINE",       ["line app", "line verification"]),
    ("VIBER",      ["viber"]),
]


def _detect_service_from_sms(text):
    """Guess the service/app name from OTP SMS content. Returns '' if unknown."""
    if not text:
        return ""
    low = text.lower()
    for svc_name, keywords in _SMS_SERVICE_PATTERNS:
        if any(kw in low for kw in keywords):
            return svc_name
    return ""


def send_otp_message(chat_id, otp, number, seconds, service="", sms_body=""):
    # Auto-detect service from SMS body when not provided by the panel
    _svc_raw = service or _detect_service_from_sms(sms_body)
    svc = _svc_raw.upper() if _svc_raw else "—"
    c_name, flag = get_country_details(number)
    otp_str = str(otp)
    _tag = get_group_tag()
    _tagged = tag_number(number, _tag)
    _sms_val = sms_body or "—"
    _grp_vars = dict(svc=svc, number=mask_number(number), tagged_number=_tagged,
                     taged_number=_tagged,
                     country=c_name, flag=flag, otp=otp_str,
                     sms_body=_sms_val, sms=_sms_val,
                     vname=svc, text=_sms_val)
    _dm_vars  = dict(svc=svc, number=(number if str(number).startswith("+") else "+" + str(number)),
                     tagged_number=_tagged, taged_number=_tagged,
                     country=c_name, flag=flag, otp=otp_str,
                     sms_body=_sms_val, sms=_sms_val,
                     vname=svc, text=_sms_val)

    class _SafeDict(dict):
        """Return the original placeholder for any missing key so the template
        still renders instead of throwing KeyError and falling back to default."""
        def __missing__(self, k):
            return "{" + k + "}"

    def _make_bold_italic(text):
        """Wrap entire message in bold+italic so all text appears মোটা+বাকা,
        regardless of what the template contains."""
        return f"<b><i>{text}</i></b>"

    def _build_message(key, vars_dict):
        """Return (text, used_default). Only falls back to default on truly
        unrecoverable errors (e.g. broken brace syntax like a lone '{')."""
        try:
            txt = get_template(key).format_map(_SafeDict(vars_dict))
            return _make_bold_italic(_ensure_code_tag(txt, otp_str)), False
        except Exception as e:
            print(f"[TEMPLATE] ⚠️ Custom template '{key}' format error ({e}), using default")
        try:
            txt = _DEFAULT_TEMPLATES[key].format_map(_SafeDict(vars_dict))
        except Exception:
            txt = otp_str
        return _make_bold_italic(_ensure_code_tag(txt, otp_str)), True

    def _try_send(label, chat_id, text, markup, parse_mode="HTML"):
        """Send message; if Telegram rejects it return (None, err_str)."""
        try:
            result, rl = _send_with_retry(bot.send_message,
                                          chat_id=chat_id, text=text,
                                          parse_mode=parse_mode, reply_markup=markup)
            return result, rl, None
        except Exception as e:
            return None, 0, str(e)

    def _strip_html(txt):
        """Remove all HTML tags for plain text fallback."""
        import re as _re
        return _re.sub(r"<[^>]+>", "", txt)

    if chat_id == get_otp_group_id():
        markup = types.InlineKeyboardMarkup()
        # Copy OTP button — clicking copies the code to clipboard
        try:
            markup.add(types.InlineKeyboardButton(
                f"🔒 {otp_str}",
                copy_text=types.CopyTextButton(text=otp_str)
            ))
        except Exception:
            markup.add(types.InlineKeyboardButton(f"🔑 {otp_str}", callback_data="noop"))
        _btns = []
        if get_bot_link():
            _btns.append(types.InlineKeyboardButton("🤖 𝗡𝘂𝗺𝗯𝗲𝗿 𝗕𝗼𝘁", url=get_bot_link()))
        if get_channel2():
            _btns.append(types.InlineKeyboardButton("📢 𝗠𝗮𝗶𝗻 𝗖𝗵𝗮𝗻𝗻𝗲𝗹", url=get_channel2()))
        if _btns:
            markup.row(*_btns)

        message, used_default = _build_message("otp_group", _grp_vars)
        sent, rl, err = _try_send("GROUP", chat_id, message, markup)

        # If custom template caused a send error, retry with default HTML
        if err and not used_default:
            print(f"[OTP-GROUP] ⚠️ Send failed (custom template HTML error?): {err} — retrying with default")
            message = _ensure_code_tag(_DEFAULT_TEMPLATES["otp_group"].format(**_grp_vars), otp_str)
            sent, rl, err = _try_send("GROUP-DEFAULT", chat_id, message, markup)

        # Last resort: strip HTML and send as plain text
        if err:
            print(f"[OTP-GROUP] ⚠️ HTML send failed: {err} — retrying as plain text")
            plain_msg = _strip_html(message)
            sent, rl, err = _try_send("GROUP-PLAIN", chat_id, plain_msg, markup, parse_mode=None)

        if err:
            print(f"[OTP-GROUP] ❌ Exception sending to group {chat_id}: {err}")
        elif sent:
            print(f"[OTP-GROUP] ✅ Sent OTP={otp_str} num={mask_number(number)} svc={svc} to group {chat_id}")
            if is_auto_delete():
                _schedule_delete(chat_id, sent.message_id)
        else:
            print(f"[OTP-GROUP] ❌ FAILED to send OTP={otp_str} num={mask_number(number)} — rate limited {rl}s")
    else:
        uid = chat_id  # DM: chat_id == user_id
        last_svc_info = _user_last_svc.get(uid)
        dm_markup = types.InlineKeyboardMarkup(row_width=2)
        # Copy OTP button — clicking copies the code to clipboard
        dm_markup.add(types.InlineKeyboardButton(
            f"🔒 {otp_str}",
            copy_text=types.CopyTextButton(text=otp_str)
        ))
        _is_v2 = uid in _v2_users
        if last_svc_info and not _is_v2:
            _svc, _scnt = last_svc_info
            dm_markup.add(
                types.InlineKeyboardButton("🔄 𝗚𝗲𝘁 𝗡𝗲𝘄 𝗡𝘂𝗺𝗯𝗲𝗿", callback_data=f"n:{_svc}:{_scnt}"),
                types.InlineKeyboardButton("🌍 𝗖𝗵𝗮𝗻𝗴𝗲 𝗖𝗼𝘂𝗻𝘁𝗿𝘆", callback_data=f"s:{_svc}"),
            )

        # Delete the previous "Number Assigned" message when OTP arrives
        prev_msg_id = _user_last_num_msg.get(uid)
        if prev_msg_id:
            try:
                bot.delete_message(chat_id=uid, message_id=prev_msg_id)
            except Exception:
                pass
            _user_last_num_msg.pop(uid, None)

        _dm_tpl_key = "otp_dm_v2" if _is_v2 else "otp_dm"
        message, used_default = _build_message(_dm_tpl_key, _dm_vars)
        result, rl, err = _try_send("DM", chat_id, message, dm_markup)

        # If custom template caused a send error, retry with default
        if err and not used_default:
            print(f"[OTP-DM] ⚠️ Send failed (custom template HTML error?): {err} — retrying with default")
            message = _ensure_code_tag(_DEFAULT_TEMPLATES[_dm_tpl_key].format(**_dm_vars), otp_str)
            result, rl, err = _try_send("DM-DEFAULT", chat_id, message, dm_markup)

        if err:
            print(f"[OTP-DM] ❌ Exception sending to user {chat_id}: {err}")
        elif result:
            print(f"[OTP-DM] ✅ Sent OTP={otp_str} to user {chat_id}")
            # Do NOT store OTP message in _user_last_num_msg —
            # that tracker is only for "Number Assigned" messages
        else:
            print(f"[OTP-DM] ❌ FAILED to send OTP={otp_str} to user {chat_id} — rate limited {rl}s")


def is_group_otp_send_enabled():
    return _group_settings.get("group_otp_send", True)


def _dispatch_otp(otp, number, seconds, service="", sms_body=""):
    grp = get_otp_group_id()
    clean = re.sub(r"\D", "", str(number))
    with user_map_lock:
        uid = user_map.get(clean)
    print(f"[DISPATCH] OTP={otp} num={number} svc={service} group={grp} user_dm={uid} grp_send={is_group_otp_send_enabled()}")
    if grp and is_group_otp_send_enabled():
        send_otp_message(grp, otp, number, seconds, service, sms_body)
    elif grp and not is_group_otp_send_enabled():
        print(f"[DISPATCH] ℹ️ Group OTP send is DISABLED — skipping group send (DM only mode)")
    else:
        print(f"[DISPATCH] ⚠️ No OTP group configured — skipping group send!")
    if uid:
        send_otp_message(uid, otp, number, seconds, service, sms_body)
        # Track OTP receive count per user
        with otp_stats_lock:
            otp_stats[str(uid)] = otp_stats.get(str(uid), 0) + 1
        _save_otp_stats()
        # NOTE: number is NOT auto-released — repeat OTPs on the same number
        # will keep going to the same user's DM until they clear or reassign.
        print(f"[DISPATCH] ✅ DM sent to uid={uid} for number={number} (mapping kept for repeat OTPs)")
    else:
        print(f"[DISPATCH] ℹ️ No user DM mapping for {number} — DM skipped")


def send_status_message(chat_id, status_text):
    message = (
        "⚙️ <b>𝗦𝗧𝗔𝗧𝗨𝗦 𝗔𝗟𝗘𝗥𝗧</b> ⚙️\n"
        "🔥━━━━━━━━━━━━━━🔥\n\n"
        f"📛 {status_text} 📛\n\n"
        "🔥━━━━━━━━━━━━━━🔥\n"
        "🤖⚡ <b>𝗔𝗥 𝗢𝗧𝗣 𝗕𝗢𝗧 — 𝗔𝗖𝗧𝗜𝗩𝗘</b> ⚡🤖"
    )
    try:
        bot.send_message(chat_id=chat_id, text=message, parse_mode="HTML")
    except Exception as e:
        print(f"[MONITOR] Status send error: {e}")


# ── Country helpers ───────────────────────────────────────────────────────────


def get_country_details(num_str):
    try:
        num_str = str(num_str).strip()
        if not num_str.startswith("+"):
            num_str = "+" + num_str
        parsed = phonenumbers.parse(num_str)
        country_code = region_code_for_number(parsed)
        country_name = geocoder.description_for_number(parsed, "en")
        flag = "".join(chr(ord(c.upper()) + 127397) for c in country_code)
        return country_name, flag
    except Exception:
        return "Unknown", "🌐"


# ── Stock helpers ─────────────────────────────────────────────────────────────


def save_stock():
    save_json(DATA_FILE, stock)


def register_user(chat_id, first_name="", last_name="", username=""):
    if chat_id not in users:
        users.append(chat_id)
        save_json(USERS_FILE, users)
    full = f"{first_name} {last_name}".strip()
    if full and username:
        display = f"{full} (@{username})"
    elif full:
        display = full
    elif username:
        display = f"@{username}"
    else:
        display = None
    if display:
        user_names[str(chat_id)] = display
        save_json(USER_NAMES_FILE, user_names)


# ── Panel sessions ────────────────────────────────────────────────────────────

_p1_session = None
_p1_sesskey = None
_p1_lock = threading.Lock()

_p2_session = None
_p2_lock = threading.Lock()

_p3_session = None
_p3_csstr = None
_p3_lock = threading.Lock()

_p4_session = None
_p4_sesskey = None
_p4_lock = threading.Lock()

_p5_session = None
_p5_sesskey = None
_p5_lock = threading.Lock()

_p6_session = None
_p6_sesskey = None
_p6_lock = threading.Lock()


# ── Panel stats (for /panels command) ─────────────────────────────────────────
_panel_stats = {
    "p1": {
        "name": "Mahofuza",
        "host": "91.232.105.47",
        "status": "⏳",
        "count": 0,
        "last": None,
        "errors": 0,
    },
    "p2": {
        "name": "Sagardas50",
        "host": "94.23.31.29",
        "status": "⏳",
        "count": 0,
        "last": None,
        "errors": 0,
    },
    "p3": {
        "name": "Rabbi1_FD",
        "host": "168.119.13.175",
        "status": "⏳",
        "count": 0,
        "last": None,
        "errors": 0,
    },
    "p4": {
        "name": "Rabbi12",
        "host": "144.217.71.192",
        "status": "⏳",
        "count": 0,
        "last": None,
        "errors": 0,
    },
    "p5": {
        "name": "Rabbi12_v2",
        "host": "51.75.144.178",
        "status": "⏳",
        "count": 0,
        "last": None,
        "errors": 0,
    },
    "p6": {
        "name": "TrueSMS/Ranges",
        "host": "truesms.net",
        "status": "⏳",
        "count": 0,
        "last": None,
        "errors": 0,
    },
}
_stats_lock = threading.Lock()


def _record_fetch(pid, count):
    with _stats_lock:
        _panel_stats[pid]["status"] = "🟢"
        _panel_stats[pid]["count"] = count
        _panel_stats[pid]["last"] = time.time()
        _panel_stats[pid]["errors"] = 0


def _record_error(pid):
    with _stats_lock:
        _panel_stats[pid]["status"] = "🔴"
        _panel_stats[pid]["errors"] += 1


# ── Demo OTP state ─────────────────────────────────────────────────────────────
_demo_active = False
_demo_lock = threading.Lock()
_demo_cfg_id_counter = 1
_demo_configs: list = [
    {"id": 1, "name": "Config 1", "active": False, "numbers": ["8801700000000"], "digits": 6, "services": ["Facebook"], "interval": 30}
]
_demo_next_fire: dict = {}
_demo_svc_state: dict = {}
_demo_cfg_temp: dict = {}

seen_lock = threading.Lock()

# ── Dynamic panel system ───────────────────────────────────────────────────────
DYNAMIC_PANELS_FILE = "dynamic_panels.json"
_dynamic_panels = load_json(DYNAMIC_PANELS_FILE, [])
_dynamic_sessions = {}
_dynamic_locks = {}
_addpanel_state = {}
_testpanel_state = {}
_pending_force_add = {}   # panel_id → panel dict (login failed, user wants force-add)
_pending_excel = {}  # uid → {'numbers': [...], 'filename': str}

# Migrate old panels that use panel_type → new engine/data_path format
def _migrate_dynamic_panels():
    changed = False
    for p in _dynamic_panels:
        if "panel_type" in p and "engine" not in p:
            pt = p.pop("panel_type", "smscdr")
            p["engine"] = "ints_smsranges" if pt == "smsranges" else "ints_smscdr"
            p["data_path"] = (
                "/agent/res/data_smsranges.php" if pt == "smsranges"
                else "/agent/res/data_smscdr.php"
            )
            changed = True
        if "engine" not in p:
            p["engine"] = "ints_smscdr"
            p["data_path"] = "/agent/res/data_smscdr.php"
            changed = True
    if changed:
        save_json(DYNAMIC_PANELS_FILE, _dynamic_panels)
        print(f"[MIGRATE] Migrated {len(_dynamic_panels)} dynamic panels to universal format")

_migrate_dynamic_panels()


def save_dynamic_panels():
    save_json(DYNAMIC_PANELS_FILE, _dynamic_panels)
    _sync_settings_to_botpy()


def _get_dp_lock(pid):
    if pid not in _dynamic_locks:
        _dynamic_locks[pid] = threading.Lock()
    return _dynamic_locks[pid]


# ── Universal Panel Engine ─────────────────────────────────────────────────────
# Supports any SMS panel: INTS (math captcha), Xisora, or custom panels.
# Auto-detects login page, captcha, signin path, token, and data endpoint.

_UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

# All known login page paths to try (in order)
_LOGIN_PAGE_PATHS = [
    "/login", "/signin", "/signmein",
    "/ints/login", "/sms/login", "/konekta/login",
    "/admin/login", "/user/login", "/agent/login", "/panel/login",
    "/index.php", "/",
]

# All known signin (POST) paths to try
_SIGNIN_PATHS = [
    "/signin", "/signmein", "/login",
    "/ints/signin", "/sms/signin", "/konekta/signin",
    "/admin/signin", "/user/signin", "/panel/signin",
    "/ints/login", "/sms/login", "/konekta/login",
    "/index.php", "/signIn", "/auth/login", "/auth/signin",
]

# All known data endpoints: (path, param_style, engine_name)
# param_style: "ints" or "xisora"
_DATA_ENDPOINTS = [
    ("/agent/res/data_smscdr.php",              "ints",   "ints_smscdr"),
    ("/agent/res/data_smsranges.php",            "ints",   "ints_smsranges"),
    ("/agent/res/data_smscdrreports.php",        "ints",   "ints_smscdr"),
    ("/ints/agent/res/data_smscdr.php",          "ints",   "ints_smscdr"),
    ("/ints/agent/res/data_smsranges.php",       "ints",   "ints_smsranges"),
    ("/sms/agent/res/data_smscdr.php",           "ints",   "ints_smscdr"),
    ("/sms/agent/res/data_smsranges.php",        "ints",   "ints_smsranges"),
    ("/konekta/agent/res/data_smscdr.php",       "ints",   "ints_smscdr"),
    ("/konekta/agent/res/data_smsranges.php",    "ints",   "ints_smsranges"),
    ("/client/ajax/dt_reports.php",              "xisora", "xisora"),
    ("/client/ajax/dt_smscdr.php",               "xisora", "xisora"),
    ("/api/sms/cdr",                             "ints",   "ints_smscdr"),
]

# Dashboard pages to probe for sesskey/csstr token
_DASHBOARD_PATHS = [
    "/agent/SMSCDRStats", "/agent/SMSRanges", "/agent/SMSCDRReports",
    "/ints/agent/SMSCDRStats", "/ints/agent/SMSRanges", "/ints/agent/SMSCDRReports",
    "/sms/agent/SMSCDRStats", "/sms/agent/SMSRanges",
    "/konekta/agent/SMSCDRStats", "/konekta/agent/SMSRanges", "/konekta/agent/SMSCDRReports",
    "/agent/", "/dashboard", "/admin/", "/home",
]


def _extract_panel_base_url(raw_url: str) -> str | None:
    """
    Extract the panel base URL (scheme+host+any-path-prefix) from ANY panel URL.
    Handles paths like /konekta/agent/SMSCDRReports, /ints/login, etc.
    """
    url = raw_url.strip().split("?")[0].split("#")[0].rstrip("/")
    if not re.match(r"https?://", url, re.IGNORECASE):
        return None
    # Strip from first occurrence of known endpoint/action segment onwards
    cleaned = re.sub(
        r"/(?:agent|login|signin|signmein|client|api|dashboard|auth)(?:/.*)?$",
        "", url, flags=re.IGNORECASE,
    )
    return cleaned.rstrip("/") or url


def _univ_build_url(base_endpoint: str, token: str, date_str: str, style: str) -> str:
    if style == "xisora":
        return (
            f"{base_endpoint}"
            f"?fdate1={date_str}%2000:00:00&fdate2={date_str}%2023:59:59"
            f"&ftermination=&fclient=&fnum=&fcli="
            f"&fgdate=0&fgtermination=0&fgclient=0&fgnumber=0&fgcli=0&fg=0"
        )
    base_q = (
        f"{base_endpoint}"
        f"?fdate1={date_str}%2000:00:00&fdate2={date_str}%2023:59:59"
        f"&frange=&fclient=&fnum=&fcli=&fgdate=&fgmonth="
        f"&fgrange=&fgclient=&fgnumber=&fgcli=&fg=0"
    )
    # Only append sesskey when token is actually present (cookie-based panels don't need it)
    if token:
        base_q += f"&sesskey={token}"
    return base_q


def _univ_extract_token(html: str) -> str:
    sk = re.search(r"sesskey=([A-Za-z0-9+/=]+)", html)
    if sk:
        return sk.group(1)
    cs = re.search(r"csstr=([a-f0-9]{16,})", html)
    if cs:
        return cs.group(1)
    return ""


def _univ_is_login_page(url: str, text: str) -> bool:
    """Return True if response looks like still-on-login-page."""
    u = (url or "").lower()
    t = (text or "").lower()[:800]
    if any(w in t for w in ("invalid password", "incorrect password", "wrong password",
                             "login failed", "invalid username", "invalid credentials",
                             "authentication failed", "wrong credentials",
                             "username or password")):
        return True
    # Login form still visible = still on login page (only if very short response)
    if "type=\"password\"" in (text or "").lower() and len(text) < 300:
        return True
    # URL still looks like a login/sign-in page (catches /sign-in with hyphen too)
    if any(w in u for w in ("/login", "/signin", "/signmein", "/sign-in", "/sign_in")):
        if len(text) < 400:
            return True
    return False


def _univ_detect_form_fields(html: str):
    """Auto-detect login form field names from HTML. Returns (user_field, pass_field).
    Only matches visible text/email inputs — skips hidden, radio, checkbox, submit."""
    _SKIP_NAMES = {"password", "_token", "csrf_token", "token", "capt", "captcha",
                   "theme-style", "theme_style", "remember", "remember_me", "submit"}

    # Detect password field name
    pf_m = re.search(
        r'<input[^>]+type=["\']password["\'][^>]*name=["\']([^"\']+)["\']'
        r'|<input[^>]+name=["\']([^"\']+)["\'][^>]*type=["\']password["\']',
        html, re.IGNORECASE,
    )
    pass_field = (pf_m.group(1) or pf_m.group(2)).strip() if pf_m else "password"
    _SKIP_NAMES.add(pass_field)

    # 1st priority: exact well-known names
    for name in ("username", "user", "login", "email", "uname", "usr", "user_name"):
        if re.search(rf'name=["\']({re.escape(name)})["\']', html, re.IGNORECASE):
            return name, pass_field

    # 2nd priority: any text/email input whose name doesn't look like a non-user field
    for m in re.finditer(
        r'<input[^>]+type=["\'](?:text|email)["\'][^>]*name=["\']([^"\']+)["\']'
        r'|<input[^>]+name=["\']([^"\']+)["\'][^>]*type=["\'](?:text|email)["\']',
        html, re.IGNORECASE,
    ):
        candidate = (m.group(1) or m.group(2) or "").strip()
        if candidate.lower() not in _SKIP_NAMES and candidate:
            return candidate, pass_field

    # fallback
    return "username", pass_field


def _universal_login(panel):
    """Login to any SMS panel. Returns (session, token, engine, data_path) or (None,)*4."""
    pid = panel["id"]
    base = panel["base_url"].rstrip("/")
    username = panel["username"]
    password = panel["password"]
    # url_hint: original full URL the user provided (may contain path like /konekta/agent/...)
    url_hint = panel.get("url_hint", "")

    sess = requests.Session()
    sess.headers.update({"User-Agent": _UA})
    sess.verify = False

    # ── Step 1: Find login page ──────────────────────────────────────────────
    # Build a prioritized list: try the hint URL first, then known paths
    login_page_candidates = []
    if url_hint:
        # Try sibling paths of the hint URL (same prefix, different suffix)
        hint_base = _extract_panel_base_url(url_hint) or base
        for lp in ["/login", "/signin", "/signmein", "/"]:
            login_page_candidates.append(hint_base + lp)
    for lp in _LOGIN_PAGE_PATHS:
        login_page_candidates.append(base + lp)
    # Deduplicate while preserving order
    seen_lp = set()
    login_page_candidates = [x for x in login_page_candidates if not (x in seen_lp or seen_lp.add(x))]

    login_html = ""
    login_url_used = ""
    for candidate in login_page_candidates:
        try:
            r = sess.get(candidate, timeout=12, verify=False)
            txt_lo = r.text.lower()
            if r.status_code == 200 and (
                "password" in txt_lo or "username" in txt_lo or "login" in txt_lo
            ):
                login_html = r.text
                login_url_used = candidate
                print(f"[{pid}] Login page found: {candidate}")
                break
        except Exception:
            continue

    if not login_html:
        print(f"[{pid}] ❌ Login page not found at {base}")
        return None, None, None, None

    # ── Step 2: Build post data ──────────────────────────────────────────────
    user_field, pass_field = _univ_detect_form_fields(login_html)
    post_data: dict = {user_field: username, pass_field: password}
    print(f"[{pid}] Form fields: {user_field}={username}, {pass_field}=***")

    # Math captcha
    m_cap = re.search(r"[Ww]hat is (\d+) \+ (\d+)", login_html)
    if m_cap:
        ans = int(m_cap.group(1)) + int(m_cap.group(2))
        post_data["capt"] = ans
        print(f"[{pid}] Math captcha: {m_cap.group(1)}+{m_cap.group(2)}={ans}")

    # Collect ALL hidden fields from the login form (CSRF tokens, session seeds, etc.)
    for hf in re.finditer(
        r'<input[^>]+type=["\']hidden["\'][^>]*name=["\']([^"\']+)["\'][^>]*value=["\']([^"\']*)["\']'
        r'|<input[^>]+name=["\']([^"\']+)["\'][^>]*type=["\']hidden["\'][^>]*value=["\']([^"\']*)["\']',
        login_html, re.IGNORECASE,
    ):
        n = (hf.group(1) or hf.group(3) or "").strip()
        v = (hf.group(2) or hf.group(4) or "").strip()
        if n and n.lower() not in (user_field.lower(), pass_field.lower()):
            post_data[n] = v

    # ── Step 3: Try signin paths ─────────────────────────────────────────────
    # Build candidate signin URLs: derive from login page URL first, then fallbacks
    login_dir = re.sub(r"/[^/]+$", "", login_url_used)  # directory of login page
    signin_candidates = []
    for sp in ["/signin", "/signmein", "/login"]:
        signin_candidates.append(login_dir + sp)   # same directory
    for sp in _SIGNIN_PATHS:
        signin_candidates.append(base + sp)
    # Deduplicate
    seen_sp = set()
    signin_candidates = [x for x in signin_candidates if not (x in seen_sp or seen_sp.add(x))]

    logged_sess = None
    login_resp_text = ""
    for sp_url in signin_candidates:
        try:
            r2 = sess.post(
                sp_url, data=post_data, timeout=12, allow_redirects=True, verify=False,
                headers={"Referer": login_url_used},
            )
            if r2.status_code in (200, 201, 302) and not _univ_is_login_page(r2.url, r2.text):
                logged_sess = sess
                login_resp_text = r2.text
                print(f"[{pid}] ✅ Signed in via {sp_url} → {r2.url}")
                break
        except Exception:
            continue

    if not logged_sess:
        print(f"[{pid}] ❌ Login failed — all signin paths exhausted")
        return None, None, None, None

    # ── Step 3b: Validate session by probing the original URL (soft check) ──────
    # Only used as a hint — never hard-fails a successful login POST
    if url_hint:
        try:
            chk = logged_sess.get(url_hint, timeout=10, verify=False, allow_redirects=True)
            if _univ_is_login_page(chk.url, chk.text):
                print(f"[{pid}] ⚠️ Hint URL looks like login page after signin — proceeding anyway")
            elif len(chk.text) < 50:
                print(f"[{pid}] ⚠️ Hint page very short ({len(chk.text)}b) — proceeding anyway")
            else:
                print(f"[{pid}] ✅ Session validated via {url_hint} ({len(chk.text)}b)")
                login_resp_text = login_resp_text or chk.text
        except Exception as e:
            print(f"[{pid}] ⚠️ Session validation skipped: {e}")

    # ── Step 4: Extract session token ────────────────────────────────────────
    token = _univ_extract_token(login_resp_text)
    if not token:
        # Probe dashboard pages — include hint URL itself
        dash_candidates = []
        if url_hint:
            dash_candidates.append(url_hint)
        hint_base2 = _extract_panel_base_url(url_hint) if url_hint else base
        for dp in _DASHBOARD_PATHS:
            if hint_base2 and hint_base2 != base:
                dash_candidates.append(hint_base2 + dp)
            dash_candidates.append(base + dp)
        for dash_url in dash_candidates:
            try:
                rd = logged_sess.get(dash_url, timeout=10, verify=False)
                token = _univ_extract_token(rd.text)
                if token:
                    print(f"[{pid}] Token found via {dash_url}")
                    break
            except Exception:
                continue
    print(f"[{pid}] Token: {'found (' + token[:8] + '...)' if token else 'empty (cookie-based)'}")

    # ── Step 5: Probe data endpoints ─────────────────────────────────────────
    today = time.strftime("%Y-%m-%d")
    hint_base3 = _extract_panel_base_url(url_hint) if url_hint else None

    # Step 5a: Scrape the dashboard/hint page HTML to extract AJAX data URLs
    # Many panels embed the data URL directly in their JS (ajax: "/path/to/data.php")
    scraped_ep_candidates = []
    pages_to_scrape = []
    if url_hint:
        pages_to_scrape.append(url_hint)
    for dp in _DASHBOARD_PATHS:
        if hint_base3 and hint_base3 != base:
            pages_to_scrape.append(hint_base3 + dp)
        pages_to_scrape.append(base + dp)
    for scrape_url in pages_to_scrape[:8]:  # limit to avoid slow startup
        try:
            rp = logged_sess.get(scrape_url, timeout=10, verify=False)
            if rp.status_code != 200:
                continue
            pg = rp.text
            # Look for ajax/url patterns pointing to data PHP files
            for m in re.finditer(
                r'''["']([^"']*(?:data_sms|dt_reports|dt_sms|cdr|reports)[^"']*\.php)['"''',
                pg, re.IGNORECASE
            ):
                raw = m.group(1)
                # Convert to absolute URL
                if raw.startswith("http"):
                    abs_ep = raw
                elif raw.startswith("/"):
                    parsed_host = re.match(r"(https?://[^/]+)", scrape_url)
                    abs_ep = (parsed_host.group(1) if parsed_host else base) + raw
                else:
                    abs_ep = base + "/" + raw.lstrip("/")
                style = "xisora" if "dt_reports" in raw or "dt_sms" in raw else "ints"
                eng = "xisora" if style == "xisora" else "ints_smscdr"
                scraped_ep_candidates.append((abs_ep, raw, style, eng))
                print(f"[{pid}] 🔎 Scraped data URL from {scrape_url}: {raw}")
        except Exception:
            continue

    # Step 5b: Build known-path candidates (hint_base first, then base)
    known_ep_candidates = []
    for ep_path, style, eng_name in _DATA_ENDPOINTS:
        if hint_base3 and hint_base3 != base:
            known_ep_candidates.append((hint_base3 + ep_path, ep_path, style, eng_name))
        known_ep_candidates.append((base + ep_path, ep_path, style, eng_name))

    # Combine: scraped first (highest confidence), then known paths
    all_ep_candidates = scraped_ep_candidates + known_ep_candidates

    for full_ep, ep_path, style, eng_name in all_ep_candidates:
        try:
            test_url = _univ_build_url(full_ep, token, today, style)
            rr = logged_sess.get(
                test_url, timeout=10, verify=False,
                headers={"X-Requested-With": "XMLHttpRequest"},
            )
            body = rr.text.strip()
            print(f"[{pid}] Probe {full_ep} → HTTP {rr.status_code}, body={len(body)}b, starts={body[:30]!r}")
            if rr.status_code == 200 and body and not body.startswith("<"):
                try:
                    data = json.loads(body)
                    if "aaData" in data:
                        resolved_path = ep_path if full_ep.startswith(base) else ("/" + ep_path.lstrip("/"))
                        print(f"[{pid}] ✅ Data endpoint: {full_ep} (engine={eng_name})")
                        if hint_base3 and hint_base3 != base and full_ep.startswith(hint_base3):
                            panel["base_url"] = hint_base3
                        return logged_sess, token, eng_name, resolved_path
                except Exception:
                    pass
        except Exception as probe_err:
            print(f"[{pid}] Probe error {full_ep}: {probe_err}")
            continue

    # ── Step 5c: HTML table scraping fallback ────────────────────────────────
    # For panels that render data directly in HTML (no AJAX JSON endpoint)
    html_scrape_url = None
    html_pages_to_try = []
    if url_hint:
        html_pages_to_try.append(url_hint)
    for dp in _DASHBOARD_PATHS:
        if hint_base3 and hint_base3 != base:
            html_pages_to_try.append(hint_base3 + dp)
        html_pages_to_try.append(base + dp)
    for pg_url in html_pages_to_try[:6]:
        try:
            rp = logged_sess.get(pg_url, timeout=10, verify=False)
            if rp.status_code != 200 or _univ_is_login_page(rp.url, rp.text):
                continue
            from bs4 import BeautifulSoup
            soup = BeautifulSoup(rp.text, "lxml")
            tables = soup.find_all("table")
            for tbl in tables:
                rows = tbl.find_all("tr")
                if len(rows) >= 3:  # at least header + 2 data rows
                    print(f"[{pid}] 🔎 HTML scraping fallback: found table with {len(rows)} rows at {pg_url}")
                    html_scrape_url = pg_url
                    break
            if html_scrape_url:
                break
        except Exception:
            continue
    if html_scrape_url:
        return logged_sess, token, "html_scrape", html_scrape_url

    # Login succeeded but no endpoint matched → return with INTS default
    print(f"[{pid}] ⚠️ Login OK but no data endpoint found — using default")
    return logged_sess, token, "ints_smscdr", "/agent/res/data_smscdr.php"


def _universal_fetch(panel):
    """Fetch OTPs from any panel using the universal engine."""
    pid = panel["id"]
    base = panel["base_url"].rstrip("/")
    engine = panel.get("engine", "ints_smscdr")
    # IVA SMS has its own fetcher
    if engine == "iva_sms":
        return _iva_fetch(panel)
    # API Key panels have their own fetcher
    if engine == "api_key":
        return _api_key_fetch(panel)
    data_path = panel.get("data_path", "/agent/res/data_smscdr.php")
    style = "xisora" if engine == "xisora" else "ints"
    found = {}

    with _get_dp_lock(pid):
        sd = _dynamic_sessions.get(pid, {})
        if not sd.get("session"):
            sess, tok, det_eng, det_path = _universal_login(panel)
            if not sess:
                _record_error(pid)
                return found
            if det_eng and (det_eng != engine or det_path != data_path):
                panel["engine"] = det_eng
                panel["data_path"] = det_path
                engine = det_eng
                data_path = det_path
                style = "xisora" if engine == "xisora" else "ints"
                save_dynamic_panels()
            _dynamic_sessions[pid] = {"session": sess, "token": tok}
            sd = _dynamic_sessions[pid]

        sess = sd["session"]
        token = sd.get("token", "")
        today = time.strftime("%Y-%m-%d")
        full_ep = base + data_path
        hdrs = {"X-Requested-With": "XMLHttpRequest"}

        def _do_get():
            return sess.get(
                _univ_build_url(full_ep, token, today, style),
                headers=hdrs, timeout=15, verify=False,
            )

        # ── HTML scraping engine ──────────────────────────────────────────────
        if engine == "html_scrape":
            page_url = data_path  # data_path is the full page URL for this engine
            try:
                rp = sess.get(page_url, timeout=15, verify=False)
                if rp.status_code != 200 or _univ_is_login_page(rp.url, rp.text):
                    print(f"[{pid}] Session expired (html_scrape) — re-login")
                    _dynamic_sessions[pid] = {}
                    sess2, tok2, det_eng, det_path = _universal_login(panel)
                    if not sess2:
                        _record_error(pid)
                        return found
                    panel["engine"] = det_eng
                    panel["data_path"] = det_path
                    engine = det_eng
                    data_path = det_path
                    save_dynamic_panels()
                    _dynamic_sessions[pid] = {"session": sess2, "token": tok2}
                    rp = sess2.get(det_path if det_eng == "html_scrape" else page_url,
                                   timeout=15, verify=False)
                from bs4 import BeautifulSoup
                soup = BeautifulSoup(rp.text, "lxml")
                # Find header row to map column positions
                tbl = None
                for t in soup.find_all("table"):
                    if len(t.find_all("tr")) >= 3:
                        tbl = t
                        break
                if not tbl:
                    _record_fetch(pid, 0)
                    return found
                headers = [th.get_text(strip=True).lower()
                           for th in (tbl.find("tr").find_all(["th", "td"]))]
                # Heuristic column detection
                num_col = next((i for i, h in enumerate(headers)
                                if any(w in h for w in ("number", "msisdn", "phone", "mobile", "num"))), None)
                txt_col = next((i for i, h in enumerate(headers)
                                if any(w in h for w in ("message", "sms", "text", "body", "msg"))), None)
                svc_col = next((i for i, h in enumerate(headers)
                                if any(w in h for w in ("service", "route", "dest", "sender"))), None)
                row_count = 0
                for tr in tbl.find_all("tr")[1:]:  # skip header
                    cells = tr.find_all(["td", "th"])
                    if not cells:
                        continue
                    row_count += 1
                    number = cells[num_col].get_text(strip=True) if num_col is not None and num_col < len(cells) else ""
                    sms_txt = cells[txt_col].get_text(strip=True) if txt_col is not None and txt_col < len(cells) else ""
                    service = cells[svc_col].get_text(strip=True) if svc_col is not None and svc_col < len(cells) else ""
                    # Fallback: scan all cells for phone-like pattern and OTP
                    if not number:
                        for c in cells:
                            ct = c.get_text(strip=True)
                            if re.match(r"^\+?\d{7,15}$", ct):
                                number = ct
                                break
                    if not sms_txt:
                        for c in cells:
                            ct = c.get_text(strip=True)
                            if len(ct) > 10 and re.search(r"\d{4,8}", ct):
                                sms_txt = ct
                                break
                    otp = extract_otp_from_sms(sms_txt)
                    if number and otp:
                        key = f"{number}:{sms_txt}"
                        found[key] = (number, otp, sms_txt, service)
                _record_fetch(pid, row_count)
                if found:
                    print(f"[{pid}] ✅ HTML-scraped {row_count} rows, {len(found)} OTPs")
            except Exception as e:
                print(f"[{pid}] HTML scrape error: {e}")
                _record_error(pid)
                _dynamic_sessions[pid] = {}
            return found

        # ── JSON / AJAX engine (default) ──────────────────────────────────────
        try:
            r = _do_get()
            body = r.text.strip()
            if r.status_code != 200 or not body or body.startswith("<") or "Direct Script" in body:
                print(f"[{pid}] Session expired — re-login")
                _dynamic_sessions[pid] = {}
                sess2, tok2, det_eng, det_path = _universal_login(panel)
                if not sess2:
                    _record_error(pid)
                    return found
                if det_eng:
                    panel["engine"] = det_eng
                    panel["data_path"] = det_path
                    engine = det_eng
                    data_path = det_path
                    style = "xisora" if engine == "xisora" else "ints"
                    full_ep = base + data_path
                    save_dynamic_panels()
                _dynamic_sessions[pid] = {"session": sess2, "token": tok2}
                sd = _dynamic_sessions[pid]
                sess = sess2
                token = tok2
                r = _do_get()
                body = r.text.strip()

            rows = json.loads(body).get("aaData", [])
            for row in rows:
                parsed = _univ_parse_row(row, engine)
                if not parsed:
                    continue
                number, service, sms_txt = parsed
                otp = extract_otp_from_sms(sms_txt)
                if otp:
                    key = f"{number}:{sms_txt}"
                    found[key] = (number, otp, sms_txt, service)
            _record_fetch(pid, len(rows))
            if found:
                print(f"[{pid}] ✅ Fetched {len(rows)} rows, {len(found)} OTPs")
        except Exception as e:
            print(f"[{pid}] Fetch error: {e}")
            _record_error(pid)
            _dynamic_sessions[pid] = {}
    return found


def _univ_parse_row(row, engine):
    """Parse one aaData row. Returns (number, service, sms_text) or None."""
    try:
        if not row:
            return None
        cells = [str(c).strip() for c in row]
        # Standard INTS layout: [date, range, number, cli/service, client, sms, ...]
        # extract_otp_from_sms already enforces real-SMS validation (≥4 alpha chars)
        if len(cells) > 5:
            number  = cells[2]
            service = cells[3]
            sms_txt = cells[5]
            if number and extract_otp_from_sms(sms_txt):
                return number, service, sms_txt
        # Shorter rows (some panels only have 5 cols)
        if len(cells) > 4:
            number  = cells[2]
            service = cells[3]
            sms_txt = cells[4]
            if number and extract_otp_from_sms(sms_txt):
                return number, service, sms_txt
    except Exception:
        pass
    return None


# ── API Key Panel Engine ───────────────────────────────────────────────────────

_API_KEY_ENDPOINTS = [
    "/api/sms",
    "/api/messages",
    "/api/received",
    "/api/sms/received",
    "/api/v1/sms",
    "/api/v1/messages",
    "/api/inbox",
    "/api/sms/list",
    "/api/data",
    "/sms",
    "/api/otp",
]

_API_KEY_PARAMS = ["api_key", "token", "key", "apikey", "access_token"]


def _api_key_test(base_url, api_key):
    """Try common API endpoints with key. Returns (path, param_style) or (None, None).
    param_style: 'bearer', 'api_key=xxx', 'token=xxx', etc.
    """
    base = base_url.rstrip("/")
    hdrs = {"Accept": "application/json", "User-Agent": _UA}
    for path in _API_KEY_ENDPOINTS:
        # Try query param variants
        for param in _API_KEY_PARAMS:
            try:
                r = requests.get(
                    f"{base}{path}?{param}={api_key}",
                    timeout=8, verify=False, headers=hdrs,
                )
                if r.status_code == 200:
                    try:
                        data = r.json()
                        if isinstance(data, list):
                            print(f"[APIKEY-TEST] ✅ {path}?{param}= → array")
                            return path, param
                        if isinstance(data, dict):
                            for dk in ("data", "sms", "messages", "records", "result", "items", "list"):
                                if dk in data and isinstance(data[dk], list):
                                    print(f"[APIKEY-TEST] ✅ {path}?{param}= → dict.{dk}")
                                    return path, f"{param}|resp={dk}"
                            if any(k in data for k in ("status", "success", "ok", "error")):
                                print(f"[APIKEY-TEST] ✅ {path}?{param}= → status obj")
                                return path, param
                    except Exception:
                        pass
            except Exception:
                continue
        # Try Authorization: Bearer header
        try:
            r = requests.get(
                f"{base}{path}", timeout=8, verify=False,
                headers={**hdrs, "Authorization": f"Bearer {api_key}"},
            )
            if r.status_code == 200:
                try:
                    data = r.json()
                    if isinstance(data, (list, dict)):
                        print(f"[APIKEY-TEST] ✅ {path} → Bearer header")
                        return path, "bearer"
                except Exception:
                    pass
        except Exception:
            continue
    return None, None


def _api_key_fetch(panel):
    """Fetch OTPs from an API-key authenticated panel."""
    pid      = panel["id"]
    base     = panel["base_url"].rstrip("/")
    api_key  = panel.get("api_key", "")
    data_path = panel.get("data_path", "/api/sms")
    param_style = panel.get("api_key_param", "api_key")
    found = {}
    try:
        hdrs = {"Accept": "application/json", "User-Agent": _UA}
        if param_style == "bearer":
            url = f"{base}{data_path}"
            hdrs["Authorization"] = f"Bearer {api_key}"
        elif "|resp=" in param_style:
            pname = param_style.split("|resp=")[0]
            url = f"{base}{data_path}?{pname}={api_key}"
        else:
            url = f"{base}{data_path}?{param_style}={api_key}"

        r = requests.get(url, timeout=15, verify=False, headers=hdrs)
        if r.status_code != 200:
            print(f"[{pid}] API key fetch HTTP {r.status_code}")
            _record_error(pid)
            return found

        raw = r.json()
        rows = []
        if isinstance(raw, list):
            rows = raw
        elif isinstance(raw, dict):
            resp_key = param_style.split("|resp=")[1] if "|resp=" in param_style else None
            if resp_key and resp_key in raw:
                rows = raw[resp_key]
            else:
                for dk in ("data", "otps", "sms", "messages", "records", "result", "items", "list"):
                    val = raw.get(dk)
                    if isinstance(val, list):
                        rows = val
                        break
                    elif isinstance(val, dict):
                        # Handle nested format e.g. fastxotps: {data: {otps: [...]}}
                        for inner in ("otps", "data", "sms", "messages", "records", "result", "items", "list"):
                            inner_val = val.get(inner)
                            if isinstance(inner_val, list):
                                rows = inner_val
                                break
                        if rows:
                            break

        for row in rows:
            number = sms_txt = service = ""
            if isinstance(row, dict):
                number  = str(row.get("number") or row.get("phone") or row.get("msisdn") or
                              row.get("from") or row.get("sender") or row.get("to") or "").strip()
                sms_txt = str(row.get("message") or row.get("sms") or row.get("text") or
                              row.get("body") or row.get("content") or "").strip()
                service = str(row.get("service") or row.get("sender") or row.get("source") or "").strip()
            elif isinstance(row, list):
                if len(row) > 2: number  = str(row[2]).strip()
                if len(row) > 5: sms_txt = str(row[5]).strip()
                elif len(row) > 3: sms_txt = str(row[3]).strip()
                if len(row) > 3: service = str(row[3]).strip()
            # Auto-detect service from SMS content when panel doesn't provide it
            if not service and sms_txt:
                service = _detect_service_from_sms(sms_txt)
            otp = extract_otp_from_sms(sms_txt)
            if number and otp:
                key = f"{number}:{sms_txt}"
                found[key] = (number, otp, sms_txt, service)

        _record_fetch(pid, len(rows))
        if found:
            print(f"[{pid}] ✅ API key: {len(rows)} rows, {len(found)} OTPs")
    except Exception as e:
        print(f"[{pid}] API key fetch error: {e}")
        _record_error(pid)
    return found


# ─────────────────────────────────────────────────────────────────────────────

# ── FastXOTPs (fastxotps.com) helpers ─────────────────────────────────────────
FASTX_BASE = "https://fastxotps.com"
FASTX_API_KEY = "MURAD_979BB07726A593010D1BA4A2"


# Keywords to detect synthetic services from OTP message content
_FASTX_MSG_SERVICE_KEYWORDS = {
    "INSTAGRAM": ["instagram"],
    "TIKTOK":    ["tiktok", "tik tok"],
    "SNAPCHAT":  ["snapchat"],
    "TWITTER":   ["twitter", "x.com"],
    "GOOGLE":    ["google"],
    "YOUTUBE":   ["youtube"],
    "LINKEDIN":  ["linkedin"],
    "AMAZON":    ["amazon"],
}


def _fastx_detect_extra_services(base_services):
    """
    Inject synthetic services (INSTAGRAM, etc.) derived from parent carrier ranges.
    INSTAGRAM inherits all FACEBOOK ranges since they share the same carrier routes.
    Also scans recent OTPs for any extra ranges not already covered.
    Returns the enriched services list.
    """
    # Map synthetic service -> which liveaccess services share its carrier routes
    _carrier_parents = {
        "INSTAGRAM": {"FACEBOOK", "FB"},
        "TIKTOK":    {"FACEBOOK", "TWILIO"},
        "SNAPCHAT":  {"TWILIO"},
    }
    # Preferred insertion position (insert after parent)
    _insert_after = {
        "INSTAGRAM": {"FACEBOOK", "FB"},
        "TIKTOK":    {"FACEBOOK", "TWILIO"},
        "SNAPCHAT":  {"TWILIO"},
    }

    existing_sids = {s.get("sid", "").upper() for s in base_services}
    now_ms = int(time.time() * 1000)

    # Build a flat list of (prefix, range_str) from all base service ranges
    all_prefixes = []
    for svc in base_services:
        for rng in svc.get("ranges", []):
            prefix = rng.rstrip("X")
            all_prefixes.append((prefix, rng))
    all_prefixes.sort(key=lambda x: -len(x[0]))

    # Step 1: seed synthetic ranges from parent carrier's liveaccess ranges
    synthetic: dict = {}  # sid -> {ranges: set, last_at: int}
    for sid, parents in _carrier_parents.items():
        if sid.upper() in existing_sids:
            continue  # already in liveaccess, handle via merge below
        for svc in base_services:
            if svc.get("sid", "").upper() in parents:
                parent_ranges = set(svc.get("ranges", []))
                if parent_ranges:
                    entry = synthetic.setdefault(sid, {"ranges": set(), "last_at": 0})
                    entry["ranges"] |= parent_ranges
                    ts = svc.get("last_at", now_ms)
                    if ts > entry["last_at"]:
                        entry["last_at"] = ts

    # Step 2: also scan recent OTPs for any extra ranges not covered by parent
    try:
        r = requests.get(f"{FASTX_BASE}/api/otps",
                         params={"api_key": FASTX_API_KEY},
                         timeout=10, verify=False)
        if r.status_code == 200:
            d = r.json()
            otps = d.get("data", {}).get("otps", d.get("otps", []))
            if isinstance(otps, list):
                for otp in otps:
                    msg = (otp.get("message") or "").lower()
                    num = str(otp.get("number") or "").strip().lstrip("+")
                    if not num:
                        continue
                    ts = otp.get("time", now_ms)
                    for sid, keywords in _FASTX_MSG_SERVICE_KEYWORDS.items():
                        if not any(kw in msg for kw in keywords):
                            continue
                        matched_rng = None
                        for prefix, rng in all_prefixes:
                            if num.startswith(prefix):
                                matched_rng = rng
                                break
                        if not matched_rng and len(num) >= 7:
                            plen = max(4, min(9, len(num) - 6))
                            matched_rng = num[:plen] + "XXX"
                        if matched_rng:
                            entry = synthetic.setdefault(sid, {"ranges": set(), "last_at": 0})
                            entry["ranges"].add(matched_rng)
                            if ts > entry["last_at"]:
                                entry["last_at"] = ts
    except Exception as e:
        print(f"[FASTX] detect_extra_services OTP scan error: {e}")

    if not synthetic:
        return base_services

    enriched = list(base_services)
    for sid, info in synthetic.items():
        sid_upper = sid.upper()
        merged = False
        for svc in enriched:
            if svc.get("sid", "").upper() == sid_upper:
                existing_ranges = set(svc.get("ranges", []))
                new_ranges = info["ranges"] - existing_ranges
                if new_ranges:
                    svc["ranges"] = sorted(existing_ranges | info["ranges"])
                    print(f"[FASTX] Merged {len(new_ranges)} range(s) into existing {sid!r}: {sorted(new_ranges)}")
                merged = True
                break
        if not merged:
            new_svc = {
                "sid": sid,
                "ranges": sorted(info["ranges"]),
                "last_at": info["last_at"],
            }
            parents = _insert_after.get(sid, set())
            insert_idx = len(enriched)
            for i, svc in enumerate(enriched):
                if svc.get("sid", "").upper() in parents:
                    insert_idx = i + 1
                    break
            enriched.insert(insert_idx, new_svc)
            print(f"[FASTX] Injected synthetic {sid!r} with {len(new_svc['ranges'])} ranges: {new_svc['ranges']}")

    return enriched


def _fastx_liveaccess():
    """Return list of live services [{sid, ranges, last_at}] from fastxotps,
    enriched with synthetic services (INSTAGRAM, etc.) detected from recent OTPs."""
    try:
        r = requests.get(f"{FASTX_BASE}/api/liveaccess",
                         params={"api_key": FASTX_API_KEY},
                         timeout=15, verify=False)
        if r.status_code == 200:
            d = r.json()
            if d.get("status") == "ok":
                base = d.get("services", [])
                return _fastx_detect_extra_services(base)
        print(f"[FASTX] liveaccess HTTP {r.status_code}")
    except Exception as e:
        print(f"[FASTX] liveaccess error: {e}")
    return []


def _fastx_getnum(prefix):
    """Allocate a number for the given range prefix. Returns full_number str or None."""
    try:
        r = requests.post(f"{FASTX_BASE}/api/getnum",
                          params={"api_key": FASTX_API_KEY},
                          data={"prefix": prefix},
                          timeout=15, verify=False)
        if r.status_code == 200:
            resp = r.json()
            d = resp.get("data") or {}
            num = d.get("full_number") or d.get("no_plus_number")
            if num:
                return num if num.startswith("+") else "+" + num
            print(f"[FASTX] getnum no number in response: {resp}")
        else:
            print(f"[FASTX] getnum HTTP {r.status_code}: {r.text[:200]}")
    except Exception as e:
        print(f"[FASTX] getnum error: {e}")
    return None


def _v2_svc_emoji(sid):
    m = {"FACEBOOK": "🔵", "INSTAGRAM": "📸", "WHATSAPP": "💚", "TELEGRAM": "✈️",
         "TWITTER": "🐦", "TIKTOK": "🎵", "BINANCE": "🟡", "SNAPCHAT": "👻",
         "GOOGLE": "🔴", "YOUTUBE": "📺", "LINKEDIN": "💼", "AMAZON": "🛒"}
    return m.get((sid or "").upper(), "📱")


def _v2_build_console_markup(services):
    """Build inline keyboard for V2 console service list."""
    markup = types.InlineKeyboardMarkup(row_width=2)
    btns = []
    for svc in services:
        sid = svc.get("sid", "?")
        cnt = len(svc.get("ranges", []))
        if cnt:
            btns.append(types.InlineKeyboardButton(
                f"{_v2_svc_emoji(sid)} {sid} ({cnt})",
                callback_data=f"v2svc:{sid}"
            ))
    if btns:
        markup.add(*btns)
    return markup, bool(btns)


def _v2_custom_range_step(message):
    """Handle user input for CUSTOM RANGE — find matching range and allocate number."""
    uid = message.from_user.id
    txt = (message.text or "").strip()

    if _intercept_menu_btn(message):
        return
    if txt in ("❌ Cancel", "❌ cancel"):
        bot.send_message(
            message.chat.id,
            "❌ বাদ দেওয়া হয়েছে।",
            reply_markup=v2_switch_menu(),
            parse_mode="HTML",
        )
        return

    prefix = re.sub(r"[^\d]", "", txt)
    if not prefix:
        msg = bot.send_message(
            message.chat.id,
            "❌ <b>Invalid input!</b> শুধু সংখ্যা দাও (e.g. <code>8801</code>)\n\nআবার চেষ্টা করো:",
            reply_markup=types.ReplyKeyboardMarkup(resize_keyboard=True).add("❌ Cancel"),
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, _v2_custom_range_step)
        return

    bot.send_message(message.chat.id, f"⏳ <b>{prefix}</b> দিয়ে number খুঁজছি...", parse_mode="HTML")

    services = _fastx_liveaccess()
    matched_prefix = None
    matched_sid = None
    for svc in services:
        sid = svc.get("sid", "?")
        for rng in svc.get("ranges", []):
            clean = rng.rstrip("X")
            if clean.startswith(prefix) or prefix.startswith(clean):
                matched_prefix = clean
                matched_sid = sid
                break
        if matched_prefix:
            break

    if not matched_prefix:
        msg2 = bot.send_message(
            message.chat.id,
            f"❌ <b>Range পাওয়া যায়নি!</b>\n\n"
            f"<code>{prefix}</code> দিয়ে এখন কোনো live range নেই।\n\n"
            f"অন্য prefix দাও অথবা <b>LIVE RANGE</b> দেখো:",
            reply_markup=types.ReplyKeyboardMarkup(resize_keyboard=True).add("❌ Cancel"),
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg2, _v2_custom_range_step)
        return

    num = _fastx_getnum(matched_prefix)
    if not num:
        bot.send_message(
            message.chat.id,
            f"❌ <b>Number allocate হলো না!</b>\n"
            f"Range <code>{matched_prefix}</code> এ এখন number নেই। পরে চেষ্টা করো।",
            reply_markup=v2_switch_menu(),
            parse_mode="HTML",
        )
        return

    with user_map_lock:
        old_nums = [k for k, v in user_map.items() if v == uid]
        for old_clean in old_nums:
            user_map.pop(old_clean, None)
            assigned_time.pop(old_clean, None)
    if old_nums:
        _save_user_map()
    register_number(uid, num)
    c_name, flag = get_country_details(num)
    display_num = num if num.startswith("+") else "+" + num
    _user_last_svc[uid] = (matched_sid.lower(), c_name)

    refresh_kb = types.InlineKeyboardMarkup(row_width=2)
    refresh_kb.add(
        types.InlineKeyboardButton("🔄 New Number", callback_data=f"v2rng:{matched_prefix}:{matched_sid}"),
    )
    if get_otp_group_link():
        refresh_kb.add(types.InlineKeyboardButton("📢 OTP Group", url=get_otp_group_link()))

    bot.send_message(
        message.chat.id,
        f"✅ <b>CUSTOM RANGE — Number পেয়েছ!</b>\n\n"
        f"📱 <b>Service:</b> {_v2_svc_emoji(matched_sid)} <b>{matched_sid}</b>\n"
        f"🔢 <b>Range:</b> <code>{matched_prefix}XXXX</code>\n"
        f"📞 <b>Number:</b> <code>{display_num}</code>\n"
        f"🌍 <b>Country:</b> {flag} {c_name}\n\n"
        f"⏳ OTP আসলে এখানে পাঠানো হবে।",
        reply_markup=refresh_kb,
        parse_mode="HTML",
    )
    bot.send_message(message.chat.id, "🔄 V2 SWITCH menu:", reply_markup=v2_switch_menu(), parse_mode="HTML")


def _v2_show_console(chat_id):
    """Send V2 Console service list to chat_id."""
    services = _fastx_liveaccess()
    markup, has_btns = _v2_build_console_markup(services)
    if not has_btns:
        bot.send_message(chat_id,
                         "❌ <b>V2 Console:</b> কোনো live service নেই এখন। পরে চেষ্টা করো।",
                         parse_mode="HTML")
        return
    bot.send_message(
        chat_id,
        "📡 <b>V2 LIVE CONSOLE</b>\n"
        "⚡━━━━━━━━━━━━━━⚡\n\n"
        "🔴 <b>Live service select koro:</b>\n"
        "<i>এই service-গুলোতে এখন OTP আসছে</i>\n\n"
        "⚡━━━━━━━━━━━━━━⚡",
        reply_markup=markup,
        parse_mode="HTML"
    )


# ──────────────────────────────────────────────────────────────────────────────

def _start_dynamic_panel(panel):
    pid = panel["id"]
    with _stats_lock:
        _panel_stats[pid] = {
            "name": panel.get("username", pid),
            "host": panel.get("host", ""),
            "status": "⏳",
            "count": 0,
            "last": None,
            "errors": 0,
        }

    def monitor():
        global seen_otps
        print(f"[{pid}-MONITOR] Started. Pre-loading existing records...")
        existing = _universal_fetch(panel)
        with seen_lock:
            for key in existing:
                seen_otps[key] = True
            save_json(SEEN_FILE, seen_otps)
        print(f"[{pid}-MONITOR] Pre-loaded {len(existing)} records. Watching for new ones...")
        while True:
            try:
                process_new_otps(_universal_fetch(panel))
            except Exception as e:
                print(f"[{pid}-MONITOR] Loop error: {e}")
            time.sleep(POLL_INTERVAL)

    threading.Thread(target=monitor, daemon=True).start()


# ─── IVA SMS (ivasms.com) engine ─────────────────────────────────────────────

_iva_scrapers: dict = {}
_iva_lock_map: dict = {}


def _iva_get_lock(pid):
    if pid not in _iva_lock_map:
        _iva_lock_map[pid] = threading.Lock()
    return _iva_lock_map[pid]


_IVA_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.5",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection": "keep-alive",
    "Upgrade-Insecure-Requests": "1",
}


def _iva_make_session(cookie_str=""):
    """Create a plain requests.Session with browser headers + optional cookies."""
    sess = requests.Session()
    sess.headers.update(_IVA_HEADERS)
    if cookie_str:
        for part in cookie_str.split(";"):
            part = part.strip()
            if "=" in part:
                k, v = part.split("=", 1)
                sess.cookies.set(k.strip(), v.strip(), domain="ivasms.com")
    return sess


# Keep old name for compatibility
def _iva_make_scraper():
    return _iva_make_session()


def _iva_set_cookies(sess, cookie_str):
    for part in cookie_str.split(";"):
        part = part.strip()
        if "=" in part:
            k, v = part.split("=", 1)
            sess.cookies.set(k.strip(), v.strip(), domain="ivasms.com")


def _iva_login(panel):
    """Login to ivasms.com using browser cookies (plain requests — no cloudscraper).
    Railway/cloud server IPs are blocked by Cloudflare for email/password login,
    so we ONLY support cookie-based login here.
    """
    pid     = panel["id"]
    cookies = panel.get("cookie_str", "")
    base    = "https://ivasms.com"

    if not cookies:
        return False

    sess = _iva_make_session(cookies)
    try:
        r = sess.get(
            f"{base}/portal/sms/received",
            timeout=20,
            allow_redirects=True,
        )
        if r.status_code == 200 and "login" not in r.url.lower():
            _iva_scrapers[pid] = sess
            print(f"[{pid}] ✅ IVA SMS: Cookie login OK ({r.url})")
            return True
        print(f"[{pid}] ❌ IVA SMS: Cookie login failed — status={r.status_code} url={r.url}")
        return False
    except Exception as e:
        print(f"[{pid}] ❌ IVA SMS cookie login error: {e}")
        return False


def _iva_parse_page(html):
    """Parse ivasms.com SMS received page → {key: (number, otp, sms_txt, service)}."""
    found = {}

    # 1) Embedded JS array (e.g. var data = [...])
    for pat in [
        r'(?:data|rows|messages|smsList|records|smsData)\s*[:=]\s*(\[.*?\])\s*[,;]',
        r'\.DataTable\([^)]*data\s*:\s*(\[.*?\])',
    ]:
        m = re.search(pat, html, re.DOTALL | re.I)
        if m:
            try:
                records = json.loads(m.group(1))
                for rec in records:
                    if not isinstance(rec, dict):
                        continue
                    num = str(rec.get("number", rec.get("phone", rec.get("msisdn", ""))))
                    txt = str(rec.get("message", rec.get("sms", rec.get("text", rec.get("body", "")))))
                    svc = str(rec.get("service", rec.get("cli", rec.get("sender", "IVA"))))
                    otp = extract_otp_from_sms(txt)
                    if num and otp:
                        found[f"{num}:{txt}"] = (num, otp, txt, svc)
                if found:
                    return found
            except Exception:
                pass

    # 2) HTML table: look for rows with phone number + SMS content
    from html.parser import HTMLParser
    class _TblParser(HTMLParser):
        def __init__(self):
            super().__init__()
            self.rows, self._cur_row, self._cur_cell, self._in_td = [], [], [], False
        def handle_starttag(self, tag, attrs):
            if tag == "tr":   self._cur_row = []
            elif tag == "td": self._in_td = True; self._cur_cell = []
        def handle_endtag(self, tag):
            if tag == "td":
                self._cur_row.append("".join(self._cur_cell).strip())
                self._in_td = False
            elif tag == "tr" and self._cur_row:
                self.rows.append(self._cur_row)
        def handle_data(self, data):
            if self._in_td: self._cur_cell.append(data)

    parser = _TblParser()
    try:
        parser.feed(html)
    except Exception:
        pass

    phone_re = re.compile(r"^\+?\d{7,15}$")
    for row in parser.rows:
        clean = [re.sub(r"\s+", " ", c).strip() for c in row]
        nums = [c for c in clean if phone_re.match(c)]
        smses = [c for c in clean if extract_otp_from_sms(c)]
        if nums and smses:
            n, t = nums[0], smses[0]
            found[f"{n}:{t}"] = (n, extract_otp_from_sms(t), t, "IVA")

    return found


def _iva_parse_rows(rows, default_svc="IVA"):
    """Convert a list of dicts (DataTables / JSON) to found-dict entries."""
    found = {}
    for rec in rows:
        if not isinstance(rec, dict):
            # DataTables may return list-of-lists
            continue
        num = str(rec.get("number", rec.get("phone", rec.get("msisdn",
              rec.get("sender", rec.get("from", ""))))))
        txt = str(rec.get("message", rec.get("sms", rec.get("text",
              rec.get("body", rec.get("content", ""))))))
        svc = str(rec.get("service", rec.get("cli", rec.get("application",
              rec.get("app", rec.get("shortcode", default_svc))))))
        otp = extract_otp_from_sms(txt)
        if num and otp:
            found[f"{num}:{txt}"] = (num, otp, txt, svc)
    return found


def _iva_dt_post(scraper, url, csrf_token, page_html="", start=0, length=100):
    """Send a DataTables server-side POST and return parsed rows."""
    today = time.strftime("%Y-%m-%d")
    payload = {
        "draw": "1",
        "start": str(start),
        "length": str(length),
        "search[value]": "",
        "search[regex]": "false",
        "_token": csrf_token,
        "start_date": today,
        "end_date": today,
        # Common DataTables column ordering params (harmless if unused)
        "order[0][column]": "0",
        "order[0][dir]": "desc",
    }
    hdrs = {
        "X-Requested-With": "XMLHttpRequest",
        "Accept": "application/json, text/javascript, */*",
        "Referer": "https://ivasms.com/portal/sms/received",
        "Content-Type": "application/x-www-form-urlencoded",
    }
    r = scraper.post(url, data=payload, headers=hdrs, timeout=25)
    if r.status_code != 200:
        return None, r.status_code
    ct = r.headers.get("Content-Type", "")
    if "json" not in ct:
        return None, 0
    try:
        return r.json(), 200
    except Exception:
        return None, 0


def _iva_fetch(panel):
    """Fetch latest OTPs from ivasms.com.

    Strategy (in order):
      1. Load /portal/sms/received, grab CSRF token + discover AJAX URL
      2. Try DataTables POST to the discovered/candidate AJAX endpoints
      3. Fall back to plain HTML table parse of the page
    """
    pid  = panel["id"]
    base = "https://ivasms.com"
    found = {}

    with _iva_get_lock(pid):
        scraper = _iva_scrapers.get(pid)
        if not scraper:
            if not _iva_login(panel):
                _record_error(pid)
                return found
            scraper = _iva_scrapers[pid]

        today = time.strftime("%Y-%m-%d")

        # ── Step 1: load the page ─────────────────────────────────────────────
        try:
            r = scraper.get(f"{base}/portal/sms/received",
                            params={"start_date": today, "end_date": today},
                            timeout=30)
        except Exception as e:
            print(f"[{pid}] IVA page load error: {e}")
            _iva_scrapers.pop(pid, None)
            _record_error(pid)
            return found

        # Session expired?
        if r.status_code != 200 or "login" in r.url.lower():
            print(f"[{pid}] IVA session expired → re-login")
            _iva_scrapers.pop(pid, None)
            if not _iva_login(panel):
                _record_error(pid)
                return found
            scraper = _iva_scrapers[pid]
            try:
                r = scraper.get(f"{base}/portal/sms/received",
                                params={"start_date": today, "end_date": today},
                                timeout=30)
            except Exception as e:
                print(f"[{pid}] IVA page reload error: {e}")
                _record_error(pid)
                return found

        html = r.text

        # ── Step 2: extract CSRF token from page ──────────────────────────────
        csrf = ""
        for pat in [
            r'<meta[^>]+name=["\']csrf-token["\'][^>]*content=["\']([^"\']+)["\']',
            r'"_token"\s*:\s*"([^"]+)"',
            r"_token[\"']\s*:\s*[\"']([^\"']+)[\"']",
            r'<input[^>]+name=["\']_token["\'][^>]*value=["\']([^"\']+)["\']',
        ]:
            m = re.search(pat, html, re.I)
            if m:
                csrf = m.group(1)
                break

        # ── Step 3: discover DataTables AJAX URL from page JS ─────────────────
        # Looks for patterns like: ajax: '/portal/sms/received/data'  or  url: '...'
        ajax_url_candidates = []
        for pat in [
            r"""ajax\s*:\s*['"](\/portal\/sms\/[^'"]+)['"]""",
            r"""url\s*:\s*['"](\/portal\/sms\/[^'"]+)['"]""",
            r"""action\s*=\s*['"](\/portal\/sms\/[^'"]+)['"]""",
            r"""fetch\(['"](\/portal\/sms\/[^'"?]+)""",
        ]:
            for m in re.finditer(pat, html, re.I):
                ep = m.group(1)
                if ep not in ajax_url_candidates:
                    ajax_url_candidates.append(ep)

        # Add hardcoded fallback candidates (most common Laravel SMS panel patterns)
        for fallback in [
            "/portal/sms/received/data",
            "/portal/sms/received-data",
            "/portal/received/sms/data",
            "/portal/sms/datatable",
            "/portal/sms/ajax",
            "/portal/sms/list",
            "/portal/api/sms/received",
            "/portal/sms/received",          # POST to same URL (common)
        ]:
            if fallback not in ajax_url_candidates:
                ajax_url_candidates.append(fallback)

        # ── Step 4: try DataTables POST on each candidate ─────────────────────
        for ep in ajax_url_candidates:
            ep_url = base + ep if ep.startswith("/") else ep
            try:
                js, code = _iva_dt_post(scraper, ep_url, csrf, html)
                if js is None:
                    continue
                # DataTables standard: {"draw":N, "data":[...], "recordsTotal":N}
                rows = (js if isinstance(js, list)
                        else js.get("data", js.get("records",
                             js.get("rows", js.get("sms", [])))))
                if isinstance(rows, list) and rows:
                    # Rows may be list-of-dicts or list-of-lists
                    if isinstance(rows[0], dict):
                        found = _iva_parse_rows(rows)
                    else:
                        # list-of-lists: try to figure out columns from header
                        # Best-effort: assume [id, number, message, service, ...]
                        for row in rows:
                            if len(row) >= 3:
                                num = str(row[1]) if len(row) > 1 else ""
                                txt = str(row[2]) if len(row) > 2 else ""
                                svc = str(row[3]) if len(row) > 3 else "IVA"
                                otp = extract_otp_from_sms(txt)
                                if num and otp:
                                    found[f"{num}:{txt}"] = (num, otp, txt, svc)
                    if found:
                        print(f"[{pid}] ✅ IVA DataTables hit: {ep_url} → {len(found)} OTPs")
                        # Remember this working endpoint for next time
                        if panel.get("iva_ajax_url") != ep_url:
                            panel["iva_ajax_url"] = ep_url
                            save_dynamic_panels()
                        break
            except Exception as ex:
                print(f"[{pid}] IVA AJAX probe {ep}: {ex}")
                continue

        # ── Step 5: fall back to static HTML table parse ─────────────────────
        if not found:
            found = _iva_parse_page(html)
            if found:
                print(f"[{pid}] ✅ IVA HTML-table parse → {len(found)} OTPs")

        _record_fetch(pid, len(found))
        if found:
            print(f"[{pid}] ✅ IVA SMS total: {len(found)} OTPs found")
        else:
            # Debug: log first 300 chars of page so we know what we're getting
            preview = html[:300].replace("\n", " ").strip()
            print(f"[{pid}] IVA SMS: 0 OTPs. Page preview: {preview}")

    return found


# ─────────────────────────────────────────────────────────────────────────────


def extract_otp_from_sms(sms_text):
    if not sms_text:
        return None
    sms_text = str(sms_text).strip()
    if len(sms_text) < 4:
        return None
    # Pure short numeric codes (4-8 digits only) — treat directly as OTP
    # WhatsApp/Telegram panels sometimes store just the raw code
    if re.match(r"^\d{4,8}$", sms_text):
        return sms_text
    # WhatsApp format: "123-456" or "123 456" (digits separated by dash/space)
    wa_m = re.match(r"^(\d{3})[- ](\d{3})$", sms_text)
    if wa_m:
        return wa_m.group(1) + wa_m.group(2)
    # Must have at least 1 letter for longer strings — pure long digit strings
    # (e.g. phone numbers "88017XXXXXXX") are not OTPs
    if not re.search(r"[a-zA-Z]", sms_text):
        return None
    cleaned = re.sub(r"(?<=\d) (?=\d)", "", sms_text)
    cleaned = re.sub(r"(\d)-(\d)", r"\1\2", cleaned)
    cleaned = re.sub(r"(\d)\.(\d)", r"\1\2", cleaned)
    m = re.search(r"\b(\d{4,8})\b", cleaned)
    return m.group(1) if m else None


# ── Panel 1 login & fetch (Mahofuza) ─────────────────────────────────────────


def p1_login():
    global _p1_session, _p1_sesskey
    sess = requests.Session()
    sess.headers.update({"User-Agent": "Mozilla/5.0"})
    try:
        r = sess.get(P1_LOGIN_PAGE, timeout=15)
        m = re.search(r"What is (\d+) \+ (\d+)", r.text)
        if not m:
            print("[P1] Could not find captcha")
            return False
        answer = int(m.group(1)) + int(m.group(2))
        r2 = sess.post(
            P1_SIGNIN_URL,
            data={"username": P1_USER_NAME, "password": P1_PASSWORD, "capt": answer},
            timeout=15,
            allow_redirects=True,
        )
        if "login" in r2.url.lower() or "login" in r2.text.lower()[:500]:
            print("[P1] Login failed — still on login page")
            return False
        r3 = sess.get(
            P1_CDR_PAGE, timeout=15, headers={"Referer": P1_BASE_URL + "/agent/"}
        )
        sk = re.search(r"sesskey=([A-Za-z0-9+/=]+)", r3.text)
        _p1_sesskey = sk.group(1) if sk else ""
        _p1_session = sess
        print(f"[P1] Logged in. sesskey={_p1_sesskey}")
        return True
    except Exception as e:
        print(f"[P1] Login error: {e}")
        return False


def fetch_panel1():
    global _p1_session, _p1_sesskey
    found = {}
    with _p1_lock:
        try:
            today = time.strftime("%Y-%m-%d")

            def build_url():
                return (
                    f"{P1_CDR_DATA_URL}"
                    f"?fdate1={today}%2000:00:00"
                    f"&fdate2={today}%2023:59:59"
                    f"&frange=&fclient=&fnum=&fcli=&fgdate=&fgmonth="
                    f"&fgrange=&fgclient=&fgnumber=&fgcli=&fg=0"
                    f"&sesskey={_p1_sesskey or ''}"
                )

            headers = {"Referer": P1_CDR_PAGE, "X-Requested-With": "XMLHttpRequest"}
            if _p1_session is None:
                if not p1_login():
                    return found
            r = _p1_session.get(build_url(), headers=headers, timeout=15)
            body = r.text.strip()
            if (
                r.status_code != 200
                or not body
                or body.startswith("<")
                or "Direct Script" in body
            ):
                print(f"[P1] Bad response ({r.status_code}), re-logging in.")
                _p1_session = None
                if not p1_login():
                    return found
                r = _p1_session.get(build_url(), headers=headers, timeout=15)
                body = r.text.strip()
            rows = json.loads(body).get("aaData", [])
            for row in rows:
                number = str(row[2]).strip()
                service = str(row[3]).strip()
                sms_txt = str(row[5]).strip()
                otp = extract_otp_from_sms(sms_txt)
                if otp:
                    key = f"{number}:{sms_txt}"
                    found[key] = (number, otp, sms_txt, service)
            _record_fetch("p1", len(rows))
            if found:
                print(f"[P1] ✅ Fetched {len(found)} records.")
        except Exception as e:
            print(f"[P1] Fetch error: {e}")
            _record_error("p1")
            _p1_session = None
    return found


# ── Panel 2 login & fetch (Sagardas50 / XISORA) ──────────────────────────────


def p2_login():
    global _p2_session
    sess = requests.Session()
    sess.headers.update({"User-Agent": "Mozilla/5.0"})
    try:
        r = sess.post(
            P2_SIGNIN_URL,
            data={"username": P2_USER_NAME, "password": P2_PASSWORD},
            timeout=15,
            allow_redirects=True,
        )
        if "signin" in r.url.lower() or "login" in r.url.lower():
            print("[P2] Login failed — still on login page")
            return False
        _p2_session = sess
        print(f"[P2] Logged in. URL={r.url}")
        return True
    except Exception as e:
        print(f"[P2] Login error: {e}")
        return False


def fetch_panel2():
    global _p2_session
    found = {}
    with _p2_lock:
        try:
            today = time.strftime("%Y-%m-%d")
            url = (
                f"{P2_DATA_URL}"
                f"?fdate1={today}%2000:00:00"
                f"&fdate2={today}%2023:59:59"
                f"&ftermination=&fclient=&fnum=&fcli="
                f"&fgdate=0&fgtermination=0&fgclient=0&fgnumber=0&fgcli=0&fg=0"
            )
            headers = {"Referer": P2_REPORTS_PAGE, "X-Requested-With": "XMLHttpRequest"}
            if _p2_session is None:
                if not p2_login():
                    return found
            r = _p2_session.get(url, headers=headers, timeout=15)
            body = r.text.strip()
            if r.status_code != 200 or not body or body.startswith("<"):
                print(f"[P2] Bad response ({r.status_code}), re-logging in.")
                _p2_session = None
                if not p2_login():
                    return found
                r = _p2_session.get(url, headers=headers, timeout=15)
                body = r.text.strip()
            rows = json.loads(body).get("aaData", [])
            for row in rows:
                if not isinstance(row[0], str):
                    continue
                number = str(row[2]).strip()
                service = str(row[3]).strip()
                sms_txt = str(row[10]).strip()
                otp = extract_otp_from_sms(sms_txt)
                if otp:
                    key = f"{number}:{sms_txt}"
                    found[key] = (number, otp, sms_txt, service)
            _record_fetch("p2", len(rows))
            if found:
                print(f"[P2] ✅ Fetched {len(found)} records.")
        except Exception as e:
            print(f"[P2] Fetch error: {e}")
            _record_error("p2")
            _p2_session = None
    return found


# ── Shared OTP processor ──────────────────────────────────────────────────────


def process_new_otps(current):
    global seen_otps
    for key, (number, otp, sms_txt, service) in current.items():
        with seen_lock:
            if key in seen_otps:
                continue
            seen_otps[key] = True
            save_json(SEEN_FILE, seen_otps)
        clean = re.sub(r"\D", "", str(number))
        with user_map_lock:
            t_start = assigned_time.get(clean)
        seconds = int(time.time() - t_start) if t_start else 0
        _dispatch_otp(otp, number, seconds, service, sms_txt or "")
        print(
            f"[MONITOR] ✅ Forwarded OTP={otp} for {number} ({service}) in {seconds}s"
        )


# ── Global OTP monitors ───────────────────────────────────────────────────────


def panel1_monitor():
    global seen_otps
    print("[P1-MONITOR] Started. Pre-loading existing records...")
    existing = fetch_panel1()
    with seen_lock:
        for key in existing:
            seen_otps[key] = True
        save_json(SEEN_FILE, seen_otps)
    print(f"[P1-MONITOR] Pre-loaded {len(existing)} records. Watching for new ones...")
    while True:
        try:
            process_new_otps(fetch_panel1())
        except Exception as e:
            print(f"[P1-MONITOR] Loop error: {e}")
        time.sleep(POLL_INTERVAL)


def panel2_monitor():
    global seen_otps
    print("[P2-MONITOR] Started. Pre-loading existing records...")
    existing = fetch_panel2()
    with seen_lock:
        for key in existing:
            seen_otps[key] = True
        save_json(SEEN_FILE, seen_otps)
    print(f"[P2-MONITOR] Pre-loaded {len(existing)} records. Watching for new ones...")
    while True:
        try:
            process_new_otps(fetch_panel2())
        except Exception as e:
            print(f"[P2-MONITOR] Loop error: {e}")
        time.sleep(POLL_INTERVAL)


# ── Panel 3 login & fetch (Rabbi1_FD) ────────────────────────────────────────


def p3_login():
    global _p3_session, _p3_csstr
    sess = requests.Session()
    sess.headers.update({"User-Agent": "Mozilla/5.0"})
    try:
        r = sess.get(P3_LOGIN_PAGE, timeout=15)
        m = re.search(r"What is (\d+) \+ (\d+)", r.text)
        if not m:
            print("[P3] Could not find captcha")
            return False
        answer = int(m.group(1)) + int(m.group(2))
        r2 = sess.post(
            P3_SIGNIN_URL,
            data={"username": P3_USER_NAME, "password": P3_PASSWORD, "capt": answer},
            timeout=15,
            allow_redirects=True,
        )
        if "login" in r2.url.lower() or "signin" in r2.url.lower():
            print("[P3] Login failed — still on login page")
            return False
        r3 = sess.get(
            P3_CDR_PAGE, timeout=15, headers={"Referer": P3_BASE_URL + "/agent/"}
        )
        cs = re.search(r"csstr=([a-f0-9]+)", r3.text)
        _p3_csstr = cs.group(1) if cs else ""
        _p3_session = sess
        print(f"[P3] Logged in. csstr={_p3_csstr}")
        return True
    except Exception as e:
        print(f"[P3] Login error: {e}")
        return False


def fetch_panel3():
    global _p3_session, _p3_csstr
    found = {}
    with _p3_lock:
        try:
            today = time.strftime("%Y-%m-%d")

            def build_url():
                return (
                    f"{P3_CDR_DATA_URL}"
                    f"?fdate1={today}%2000:00:00"
                    f"&fdate2={today}%2023:59:59"
                    f"&frange=&fclient=&fnum=&fcli=&fgdate=&fgmonth="
                    f"&fgrange=&fgclient=&fgnumber=&fgcli=&fg=0"
                    f"&csstr={_p3_csstr or ''}"
                )

            headers = {"Referer": P3_CDR_PAGE, "X-Requested-With": "XMLHttpRequest"}
            if _p3_session is None:
                if not p3_login():
                    return found
            r = _p3_session.get(build_url(), headers=headers, timeout=15)
            body = r.text.strip()
            if (
                r.status_code != 200
                or not body
                or body.startswith("<")
                or "Direct Script" in body
            ):
                print(f"[P3] Bad response ({r.status_code}), re-logging in.")
                _p3_session = None
                if not p3_login():
                    return found
                r = _p3_session.get(build_url(), headers=headers, timeout=15)
                body = r.text.strip()
            rows = json.loads(body).get("aaData", [])
            for row in rows:
                if not isinstance(row[0], str):
                    continue
                number = str(row[2]).strip()
                service = str(row[3]).strip()
                sms_txt = str(row[5]).strip()
                otp = extract_otp_from_sms(sms_txt)
                if otp:
                    key = f"{number}:{sms_txt}"
                    found[key] = (number, otp, sms_txt, service)
            _record_fetch("p3", len(rows))
            if found:
                print(f"[P3] ✅ Fetched {len(found)} records.")
        except Exception as e:
            print(f"[P3] Fetch error: {e}")
            _record_error("p3")
            _p3_session = None
    return found


def panel3_monitor():
    global seen_otps
    print("[P3-MONITOR] Started. Pre-loading existing records...")
    existing = fetch_panel3()
    with seen_lock:
        for key in existing:
            seen_otps[key] = True
        save_json(SEEN_FILE, seen_otps)
    print(f"[P3-MONITOR] Pre-loaded {len(existing)} records. Watching for new ones...")
    while True:
        try:
            process_new_otps(fetch_panel3())
        except Exception as e:
            print(f"[P3-MONITOR] Loop error: {e}")
        time.sleep(POLL_INTERVAL)


# ── Panel 4 login & fetch (Rabbi12 / 144.217.71.192) ─────────────────────────


def p4_login():
    global _p4_session, _p4_sesskey
    sess = requests.Session()
    sess.headers.update({"User-Agent": "Mozilla/5.0"})
    try:
        r = sess.get(P4_LOGIN_PAGE, timeout=15)
        m = re.search(r"What is (\d+) \+ (\d+)", r.text)
        if not m:
            print("[P4] Could not find captcha")
            return False
        answer = int(m.group(1)) + int(m.group(2))
        r2 = sess.post(
            P4_SIGNIN_URL,
            data={"username": P4_USER_NAME, "password": P4_PASSWORD, "capt": answer},
            timeout=15,
            allow_redirects=True,
        )
        if "SMSDashboard" not in r2.url and "agent" not in r2.url:
            print(f"[P4] Login failed: {r2.url}")
            return False
        r3 = sess.get(
            P4_CDR_PAGE, timeout=15, headers={"Referer": P4_BASE_URL + "/agent/"}
        )
        sk = re.search(r"sesskey=([A-Za-z0-9+/=]+)", r3.text)
        _p4_sesskey = sk.group(1) if sk else ""
        _p4_session = sess
        print(f"[P4] Logged in. sesskey={_p4_sesskey}")
        return True
    except Exception as e:
        print(f"[P4] Login error: {e}")
        return False


def fetch_panel4():
    global _p4_session, _p4_sesskey
    found = {}
    with _p4_lock:
        if not _p4_session and not p4_login():
            return found
        today = time.strftime("%Y-%m-%d")

        def build_url():
            return (
                f"{P4_CDR_DATA_URL}"
                f"?fdate1={today}%2000:00:00&fdate2={today}%2023:59:59"
                f"&frange=&fclient=&fnum=&fcli=&fgdate=&fgmonth="
                f"&fgrange=&fgclient=&fgnumber=&fgcli=&fg=0"
                f"&sesskey={_p4_sesskey}"
            )

        headers = {"Referer": P4_CDR_PAGE, "X-Requested-With": "XMLHttpRequest"}
        try:
            r = _p4_session.get(build_url(), headers=headers, timeout=15)
            body = r.text.strip()
            if (
                r.status_code != 200
                or not body
                or body.startswith("<")
                or "Direct Script" in body
            ):
                print(f"[P4] Bad response ({r.status_code}), re-logging in.")
                _p4_session = None
                if not p4_login():
                    return found
                r = _p4_session.get(build_url(), headers=headers, timeout=15)
                body = r.text.strip()
            rows = json.loads(body).get("aaData", [])
            for row in rows:
                if not isinstance(row[0], str):
                    continue
                number = str(row[2]).strip()
                service = str(row[3]).strip()
                sms_txt = str(row[5]).strip()
                otp = extract_otp_from_sms(sms_txt)
                if otp:
                    key = f"{number}:{sms_txt}"
                    found[key] = (number, otp, sms_txt, service)
            _record_fetch("p4", len(rows))
            if found:
                print(f"[P4] ✅ Fetched {len(found)} records.")
        except Exception as e:
            print(f"[P4] Fetch error: {e}")
            _record_error("p4")
            _p4_session = None
    return found


def panel4_monitor():
    global seen_otps
    print("[P4-MONITOR] Started. Pre-loading existing records...")
    existing = fetch_panel4()
    with seen_lock:
        for key in existing:
            seen_otps[key] = True
        save_json(SEEN_FILE, seen_otps)
    print(f"[P4-MONITOR] Pre-loaded {len(existing)} records. Watching for new ones...")
    while True:
        try:
            process_new_otps(fetch_panel4())
        except Exception as e:
            print(f"[P4-MONITOR] Loop error: {e}")
        time.sleep(POLL_INTERVAL)


# ── Panel 5 login & fetch (Rabbi12_v2 / 51.75.144.178) ───────────────────────


def p5_login():
    global _p5_session, _p5_sesskey
    sess = requests.Session()
    sess.headers.update({"User-Agent": "Mozilla/5.0"})
    try:
        r = sess.get(P5_LOGIN_PAGE, timeout=15)
        m = re.search(r"What is (\d+) \+ (\d+)", r.text)
        if not m:
            print("[P5] Could not find captcha")
            return False
        answer = int(m.group(1)) + int(m.group(2))
        r2 = sess.post(
            P5_SIGNIN_URL,
            data={"username": P5_USER_NAME, "password": P5_PASSWORD, "capt": answer},
            timeout=15,
            allow_redirects=True,
        )
        if "SMSDashboard" not in r2.url and "agent" not in r2.url:
            print(f"[P5] Login failed: {r2.url}")
            return False
        r3 = sess.get(
            P5_CDR_PAGE, timeout=15, headers={"Referer": P5_BASE_URL + "/agent/"}
        )
        sk = re.search(r"sesskey=([A-Za-z0-9+/=]+)", r3.text)
        _p5_sesskey = sk.group(1) if sk else ""
        _p5_session = sess
        print(f"[P5] Logged in. sesskey={_p5_sesskey}")
        return True
    except Exception as e:
        print(f"[P5] Login error: {e}")
        return False


def fetch_panel5():
    global _p5_session, _p5_sesskey
    found = {}
    with _p5_lock:
        if not _p5_session and not p5_login():
            return found
        today = time.strftime("%Y-%m-%d")

        def build_url():
            return (
                f"{P5_CDR_DATA_URL}"
                f"?fdate1={today}%2000:00:00&fdate2={today}%2023:59:59"
                f"&frange=&fclient=&fnum=&fcli=&fgdate=&fgmonth="
                f"&fgrange=&fgclient=&fgnumber=&fgcli=&fg=0"
                f"&sesskey={_p5_sesskey}"
            )

        headers = {"Referer": P5_CDR_PAGE, "X-Requested-With": "XMLHttpRequest"}
        try:
            r = _p5_session.get(build_url(), headers=headers, timeout=15)
            body = r.text.strip()
            if (
                r.status_code != 200
                or not body
                or body.startswith("<")
                or "Direct Script" in body
            ):
                print(f"[P5] Bad response ({r.status_code}), re-logging in.")
                _p5_session = None
                if not p5_login():
                    return found
                r = _p5_session.get(build_url(), headers=headers, timeout=15)
                body = r.text.strip()
            rows = json.loads(body).get("aaData", [])
            for row in rows:
                if not isinstance(row[0], str):
                    continue
                number = str(row[2]).strip()
                service = str(row[3]).strip()
                sms_txt = str(row[5]).strip()
                otp = extract_otp_from_sms(sms_txt)
                if otp:
                    key = f"{number}:{sms_txt}"
                    found[key] = (number, otp, sms_txt, service)
            _record_fetch("p5", len(rows))
            if found:
                print(f"[P5] ✅ Fetched {len(found)} records.")
        except Exception as e:
            print(f"[P5] Fetch error: {e}")
            _record_error("p5")
            _p5_session = None
    return found


def panel5_monitor():
    global seen_otps
    print("[P5-MONITOR] Started. Pre-loading existing records...")
    existing = fetch_panel5()
    with seen_lock:
        for key in existing:
            seen_otps[key] = True
        save_json(SEEN_FILE, seen_otps)
    print(f"[P5-MONITOR] Pre-loaded {len(existing)} records. Watching for new ones...")
    while True:
        try:
            process_new_otps(fetch_panel5())
        except Exception as e:
            print(f"[P5-MONITOR] Loop error: {e}")
        time.sleep(POLL_INTERVAL)


# ── Panel 6 login & fetch (TrueSMS.net / SMSRanges) ──────────────────────────


def p6_login():
    global _p6_session, _p6_sesskey
    sess = requests.Session()
    sess.headers.update({"User-Agent": "Mozilla/5.0"})
    try:
        r = sess.get(P6_LOGIN_PAGE, timeout=20, verify=False)
        m = re.search(r"What is (\d+) \+ (\d+)", r.text)
        if m:
            answer = int(m.group(1)) + int(m.group(2))
            r2 = sess.post(
                P6_SIGNIN_URL,
                data={
                    "username": P6_USER_NAME,
                    "password": P6_PASSWORD,
                    "capt": answer,
                },
                timeout=20,
                allow_redirects=True,
                verify=False,
            )
        else:
            r2 = sess.post(
                P6_SIGNIN_URL,
                data={"username": P6_USER_NAME, "password": P6_PASSWORD},
                timeout=20,
                allow_redirects=True,
                verify=False,
            )
        if "login" in r2.url.lower() and "agent" not in r2.url.lower():
            print(f"[P6] Login failed: {r2.url}")
            return False
        r3 = sess.get(
            P6_CDR_PAGE,
            timeout=20,
            headers={"Referer": P6_BASE_URL + "/agent/"},
            verify=False,
        )
        sk = re.search(r"sesskey=([A-Za-z0-9+/=]+)", r3.text)
        cs = re.search(r"csstr=([a-f0-9]+)", r3.text)
        _p6_sesskey = sk.group(1) if sk else (cs.group(1) if cs else "")
        _p6_session = sess
        print(f"[P6] Logged in. token={_p6_sesskey[:10] if _p6_sesskey else 'none'}")
        return True
    except Exception as e:
        print(f"[P6] Login error: {e}")
        return False


def fetch_panel6():
    global _p6_session, _p6_sesskey
    found = {}
    with _p6_lock:
        try:
            today = time.strftime("%Y-%m-%d")

            def build_url():
                return (
                    f"{P6_CDR_DATA_URL}"
                    f"?fdate1={today}%2000:00:00"
                    f"&fdate2={today}%2023:59:59"
                    f"&frange=&fclient=&fnum=&fcli=&fgdate=&fgmonth="
                    f"&fgrange=&fgclient=&fgnumber=&fgcli=&fg=0"
                    f"&sesskey={_p6_sesskey or ''}"
                )

            headers = {"Referer": P6_CDR_PAGE, "X-Requested-With": "XMLHttpRequest"}
            if _p6_session is None:
                if not p6_login():
                    return found
            r = _p6_session.get(build_url(), headers=headers, timeout=20, verify=False)
            body = r.text.strip()
            if (
                r.status_code != 200
                or not body
                or body.startswith("<")
                or "Direct Script" in body
            ):
                print(f"[P6] Bad response ({r.status_code}), re-logging in.")
                _p6_session = None
                if not p6_login():
                    return found
                r = _p6_session.get(
                    build_url(), headers=headers, timeout=20, verify=False
                )
                body = r.text.strip()
            rows = json.loads(body).get("aaData", [])
            for row in rows:
                if not isinstance(row[0], str):
                    continue
                number = str(row[2]).strip()
                service = str(row[3]).strip() if len(row) > 3 else "TrueSMS"
                sms_txt = str(row[5]).strip() if len(row) > 5 else ""
                if not sms_txt and len(row) > 4:
                    sms_txt = str(row[4]).strip()
                otp = extract_otp_from_sms(sms_txt)
                if otp:
                    key = f"{number}:{sms_txt}"
                    found[key] = (number, otp, sms_txt, service)
            _record_fetch("p6", len(rows))
            if found:
                print(f"[P6] ✅ Fetched {len(found)} records.")
        except Exception as e:
            print(f"[P6] Fetch error: {e}")
            _record_error("p6")
            _p6_session = None
    return found


def panel6_monitor():
    global seen_otps
    print("[P6-MONITOR] Started (TrueSMS/SMSRanges). Pre-loading existing records...")
    existing = fetch_panel6()
    with seen_lock:
        for key in existing:
            seen_otps[key] = True
        save_json(SEEN_FILE, seen_otps)
    print(f"[P6-MONITOR] Pre-loaded {len(existing)} records. Watching for new ones...")
    while True:
        try:
            process_new_otps(fetch_panel6())
        except Exception as e:
            print(f"[P6-MONITOR] Loop error: {e}")
        time.sleep(POLL_INTERVAL)


# ── Demo OTP monitor ──────────────────────────────────────────────────────────


def demo_monitor():
    print("[DEMO] Thread started.")
    while True:
        now = time.time()
        with _demo_lock:
            configs = list(_demo_configs)
        for cfg in configs:
            if not cfg.get("active"):
                continue
            cid = cfg["id"]
            if now >= _demo_next_fire.get(cid, 0):
                _demo_next_fire[cid] = now + cfg["interval"]
                services = cfg.get("services") or ["Facebook"]
                for svc in services:
                    otp = "".join([str(random.randint(0, 9)) for _ in range(cfg["digits"])])
                    number = random.choice(cfg["numbers"])
                    try:
                        send_otp_message(get_otp_group_id(), otp, number, "—", svc, "")
                    except Exception as e:
                        print(f"[DEMO] {cfg['name']} send error ({svc}): {e}")
        time.sleep(1)


def demo_status_text():
    with _demo_lock:
        configs = list(_demo_configs)
    running = [c for c in configs if c.get("active")]
    status = f"🟢 <b>{len(running)} টি চলছে</b>" if running else "🔴 <b>সব বন্ধ</b>"
    lines = (
        f"🎭🔥 <b>DEMO OTP PANEL</b> 🔥🎭\n"
        f"⚡━━━━━━━━━━━━━━━━⚡\n\n"
        f"📡 <b>Status ▸▸</b>  {status}\n"
        f"📋 <b>Configs:</b>  {len(configs)} টি\n\n"
    )
    for cfg in configs:
        icon = "🟢" if cfg.get("active") else "🔴"
        svcs = ", ".join(cfg.get("services") or ["?"])
        nums = cfg["numbers"]
        lines += (
            f"{icon} <b>{cfg['name']}</b>\n"
            f"  💬 {svcs}  |  🔢 {cfg['digits']} digits  |  ⏱️ {cfg['interval']}s  |  📱 {len(nums)} num\n\n"
        )
    lines += "⚡━━━━━━━━━━━━━━━━⚡"
    return lines


def demo_cfg_inline_markup():
    with _demo_lock:
        configs = list(_demo_configs)
    markup = types.InlineKeyboardMarkup(row_width=2)
    for cfg in configs:
        icon = "⏹️ Stop" if cfg.get("active") else "▶️ Start"
        action = "stop" if cfg.get("active") else "start"
        markup.add(
            types.InlineKeyboardButton(
                f"{icon}  {cfg['name']}",
                callback_data=f"cfg_toggle:{cfg['id']}:{action}",
            )
        )
    return markup


def demo_menu_markup():
    m = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
    with _demo_lock:
        cfg_count = len(_demo_configs)
    m.add("➕ 𝗖𝗼𝗻𝗳𝗶𝗴 𝗬𝗼𝗴 𝗞𝗼𝗿𝗼")
    if cfg_count > 0:
        m.add("🗑️ 𝗖𝗼𝗻𝗳𝗶𝗴 𝗠𝘂𝗰𝗵𝗼")
    m.add("🔙 𝗔𝗗𝗠𝗜𝗡 𝗣𝗔𝗡𝗘𝗟")
    return m


# ── Menus ─────────────────────────────────────────────────────────────────────


def main_menu(user_id):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
    markup.add(types.KeyboardButton("☎️ 𝗩𝟭 𝗡𝗨𝗠𝗕𝗔𝗥 ☎️"))
    markup.add(types.KeyboardButton("🔄 𝗩𝟮 𝗦𝗪𝗜𝗧𝗖𝗛"))
    markup.add(types.KeyboardButton("📊 𝗦𝗧𝗢𝗖𝗞"), types.KeyboardButton("📞 𝗦𝗔𝗣𝗢𝗥𝗧"))
    if user_id in ADMIN_IDS:
        markup.add(types.KeyboardButton("⚙️ 𝗔𝗗𝗠𝗜𝗡 𝗣𝗔𝗡𝗘𝗟 ⚙️"))
    return markup


def v2_switch_menu():
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
    markup.add(types.KeyboardButton("🔴 𝗟𝗜𝗩𝗘 𝗥𝗔𝗡𝗚𝗘"))
    markup.add(types.KeyboardButton("⌨️ 𝗖𝗨𝗦𝗧𝗢𝗠 𝗥𝗔𝗡𝗚𝗘"))
    markup.add(types.KeyboardButton("🔙 𝗩𝟭 𝗦𝗪𝗜𝗧𝗖𝗛"))
    return markup


def save_services():
    save_json(SERVICES_FILE, _services)
    _sync_settings_to_botpy()


def _get_svc_map():
    return {s["label"]: s["key"] for s in _services}


SERVICE_BUTTON_MAP = {}


def show_services(message):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
    btns = [types.KeyboardButton(s["label"]) for s in _services]
    for i in range(0, len(btns), 2):
        markup.add(*btns[i:i + 2])
    markup.add(types.KeyboardButton("🔙 Main Menu"))
    bot.send_message(
        message.chat.id,
        "🛠 <b>Select Service:</b>",
        reply_markup=markup,
        parse_mode="HTML",
    )


def show_countries(chat_id, svc):
    markup = types.InlineKeyboardMarkup(row_width=2)
    btns = []
    if svc in stock:
        for cnt, nums in stock[svc].items():
            if nums:
                _, flag = get_country_details(nums[0])
                btns.append(
                    types.InlineKeyboardButton(
                        f"{flag} {cnt}", callback_data=f"n:{svc}:{cnt}"
                    )
                )
    if btns:
        markup.add(*btns)
    markup.add(
        types.InlineKeyboardButton("⬅️ 𝗕𝗮𝗰𝗸", callback_data="back_to_services")
    )
    bot.send_message(
        chat_id,
        f"🌍 <b>Country select koro:</b>",
        reply_markup=markup,
        parse_mode="HTML",
    )


# ── Handlers ──────────────────────────────────────────────────────────────────


@bot.message_handler(commands=["start"])
def start_cmd(message):
    u = message.from_user
    register_user(
        message.chat.id,
        first_name=u.first_name or "",
        last_name=u.last_name or "",
        username=u.username or "",
    )
    uname = f"@{u.username}" if u.username else (u.first_name or "User")
    uid_str = u.id
    markup = types.InlineKeyboardMarkup()
    _grp = get_otp_group_link() or CHANNEL_1
    if _grp:
        markup.add(types.InlineKeyboardButton("🔥 𝗢𝗧𝗣 𝗚𝗿𝘂𝗽 𝗝𝗢𝗜𝗡 🔥", url=_grp))
    if get_channel2():
        markup.add(types.InlineKeyboardButton("📢 𝗠𝗮𝗶𝗻 𝗖𝗵𝗮𝗻𝗻𝗲𝗹 𝗝𝗢𝗜𝗡", url=get_channel2()))
    markup.add(types.InlineKeyboardButton("✅ 𝗩𝗘𝗥𝗜𝗙𝗬 𝗞𝗢𝗥𝗢 ✅", callback_data="v"))
    bot.send_message(
        message.chat.id,
        get_template("start").format(uname=uname, uid=uid_str),
        reply_markup=markup,
        parse_mode="HTML",
    )


@bot.message_handler(commands=["test"])
def test_cmd(message):
    fake_otp = str(random.randint(100000, 999999))
    fake_number = "8801712345678"
    fake_svc = "Instagram"
    fake_secs = 12
    fake_sms = f"Your Instagram code is {fake_otp}. Don't share this code."
    send_otp_message(message.chat.id, fake_otp, fake_number, fake_secs, fake_svc, fake_sms)
    try:
        send_otp_message(get_otp_group_id(), fake_otp, fake_number, fake_secs, fake_svc, fake_sms)
        bot.send_message(
            message.chat.id, "✅ Group-eও pathano hoyeche!", parse_mode="HTML"
        )
    except Exception as e:
        bot.send_message(
            message.chat.id,
            f"⚠️ Group-e pathate parina: <code>{e}</code>",
            parse_mode="HTML",
        )


@bot.message_handler(commands=["panels"])
def panels_cmd(message):
    if message.from_user.id not in ADMIN_IDS:
        return
    with _stats_lock:
        stats = {k: dict(v) for k, v in _panel_stats.items()}
    lines = ""
    for pid in ["p1", "p2", "p3", "p4", "p5", "p6"]:
        s = stats.get(pid, {})
        if s.get("last"):
            ago = int(time.time() - s["last"])
            last_str = f"{ago}s ago"
        else:
            last_str = "never"
        err_str = f"  ⚠️ {s['errors']} err" if s.get("errors") else ""
        lines += (
            f"{s.get('status', '⏳')} <b>{s.get('name', '?')}</b>\n"
            f"   🌐 <code>{s.get('host', '?')}</code>\n"
            f"   📊 {s.get('count', 0)} records  •  🕐 {last_str}{err_str}\n\n"
        )
    with _demo_lock:
        demo_on = _demo_active
    demo_str = "🟢 Running" if demo_on else "🔴 Stopped"
    bot.send_message(
        message.chat.id,
        f"📡 <b>PANEL STATUS</b>\n"
        f"⚡━━━━━━━━━━━━━━━━⚡\n\n"
        f"{lines}"
        f"🎭 <b>Demo OTP:</b>  {demo_str}\n\n"
        f"⚡━━━━━━━━━━━━━━━━⚡\n"
        f"🔄 <i>Updates every {POLL_INTERVAL}s</i>",
        parse_mode="HTML",
    )
    caller_uid = message.from_user.id
    # Super admin sees all, others see only their own panels
    dp_copy = [
        p for p in _dynamic_panels
        if is_super_admin(caller_uid) or p.get("admin_id") == caller_uid
    ]
    if dp_copy:
        dp_lines = ""
        for p in dp_copy:
            pid = p["id"]
            with _stats_lock:
                s = _panel_stats.get(pid, {})
            st = s.get("status", "⏳")
            cnt = s.get("count", 0)
            err = s.get("errors", 0)
            t = s.get("last")
            last_str = f"{int(time.time() - t)}s ago" if t else "never"
            err_str = f"  ⚠️ {err} err" if err else ""
            dp_lines += (
                f"{st} <b>{p.get('username', '?')}</b> <code>[{pid}]</code>\n"
                f"   🌐 <code>{p.get('host', '?')}</code>\n"
                f"   📊 {cnt} records  •  🕐 {last_str}{err_str}\n\n"
            )
        bot.send_message(
            message.chat.id,
            f"📡 <b>DYNAMIC PANELS</b>\n"
            f"⚡━━━━━━━━━━━━━━━━⚡\n\n"
            f"{dp_lines}"
            f"💡 <i>/addpanel diye naya panel add koro</i>",
            parse_mode="HTML",
        )
    else:
        bot.send_message(
            message.chat.id,
            "📋 <b>Tomar kono dynamic panel nei.</b>\n\n"
            "💡 <i>/addpanel diye notun panel add koro.</i>",
            parse_mode="HTML",
        )


@bot.message_handler(commands=["broadcast"])
def broadcast_cmd(message):
    if message.from_user.id not in ADMIN_IDS:
        return
    msg = bot.send_message(
        message.chat.id,
        "✍️ <b>Broadcast content পাঠাও:</b> \n\n"
        "📝 Text\n🖼️ Photo\n🎥 Video\n🎭 Sticker\n"
        "🎞️ GIF / Animation\n🎵 Audio / Music\n🎤 Voice message\n📎 Document / APK / ZIP / PDF\n\n"
        "<i>Caption support ache — sob kichute!</i>",
        parse_mode="HTML",
    )
    bot.register_next_step_handler(msg, do_broadcast)


def _clr_service_markup():
    markup = types.InlineKeyboardMarkup(row_width=2)
    services = [
        ("facebook", "💬"),
        ("instagram", "📸"),
        ("whatsapp", "📱"),
        ("telegram", "✈️"),
        ("binance", "🪙"),
        ("pc clone", "💻"),
    ]
    for svc, icon in services:
        total = sum(len(v) for v in stock.get(svc, {}).values())
        markup.add(
            types.InlineKeyboardButton(
                f"{icon} {svc.upper()} ({total})", callback_data=f"clr_s:{svc}"
            )
        )
    markup.add(types.InlineKeyboardButton(" Clear ALL Stock", callback_data="clr_all"))
    return markup


@bot.message_handler(commands=["addpanel"])
def addpanel_cmd(message):
    if message.from_user.id not in ADMIN_IDS:
        return
    _show_addpanel_type_select(message.chat.id, message.from_user.id)


def _show_addpanel_type_select(chat_id, uid):
    markup = types.InlineKeyboardMarkup(row_width=1)
    markup.add(
        types.InlineKeyboardButton("🔑 Username + Password দিয়ে Add", callback_data="aptype:pass"),
        types.InlineKeyboardButton("🗝️ API Key দিয়ে Add", callback_data="aptype:apikey"),
    )
    bot.send_message(
        chat_id,
        "🔧🔥 <b>ADD NEW PANEL</b> 🔥🔧\n\n"
        "Panel কীভাবে add করতে চাও?\n\n"
        "🔑 <b>Username + Password</b> — সাধারণ login করে add\n"
        "🗝️ <b>API Key</b> — panel-এর API key দিয়ে add",
        reply_markup=markup,
        parse_mode="HTML",
    )


def _ap_get_url(message):
    if message.from_user.id not in ADMIN_IDS:
        return
    if _is_back(message.text):
        _addpanel_state.pop(message.from_user.id, None)
        _go_admin_panel(message)
        return
    if _intercept_menu_btn(message):
        return
    url = (message.text or "").strip()

    # Use the universal base extractor — handles ANY path prefix (/konekta, /ints, etc.)
    base_url = _extract_panel_base_url(url) if re.match(r"https?://", url, re.IGNORECASE) else None

    if not base_url:
        msg = bot.send_message(
            message.chat.id,
            "❌ <b>Valid URL dao!</b>\n\n"
            "Example:\n"
            "• <code>http://1.2.3.4</code>\n"
            "• <code>http://1.2.3.4/konekta</code>\n"
            "• <code>http://1.2.3.4/konekta/agent/SMSCDRReports</code>\n"
            "• <code>https://mypanel.com</code>",
            reply_markup=_back_admin_kb(),
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, _ap_get_url)
        return

    host_m = re.search(r"//([^/]+)", base_url)
    uid = message.from_user.id
    _addpanel_state[uid]["data"]["base_url"] = base_url
    _addpanel_state[uid]["data"]["host"] = host_m.group(1) if host_m else base_url
    _addpanel_state[uid]["data"]["url_hint"] = url  # preserve original URL as hint

    # ── IVA SMS special flow (ivasms.com) — cookie only ─────────────────────
    if "ivasms.com" in base_url.lower():
        msg = bot.send_message(
            message.chat.id,
            "🌐 <b>IVA SMS Panel detect hoise!</b>\n\n"
            "⚠️ Railway server IP theke email/password login Cloudflare block kore.\n"
            "<b>Browser cookie diye login korte hobe.</b>\n\n"
            "📋 <b>Cookie পাওয়ার নিয়ম:</b>\n"
            "1. Phone/PC-এ Chrome-এ <b>ivasms.com</b> login করো\n"
            "2. এই link open করো browser-এ:\n"
            "   <code>javascript:document.cookie</code>\n"
            "   (address bar-এ paste করো)\n"
            "   <b>অথবা</b> PC-তে: F12 → Application → Cookies → https://ivasms.com\n"
            "3. <code>laravel_session</code> value copy করো\n\n"
            "🍪 এখন cookie paste করো:\n"
            "<code>laravel_session=eyJ...</code>",
            reply_markup=_back_admin_kb(),
            parse_mode="HTML",
            disable_web_page_preview=True,
        )
        bot.register_next_step_handler(msg, _iva_ap_get_cookie)
        return
    # ─────────────────────────────────────────────────────────────────────────

    msg = bot.send_message(
        message.chat.id,
        f"✅ <b>URL set:</b> <code>{base_url}</code>\n\n"
        f"👤 <b>Step 2/3:</b> Username pathao:",
        reply_markup=_back_admin_kb(),
        parse_mode="HTML",
    )
    bot.register_next_step_handler(msg, _ap_get_user)


def _ap_get_user(message):
    if message.from_user.id not in ADMIN_IDS:
        return
    if _is_back(message.text):
        _addpanel_state.pop(message.from_user.id, None)
        _go_admin_panel(message)
        return
    if _intercept_menu_btn(message):
        return
    username = (message.text or "").strip()
    if not username:
        msg = bot.send_message(message.chat.id, "❌ Username dao:", reply_markup=_back_admin_kb())
        bot.register_next_step_handler(msg, _ap_get_user)
        return
    _addpanel_state[message.from_user.id]["data"]["username"] = username
    msg = bot.send_message(
        message.chat.id,
        f"✅ Username: <code>{username}</code>\n\n🔑 <b>Step 3/3:</b> Password pathao:",
        reply_markup=_back_admin_kb(),
        parse_mode="HTML",
    )
    bot.register_next_step_handler(msg, _ap_get_pass)


def _ap_get_pass(message):
    if message.from_user.id not in ADMIN_IDS:
        return
    uid = message.from_user.id
    if _is_back(message.text):
        _addpanel_state.pop(uid, None)
        _go_admin_panel(message)
        return
    if _intercept_menu_btn(message):
        return
    password = (message.text or "").strip()
    if not password:
        msg = bot.send_message(message.chat.id, "❌ Password dao:", reply_markup=_back_admin_kb())
        bot.register_next_step_handler(msg, _ap_get_pass)
        return
    data = _addpanel_state.get(uid, {}).get("data", {})
    data["password"] = password
    wait_msg = bot.send_message(
        message.chat.id,
        "⏳🔥 <b>Connecting & auto-detecting panel type...</b>\n"
        "<i>Login page khujchi, captcha solve korchi, data endpoint test korchi...</i>",
        parse_mode="HTML",
    )
    panel_id = f"d{int(time.time()) % 100000}"
    panel = {
        "id": panel_id,
        "host": data.get("host", ""),
        "base_url": data.get("base_url", ""),
        "url_hint": data.get("url_hint", ""),
        "username": data.get("username", ""),
        "password": password,
        "engine": "ints_smscdr",
        "data_path": "/agent/res/data_smscdr.php",
        "admin_id": uid,
    }
    chat_id = message.chat.id
    _addpanel_state.pop(uid, None)

    def _do_add():
        sess, token, det_engine, det_path = _universal_login(panel)
        try:
            bot.delete_message(chat_id, wait_msg.message_id)
        except Exception:
            pass
        if not sess:
            # Save panel data for force-add (Railway IP might be blocked by panel)
            _pending_force_add[panel_id] = panel
            force_markup = types.InlineKeyboardMarkup(row_width=1)
            force_markup.add(
                types.InlineKeyboardButton(
                    "⚠️ Force Add করো (Login Skip)",
                    callback_data=f"forceadd:{panel_id}",
                )
            )
            force_markup.add(
                types.InlineKeyboardButton("❌ বাদ দাও", callback_data=f"forceadd_cancel:{panel_id}")
            )
            bot.send_message(
                chat_id,
                "⚠️ <b>Login Verify করা গেলো না!</b>\n\n"
                "Railway server-er IP অনেক panel block করে।\n"
                "তবুও panel credentials সেভ করতে চাইলে\n"
                "<b>Force Add</b> করো — panel নিজেই পরে login করার চেষ্টা করবে।\n\n"
                f"🌐 Host: <code>{data.get('host', '')}</code>\n"
                f"👤 User: <code>{data.get('username', '')}</code>",
                reply_markup=force_markup,
                parse_mode="HTML",
            )
            return
        if det_engine:
            panel["engine"] = det_engine
            panel["data_path"] = det_path
        _dynamic_sessions[panel_id] = {"session": sess, "token": token}
        _dynamic_panels.append(panel)
        save_dynamic_panels()
        _start_dynamic_panel(panel)
        engine_label = {
            "ints_smscdr":   "INTS — SMSCDRStats",
            "ints_smsranges":"INTS — SMSRanges",
            "xisora":        "Xisora",
            "html_scrape":   "HTML Scrape",
        }.get(panel.get("engine", ""), panel.get("engine", "Auto"))
        bot.send_message(
            chat_id,
            f"✅🔥 <b>PANEL ADDED & STARTED!</b> 🔥✅\n"
            f"⚡━━━━━━━━━━━━━━━━⚡\n\n"
            f"🆔 <b>ID      ▸▸</b> <code>{panel_id}</code>\n"
            f"🌐 <b>Host    ▸▸</b> <code>{data.get('host','')}</code>\n"
            f"👤 <b>User    ▸▸</b> <code>{data.get('username','')}</code>\n"
            f"🔍 <b>Engine  ▸▸</b> <code>{engine_label}</code>\n"
            f"📂 <b>Endpoint▸▸</b> <code>{panel.get('data_path','')}</code>\n\n"
            f"📡 Monitor thread started! /panels diye check koro.",
            parse_mode="HTML",
        )

    threading.Thread(target=_do_add, daemon=True).start()


# ── IVA SMS add-panel flow (cookie only — email/pass blocked by Cloudflare) ───

def _iva_ap_get_email(message):
    """Legacy handler — redirects to cookie flow immediately."""
    _iva_ap_get_cookie(message)


def _iva_ap_get_pass(message):
    """Legacy handler — redirects to cookie flow immediately."""
    _iva_ap_get_cookie(message)


def _iva_ap_get_cookie(message):
    """Collect browser cookie and connect to ivasms.com."""
    uid = message.from_user.id
    if uid not in ADMIN_IDS:
        return
    if _is_back(message.text):
        _addpanel_state.pop(uid, None)
        _go_admin_panel(message)
        return
    if _intercept_menu_btn(message):
        return
    cookie_str = (message.text or "").strip()
    if not cookie_str or "=" not in cookie_str:
        msg = bot.send_message(message.chat.id,
            "❌ <b>Cookie dao!</b>\n\n"
            "Format: <code>laravel_session=eyJ0...</code>\n\n"
            "📱 <b>Phone-এ কীভাবে পাবে:</b>\n"
            "1. Chrome-এ ivasms.com login করো\n"
            "2. Address bar-এ type করো:\n"
            "   <code>javascript:alert(document.cookie)</code>\n"
            "3. Popup-এ যা আসবে সেটা copy করো\n\n"
            "💻 <b>PC-তে:</b> F12 → Application → Cookies → ivasms.com → laravel_session copy",
            reply_markup=_back_admin_kb(), parse_mode="HTML",
            disable_web_page_preview=True)
        bot.register_next_step_handler(msg, _iva_ap_get_cookie)
        return
    _iva_do_connect(message, cookie_str=cookie_str)


def _iva_do_connect(message, cookie_str):
    """Build panel dict and connect using cookie."""
    uid = message.from_user.id
    _addpanel_state.pop(uid, None)
    chat_id = message.chat.id

    panel_id = f"iva{int(time.time()) % 100000}"
    panel = {
        "id": panel_id,
        "host": "ivasms.com",
        "base_url": "https://ivasms.com",
        "url_hint": "https://ivasms.com/portal/sms/received",
        "username": "ivasms",
        "password": "",
        "cookie_str": cookie_str,
        "engine": "iva_sms",
        "data_path": "/portal/sms/received",
        "admin_id": uid,
    }

    wait_msg = bot.send_message(chat_id,
        "⏳ <b>IVA SMS — cookie diye login korchi...</b>", parse_mode="HTML")

    def _do():
        ok = _iva_login(panel)
        try:
            bot.delete_message(chat_id, wait_msg.message_id)
        except Exception:
            pass

        if not ok:
            msg2 = bot.send_message(chat_id,
                "❌ <b>Cookie kaj koreni!</b>\n\n"
                "Possible karon:\n"
                "• Cookie expire hoyeche (fresh login koro)\n"
                "• Pura cookie copy hoy nai\n\n"
                "Abar fresh cookie pathao:\n"
                "<code>laravel_session=eyJ0...</code>",
                reply_markup=_back_admin_kb(), parse_mode="HTML")
            _addpanel_state[uid] = {"step": "iva_cookie", "data": {}}
            bot.register_next_step_handler(msg2, _iva_ap_get_cookie)
            return

        _dynamic_panels.append(panel)
        save_dynamic_panels()
        _start_dynamic_panel(panel)

        bot.send_message(chat_id,
            f"✅🔥 <b>IVA SMS PANEL ADDED!</b> 🔥✅\n"
            f"⚡━━━━━━━━━━━━━━━━⚡\n\n"
            f"🆔 <b>ID     ▸▸</b> <code>{panel_id}</code>\n"
            f"🌐 <b>Host   ▸▸</b> <code>ivasms.com</code>\n"
            f"🔑 <b>Login  ▸▸</b> <code>Cookie ✅</code>\n\n"
            f"📡 Monitor started! New OTP ashle group-e pathabe.\n"
            f"⚠️ Cookie expire hole: <code>/ivacookie</code>",
            parse_mode="HTML")

    threading.Thread(target=_do, daemon=True).start()


# ── API Key Panel Add Flow ─────────────────────────────────────────────────────

_apk_state = {}   # uid → {"url": ..., "api_key": ...}


def _apk_start(message):
    """Ask for panel URL (Step 1 of API key flow)."""
    uid = message.from_user.id
    if uid not in ADMIN_IDS:
        return
    _apk_state[uid] = {}
    msg = bot.send_message(
        message.chat.id,
        "🗝️🔥 <b>API KEY দিয়ে PANEL ADD</b> 🔥🗝️\n\n"
        "📡 <b>Step 1/2:</b> Panel-এর URL পাঠাও\n\n"
        "✅ <b>যেকোনো format চলবে:</b>\n"
        "• <code>http://1.2.3.4</code>\n"
        "• <code>http://1.2.3.4/api</code>\n"
        "• <code>https://mypanel.com</code>\n"
        "• <code>https://mypanel.com/api/sms</code>",
        reply_markup=_back_admin_kb(),
        parse_mode="HTML",
    )
    bot.register_next_step_handler(msg, _apk_get_url)


def _apk_get_url(message):
    uid = message.from_user.id
    if uid not in ADMIN_IDS:
        return
    if _is_back(message.text):
        _apk_state.pop(uid, None)
        _go_admin_panel(message)
        return
    if _intercept_menu_btn(message):
        return
    url = (message.text or "").strip()
    if not re.match(r"https?://", url, re.IGNORECASE):
        msg = bot.send_message(
            message.chat.id,
            "❌ Valid URL দাও! (http:// বা https:// দিয়ে শুরু করো)",
            reply_markup=_back_admin_kb(), parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, _apk_get_url)
        return
    # Extract base URL
    m = re.match(r"(https?://[^/]+(?:/[^?#]*)?)", url, re.IGNORECASE)
    base_url = m.group(1).rstrip("/") if m else url.rstrip("/")
    # If URL contains known API paths, strip them to get clean base
    for suffix in ["/api/sms", "/api/messages", "/api/received", "/api/v1", "/api"]:
        if base_url.lower().endswith(suffix):
            base_url = base_url[: -len(suffix)]
            break
    _apk_state[uid]["base_url"] = base_url
    host_m = re.search(r"//([^/]+)", base_url)
    _apk_state[uid]["host"] = host_m.group(1) if host_m else base_url

    msg = bot.send_message(
        message.chat.id,
        f"✅ URL: <code>{base_url}</code>\n\n"
        "🗝️ <b>Step 2/2:</b> Panel-এর <b>API Key</b> পাঠাও:\n\n"
        "<i>Panel-এর settings/profile/API section থেকে copy করো।</i>",
        reply_markup=_back_admin_kb(),
        parse_mode="HTML",
    )
    bot.register_next_step_handler(msg, _apk_get_key)


def _apk_get_key(message):
    uid = message.from_user.id
    if uid not in ADMIN_IDS:
        return
    if _is_back(message.text):
        _apk_state.pop(uid, None)
        _go_admin_panel(message)
        return
    if _intercept_menu_btn(message):
        return
    api_key = (message.text or "").strip()
    if not api_key:
        msg = bot.send_message(message.chat.id, "❌ API Key দাও:", reply_markup=_back_admin_kb())
        bot.register_next_step_handler(msg, _apk_get_key)
        return

    base_url = _apk_state.get(uid, {}).get("base_url", "")
    host     = _apk_state.get(uid, {}).get("host", "")
    _apk_state.pop(uid, None)
    chat_id  = message.chat.id

    wait_msg = bot.send_message(
        chat_id,
        "⏳🔍 <b>API Key test করছি...</b>\n"
        "<i>Common endpoints probe করছি, একটু অপেক্ষা করো...</i>",
        parse_mode="HTML",
    )

    def _do():
        panel_id   = f"apk{int(time.time()) % 100000}"
        det_path, det_param = _api_key_test(base_url, api_key)
        try:
            bot.delete_message(chat_id, wait_msg.message_id)
        except Exception:
            pass

        if not det_path:
            # Force-add option — user may know their endpoint
            _apk_state[uid] = {}
            markup = types.InlineKeyboardMarkup(row_width=1)
            markup.add(
                types.InlineKeyboardButton(
                    "⚠️ Force Add করো (Endpoint Manual দেব)",
                    callback_data=f"apkforce:{panel_id}|{base_url}|{api_key}",
                ),
                types.InlineKeyboardButton("❌ বাদ দাও", callback_data=f"apkforce_cancel"),
            )
            bot.send_message(
                chat_id,
                "⚠️ <b>API Endpoint auto-detect হলো না!</b>\n\n"
                f"🌐 Host: <code>{host}</code>\n"
                f"🗝️ Key: <code>{api_key[:8]}...</code>\n\n"
                "সম্ভাব্য কারণ:\n"
                "• এই panel-এ API নেই\n"
                "• API key ভুল\n"
                "• Panel-এর custom endpoint আছে\n\n"
                "তবুও force add করতে চাইলে, পরে <b>/editpanel</b> দিয়ে endpoint দিতে পারবে।",
                reply_markup=markup,
                parse_mode="HTML",
            )
            return

        panel = {
            "id": panel_id,
            "host": host,
            "base_url": base_url,
            "url_hint": f"{base_url}{det_path}",
            "username": f"api:{host}",
            "password": "",
            "api_key": api_key,
            "api_key_param": det_param,
            "engine": "api_key",
            "data_path": det_path,
            "admin_id": uid,
        }
        _dynamic_panels.append(panel)
        save_dynamic_panels()
        _start_dynamic_panel(panel)

        bot.send_message(
            chat_id,
            f"✅🔥 <b>API KEY PANEL ADDED!</b> 🔥✅\n"
            f"⚡━━━━━━━━━━━━━━━━⚡\n\n"
            f"🆔 <b>ID       ▸▸</b> <code>{panel_id}</code>\n"
            f"🌐 <b>Host     ▸▸</b> <code>{host}</code>\n"
            f"🗝️ <b>API Key  ▸▸</b> <code>{api_key[:12]}...</code>\n"
            f"📂 <b>Endpoint ▸▸</b> <code>{det_path}</code>\n"
            f"🔐 <b>Auth     ▸▸</b> <code>{det_param}</code>\n\n"
            f"📡 Monitor thread started! /panels দিয়ে status দেখো।",
            parse_mode="HTML",
        )

    threading.Thread(target=_do, daemon=True).start()


# ── IVA SMS cookie update command ─────────────────────────────────────────────

_iva_cookie_update_state: dict = {}


def _iva_find_panel(panel_id=None):
    """Find any iva_sms panel — checks dynamic_panels AND _BUILTIN_PANELS."""
    all_panels = list(_dynamic_panels) + [p for p in _BUILTIN_PANELS if p not in _dynamic_panels]
    for p in all_panels:
        if p.get("engine") == "iva_sms" and (not panel_id or p["id"] == panel_id):
            return p
    return None


@bot.message_handler(commands=["ivacookie"])
def _iva_cookie_cmd(message):
    uid = message.from_user.id
    if uid not in ADMIN_IDS:
        return
    args = message.text.split()[1:] if message.text else []
    panel_id = args[0] if args else None

    iva_panel = _iva_find_panel(panel_id)

    if not iva_panel:
        bot.send_message(message.chat.id,
            "❌ <b>IVA SMS panel paoa jai nai.</b>\n"
            "Bot restart koro — bp10 auto-load hobe.",
            parse_mode="HTML")
        return

    _iva_cookie_update_state[uid] = iva_panel["id"]
    msg = bot.send_message(
        message.chat.id,
        f"🍪 <b>IVA SMS — Cookie Login</b>\n"
        f"Panel ID: <code>{iva_panel['id']}</code>\n\n"
        f"📋 <b>Steps:</b>\n"
        f"1. Chrome/Firefox-এ <a href='https://ivasms.com/portal/login'>ivasms.com</a> login করো\n"
        f"2. F12 → Application → Cookies → ivasms.com\n"
        f"3. <code>laravel_session</code> value copy করো\n"
        f"4. নিচে paste করো:\n\n"
        f"<code>laravel_session=XXXXXXX</code>\n\n"
        f"<i>(cf_clearance থাকলে সেটাও add করো: <code>cf_clearance=XXX; laravel_session=XXX</code>)</i>",
        reply_markup=_back_admin_kb(),
        parse_mode="HTML",
        disable_web_page_preview=True,
    )
    bot.register_next_step_handler(msg, _iva_cookie_update_step)


def _iva_cookie_update_step(message):
    uid = message.from_user.id
    if _is_back(message.text):
        _iva_cookie_update_state.pop(uid, None)
        _go_admin_panel(message)
        return
    panel_id = _iva_cookie_update_state.pop(uid, None)
    if not panel_id:
        return
    cookie_str = (message.text or "").strip()
    if not cookie_str or "=" not in cookie_str:
        bot.send_message(message.chat.id, "❌ Valid cookie format dao (laravel_session=XXX).", parse_mode="HTML")
        return

    # Update in dynamic_panels first
    updated = False
    for p in _dynamic_panels:
        if p["id"] == panel_id:
            p["cookie_str"] = cookie_str
            save_dynamic_panels()
            updated = True
            break

    # Also update BUILTIN_PANELS in-memory (so _iva_login picks it up)
    for p in _BUILTIN_PANELS:
        if p["id"] == panel_id:
            p["cookie_str"] = cookie_str
            updated = True
            break

    if not updated:
        bot.send_message(message.chat.id, "❌ Panel paoa jai nai.", parse_mode="HTML")
        return

    _iva_scrapers.pop(panel_id, None)  # force re-login with new cookie

    wait_msg = bot.send_message(message.chat.id,
        "⏳ <b>Notun cookie diye login korchi...</b>", parse_mode="HTML")

    def _try_reconnect():
        panel = _iva_find_panel(panel_id)
        ok = _iva_login(panel) if panel else False
        try:
            bot.delete_message(message.chat.id, wait_msg.message_id)
        except Exception:
            pass
        if ok:
            bot.send_message(message.chat.id,
                "✅🔥 <b>IVA SMS — Cookie login SUCCESSFUL!</b>\n"
                "Panel ekhon active — OTP ashle group-e pathabe. 🟢",
                parse_mode="HTML")
        else:
            bot.send_message(message.chat.id,
                "❌ <b>Cookie-o kaj koreni!</b>\n\n"
                "Cookie expire hoye giyeche ba bhul.\n"
                "Abar fresh cookie nao browser theke ar pathao: /ivacookie",
                parse_mode="HTML")

    threading.Thread(target=_try_reconnect, daemon=True).start()


# ── IVA SMS test command (/ivatest) ───────────────────────────────────────────

@bot.message_handler(commands=["ivatest"])
def _iva_test_cmd(message):
    uid = message.from_user.id
    if uid not in ADMIN_IDS:
        return

    # Find ivasms panel (bp10 or any iva_sms engine panel)
    iva_panel = None
    for p in _dynamic_panels:
        if p.get("engine") == "iva_sms":
            iva_panel = p
            break
    # Also check BUILTIN_PANELS
    if not iva_panel:
        for p in _BUILTIN_PANELS:
            if p.get("engine") == "iva_sms":
                iva_panel = p
                break

    if not iva_panel:
        bot.send_message(message.chat.id,
            "❌ <b>IVA SMS panel paoa jai nai!</b>\n"
            "Bot restart koro — bp10 auto-load hobe.",
            parse_mode="HTML")
        return

    wait_msg = bot.send_message(message.chat.id,
        "⏳ <b>ivasms.com theke data fetch korchi...</b>",
        parse_mode="HTML")

    def _do_test():
        try:
            bot.delete_message(message.chat.id, wait_msg.message_id)
        except Exception:
            pass

        records = _iva_fetch(iva_panel)

        if not records:
            bot.send_message(message.chat.id,
                "⚠️ <b>IVA SMS:</b> Ekhon kono OTP nai panel-e.\n"
                "Panel-e kono SMS thakle dekhabe — try koro /ivatest abar kichukhon por.",
                parse_mode="HTML")
            return

        grp = get_otp_group_id()
        items = list(records.values())[:3]

        bot.send_message(message.chat.id,
            f"✅ <b>IVA panel theke {len(records)} ta record paoa geche.</b>\n"
            f"Ekhon <b>{len(items)} ta</b> group-e send korchi...",
            parse_mode="HTML")

        for number, otp, sms_txt, service in items:
            svc = service if service else "IVA"
            if grp:
                send_otp_message(grp, otp, number, 0, svc, sms_txt or "")
            send_otp_message(uid, otp, number, 0, svc, sms_txt or "")

        bot.send_message(message.chat.id,
            f"🔥 <b>Done!</b> {len(items)} ta OTP group+DM-e pathanো hoise.\n"
            f"IVA panel <b>OK</b> ache! 🟢",
            parse_mode="HTML")

    threading.Thread(target=_do_test, daemon=True).start()


# ── Test Panel flow (test without saving) ─────────────────────────────────────

def _tp_get_url(message):
    uid = message.from_user.id
    if uid not in ADMIN_IDS:
        return
    if _is_back(message.text):
        _testpanel_state.pop(uid, None)
        _go_admin_panel(message)
        return
    if _intercept_menu_btn(message):
        return
    url = (message.text or "").strip()
    base_url = _extract_panel_base_url(url) if re.match(r"https?://", url, re.IGNORECASE) else None
    if not base_url:
        msg = bot.send_message(
            message.chat.id,
            "❌ <b>Valid URL dao!</b>\n\nExample: <code>http://1.2.3.4/konekta/agent/SMSCDRReports</code>",
            reply_markup=_back_admin_kb(),
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, _tp_get_url)
        return
    _testpanel_state[uid]["data"]["base_url"] = base_url
    _testpanel_state[uid]["data"]["url_hint"] = url
    msg = bot.send_message(
        message.chat.id,
        f"✅ <b>URL:</b> <code>{base_url}</code>\n\n👤 Username pathao:",
        reply_markup=_back_admin_kb(),
        parse_mode="HTML",
    )
    bot.register_next_step_handler(msg, _tp_get_user)


def _tp_get_user(message):
    uid = message.from_user.id
    if uid not in ADMIN_IDS:
        return
    if _is_back(message.text):
        _testpanel_state.pop(uid, None)
        _go_admin_panel(message)
        return
    if _intercept_menu_btn(message):
        return
    username = (message.text or "").strip()
    if not username:
        msg = bot.send_message(message.chat.id, "❌ Username dao:", reply_markup=_back_admin_kb())
        bot.register_next_step_handler(msg, _tp_get_user)
        return
    _testpanel_state[uid]["data"]["username"] = username
    msg = bot.send_message(
        message.chat.id,
        f"✅ Username: <code>{username}</code>\n\n🔑 Password pathao:",
        reply_markup=_back_admin_kb(),
        parse_mode="HTML",
    )
    bot.register_next_step_handler(msg, _tp_get_pass_test)


def _tp_get_pass_test(message):
    uid = message.from_user.id
    if uid not in ADMIN_IDS:
        return
    if _is_back(message.text):
        _testpanel_state.pop(uid, None)
        _go_admin_panel(message)
        return
    if _intercept_menu_btn(message):
        return
    password = (message.text or "").strip()
    if not password:
        msg = bot.send_message(message.chat.id, "❌ Password dao:", reply_markup=_back_admin_kb())
        bot.register_next_step_handler(msg, _tp_get_pass_test)
        return
    data = _testpanel_state.get(uid, {}).get("data", {})
    wait_msg = bot.send_message(
        message.chat.id,
        "⏳🔍 <b>Testing panel...</b>\n"
        "<i>Login try korchi, token khujchi, data endpoint probe korchi...</i>",
        parse_mode="HTML",
    )
    panel = {
        "id": f"test_{uid}",
        "host": data.get("base_url", ""),
        "base_url": data.get("base_url", ""),
        "url_hint": data.get("url_hint", ""),
        "username": data.get("username", ""),
        "password": password,
        "engine": "ints_smscdr",
        "data_path": "/agent/res/data_smscdr.php",
    }

    def _do_test():
        sess, token, det_engine, det_path = _universal_login(panel)
        try:
            bot.delete_message(message.chat.id, wait_msg.message_id)
        except Exception:
            pass

        if not sess:
            bot.send_message(
                message.chat.id,
                "❌🔥 <b>TEST FAILED!</b> 🔥❌\n\n"
                "⚡━━━━━━━━━━━━━━━━⚡\n\n"
                f"🌐 <b>URL      ▸▸</b> <code>{data.get('base_url','')}</code>\n"
                f"👤 <b>User     ▸▸</b> <code>{data.get('username','')}</code>\n"
                f"📡 <b>Status   ▸▸</b> ❌ Login başarısız\n\n"
                "❌ <b>Possible karon:</b>\n"
                "• URL thik nei\n"
                "• Username/password vul\n"
                "• Panel offline ache",
                parse_mode="HTML",
                reply_markup=_back_admin_kb(),
            )
            _testpanel_state.pop(uid, None)
            return

        # ── Login success — now fetch existing OTPs ───────────────────────────
        engine_label = {
            "ints_smscdr":    "✅ INTS — SMSCDRStats",
            "ints_smsranges": "✅ INTS — SMSRanges",
            "xisora":         "✅ Xisora",
            "html_scrape":    "✅ HTML Scrape",
        }.get(det_engine or "", f"✅ {det_engine or 'Auto'}")
        tok_display = f"<code>{token[:12]}...</code>" if token else "<i>cookie-based</i>"

        # Update panel with detected engine/path and store session
        panel["engine"] = det_engine or "ints_smscdr"
        panel["data_path"] = det_path or "/agent/res/data_smscdr.php"
        _dynamic_sessions[panel["id"]] = {"session": sess, "token": token}

        fetch_msg = bot.send_message(
            message.chat.id,
            "⏳ <b>Login OK!</b> SMS report theke OTP fetch korchi...",
            parse_mode="HTML",
        )

        found_otps = _universal_fetch(panel)

        try:
            bot.delete_message(message.chat.id, fetch_msg.message_id)
        except Exception:
            pass

        # Clean up temp session
        _dynamic_sessions.pop(panel["id"], None)

        # ── Send up to 6 OTPs to admin's configured group ────────────────────
        admin_group_id = get_admin_setting(uid, "otp_group_id", None)
        target_group = admin_group_id or get_otp_group_id()

        sent_count = 0
        MAX_SEND = 6
        otp_items = list(found_otps.values())  # [(number, otp, sms_txt, service)]

        if otp_items and target_group:
            for number, otp_val, sms_txt, service in otp_items[:MAX_SEND]:
                try:
                    send_otp_message(target_group, otp_val, number, "—", service, sms_txt or "")
                    sent_count += 1
                    time.sleep(0.4)
                except Exception:
                    pass

        # ── Summary message to admin ──────────────────────────────────────────
        total_found = len(otp_items)
        if total_found == 0:
            otp_summary = "⚠️ <i>Panel e aj kono OTP record nei (empty).</i>"
        elif not target_group:
            otp_summary = (
                f"⚠️ <b>{total_found}টা OTP</b> panel e ache kintu kono group configure kora nei!\n"
                "Settings theke group set koro."
            )
        else:
            otp_summary = (
                f"📤 <b>{sent_count}টা OTP</b> group e send kora hoise "
                f"({total_found}টার moddhe theke)."
            )

        bot.send_message(
            message.chat.id,
            "✅🔍 <b>TEST SUCCESS!</b> 🔍✅\n\n"
            "⚡━━━━━━━━━━━━━━━━⚡\n\n"
            f"🌐 <b>URL      ▸▸</b> <code>{data.get('base_url','')}</code>\n"
            f"👤 <b>User     ▸▸</b> <code>{data.get('username','')}</code>\n"
            f"🔍 <b>Engine   ▸▸</b> {engine_label}\n"
            f"📂 <b>Endpoint ▸▸</b> <code>{det_path or '/agent/res/data_smscdr.php'}</code>\n"
            f"🔑 <b>Token    ▸▸</b> {tok_display}\n\n"
            "⚡━━━━━━━━━━━━━━━━⚡\n\n"
            f"{otp_summary}\n\n"
            "✅ <i>Panel thik ache! Add Panel diye save korte paro.</i>",
            parse_mode="HTML",
            reply_markup=_back_admin_kb(),
        )
        _testpanel_state.pop(uid, None)

    threading.Thread(target=_do_test, daemon=True).start()


def _svc_get_label(message):
    if message.from_user.id not in ADMIN_IDS:
        return
    if _is_back(message.text):
        _addservice_state.pop(message.from_user.id, None)
        _go_admin_panel(message)
        return
    if _intercept_menu_btn(message):
        return
    label = (message.text or "").strip()
    if not label:
        msg = bot.send_message(message.chat.id, "❌ Label dao:", reply_markup=_back_admin_kb())
        bot.register_next_step_handler(msg, _svc_get_label)
        return
    _addservice_state[message.from_user.id]["label"] = label
    msg = bot.send_message(
        message.chat.id,
        f"✅ Label: <b>{label}</b>\n\n"
        "🔑 <b>Step 2/2:</b> Internal key dao (lowercase, no space)\n"
        "<i>Example: telegram, binance, tiktok</i>",
        reply_markup=_back_admin_kb(),
        parse_mode="HTML",
    )
    bot.register_next_step_handler(msg, _svc_get_key)


def _svc_get_key(message):
    if message.from_user.id not in ADMIN_IDS:
        return
    if _is_back(message.text):
        _addservice_state.pop(message.from_user.id, None)
        _go_admin_panel(message)
        return
    if _intercept_menu_btn(message):
        return
    key = (message.text or "").strip().lower()
    if not key:
        msg = bot.send_message(message.chat.id, "❌ Key dao:", reply_markup=_back_admin_kb())
        bot.register_next_step_handler(msg, _svc_get_key)
        return
    label = _addservice_state.get(message.from_user.id, {}).get("label", "")
    existing_keys = [s["key"] for s in _services]
    if key in existing_keys:
        msg = bot.send_message(
            message.chat.id,
            f"❌ Key <code>{key}</code> already ache! Onnyo key dao:",
            reply_markup=_back_admin_kb(),
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, _svc_get_key)
        return
    _services.append({"label": label, "key": key})
    save_services()
    _addservice_state.pop(message.from_user.id, None)
    _go_admin_panel(
        message,
        f"✅🔥 <b>Service Added!</b>\n\n"
        f"🏷️ Label: <b>{label}</b>\n"
        f"🔑 Key: <code>{key}</code>\n\n"
        f"<i>Service menu-te dekha jabe!</i>",
    )


@bot.message_handler(commands=["listpanels"])
def listpanels_cmd(message):
    if message.from_user.id not in ADMIN_IDS:
        return
    caller_uid = message.from_user.id
    my_panels = [
        p for p in _dynamic_panels
        if is_super_admin(caller_uid) or p.get("admin_id") == caller_uid
    ]
    if not my_panels:
        bot.send_message(
            message.chat.id,
            "📋 <b>Tomar kono dynamic panel nei.</b>\n💡 /addpanel diye add koro.",
            parse_mode="HTML",
        )
        return
    lines = "📋🔥 <b>DYNAMIC PANELS LIST</b> 🔥📋\n⚡━━━━━━━━━━━━━━━━⚡\n\n"
    for p in my_panels:
        pid = p["id"]
        with _stats_lock:
            s = _panel_stats.get(pid, {})
        st = s.get("status", "⏳")
        lines += (
            f"{st} 🆔 <code>{pid}</code>\n"
            f"   🌐 <code>{p.get('host', '?')}</code>\n"
            f"   👤 {p.get('username', '?')}\n\n"
        )
    lines += "🗑️ Remove: <code>/removepanel [ID]</code>"
    bot.send_message(message.chat.id, lines, parse_mode="HTML")


@bot.message_handler(commands=["removepanel"])
def removepanel_cmd(message):
    if message.from_user.id not in ADMIN_IDS:
        return
    args = message.text.split(maxsplit=1)
    if len(args) < 2:
        bot.send_message(
            message.chat.id,
            "❌ Panel ID dao:\n<code>/removepanel d12345</code>\n\n"
            "💡 /listpanels diye ID dekho.",
            parse_mode="HTML",
        )
        return
    caller_uid = message.from_user.id
    pid = args[1].strip()
    target = next((p for p in _dynamic_panels if p["id"] == pid), None)
    if not target:
        bot.send_message(message.chat.id, f"❌ Panel <code>{pid}</code> not found.\n💡 /listpanels diye ID check koro.", parse_mode="HTML")
        return
    if not is_super_admin(caller_uid) and target.get("admin_id") != caller_uid:
        bot.send_message(message.chat.id, "❌ <b>Ei panel tomar na — remove korte parba na!</b>", parse_mode="HTML")
        return
    _dynamic_panels[:] = [p for p in _dynamic_panels if p["id"] != pid]
    save_dynamic_panels()
    with _stats_lock:
        _panel_stats.pop(pid, None)
    _dynamic_sessions.pop(pid, None)
    _dynamic_locks.pop(pid, None)
    bot.send_message(message.chat.id, f"✅🔥 Panel <code>{pid}</code> removed!\n<i>Monitor thread will stop naturally.</i>", parse_mode="HTML")


@bot.callback_query_handler(func=lambda call: True)
def callback_handler(call):
    global stock
    try:
        data = call.data

        # ── Force Add Panel (Railway IP blocked) ─────────────────────────────
        if data.startswith("forceadd:"):
            if call.from_user.id not in ADMIN_IDS:
                bot.answer_callback_query(call.id, "❌ Permission নেই!")
                return
            pid = data.split(":", 1)[1]
            panel = _pending_force_add.pop(pid, None)
            if not panel:
                bot.answer_callback_query(call.id, "❌ Panel data expire হয়ে গেছে। আবার চেষ্টা করো।")
                return
            _dynamic_panels.append(panel)
            save_dynamic_panels()
            _start_dynamic_panel(panel)
            bot.answer_callback_query(call.id, "✅ Panel Force Add হয়েছে!")
            try:
                bot.edit_message_text(
                    f"✅🔥 <b>PANEL FORCE ADDED!</b>\n\n"
                    f"🆔 <b>ID:</b> <code>{pid}</code>\n"
                    f"🌐 <b>Host:</b> <code>{panel.get('host', '')}</code>\n"
                    f"👤 <b>User:</b> <code>{panel.get('username', '')}</code>\n\n"
                    f"⚠️ Login এখনো verify হয়নি — panel নিজেই login করার চেষ্টা করবে।\n"
                    f"/panels দিয়ে status দেখো।",
                    call.message.chat.id,
                    call.message.message_id,
                    parse_mode="HTML",
                )
            except Exception:
                pass
            return

        if data.startswith("forceadd_cancel:"):
            pid = data.split(":", 1)[1]
            _pending_force_add.pop(pid, None)
            bot.answer_callback_query(call.id, "বাদ দেওয়া হয়েছে।")
            try:
                bot.edit_message_text(
                    "❌ Panel add বাতিল করা হয়েছে।\n/addpanel দিয়ে আবার চেষ্টা করো।",
                    call.message.chat.id,
                    call.message.message_id,
                    parse_mode="HTML",
                )
            except Exception:
                pass
            return

        # ── API Key Panel type selection ──────────────────────────────────────
        if data == "aptype:pass":
            uid = call.from_user.id
            if uid not in ADMIN_IDS:
                bot.answer_callback_query(call.id, "❌ Permission নেই!")
                return
            bot.answer_callback_query(call.id)
            try:
                bot.edit_message_reply_markup(call.message.chat.id, call.message.message_id, reply_markup=None)
            except Exception:
                pass
            _addpanel_state[uid] = {"step": "url", "data": {}}
            msg = bot.send_message(
                call.message.chat.id,
                "🔧🔥 <b>ADD NEW PANEL</b> 🔥🔧\n\n"
                "📡 <b>Step 1/3:</b> Panel-এর URL পাঠাও\n\n"
                "✅ <b>যেকোনো format চলবে:</b>\n"
                "• <code>http://1.2.3.4</code>\n"
                "• <code>http://1.2.3.4/ints</code>\n"
                "• <code>http://1.2.3.4/konekta</code>\n"
                "• <code>https://truesms.net</code>\n\n"
                "🤖 <i>Login, captcha, endpoint সব auto-detect হবে!</i>",
                reply_markup=_back_admin_kb(),
                parse_mode="HTML",
            )
            bot.register_next_step_handler(msg, _ap_get_url)
            return

        if data == "aptype:apikey":
            uid = call.from_user.id
            if uid not in ADMIN_IDS:
                bot.answer_callback_query(call.id, "❌ Permission নেই!")
                return
            bot.answer_callback_query(call.id)
            try:
                bot.edit_message_reply_markup(call.message.chat.id, call.message.message_id, reply_markup=None)
            except Exception:
                pass
            _apk_state[uid] = {}
            msg = bot.send_message(
                call.message.chat.id,
                "🗝️🔥 <b>API KEY দিয়ে PANEL ADD</b> 🔥🗝️\n\n"
                "📡 <b>Step 1/2:</b> Panel-এর URL পাঠাও\n\n"
                "✅ <b>যেকোনো format চলবে:</b>\n"
                "• <code>http://1.2.3.4</code>\n"
                "• <code>http://1.2.3.4/api</code>\n"
                "• <code>https://mypanel.com</code>\n"
                "• <code>https://mypanel.com/api/sms</code>",
                reply_markup=_back_admin_kb(),
                parse_mode="HTML",
            )
            bot.register_next_step_handler(msg, _apk_get_url)
            return

        if data.startswith("apkforce:"):
            uid = call.from_user.id
            if uid not in ADMIN_IDS:
                bot.answer_callback_query(call.id, "❌ Permission নেই!")
                return
            bot.answer_callback_query(call.id, "✅ Force Add হচ্ছে...")
            try:
                rest = data[len("apkforce:"):].split("|", 2)
                panel_id = rest[0]
                base_url = rest[1] if len(rest) > 1 else ""
                api_key  = rest[2] if len(rest) > 2 else ""
            except Exception:
                bot.send_message(call.message.chat.id, "❌ Data parse error।", parse_mode="HTML")
                return
            host_m = re.search(r"//([^/]+)", base_url)
            host   = host_m.group(1) if host_m else base_url
            panel = {
                "id": panel_id,
                "host": host,
                "base_url": base_url,
                "url_hint": f"{base_url}/api/sms",
                "username": f"api:{host}",
                "password": "",
                "api_key": api_key,
                "api_key_param": "api_key",
                "engine": "api_key",
                "data_path": "/api/sms",
                "admin_id": uid,
            }
            _dynamic_panels.append(panel)
            save_dynamic_panels()
            _start_dynamic_panel(panel)
            try:
                bot.edit_message_text(
                    f"✅ <b>API KEY PANEL FORCE ADDED!</b>\n\n"
                    f"🆔 <b>ID:</b> <code>{panel_id}</code>\n"
                    f"🌐 <b>Host:</b> <code>{host}</code>\n"
                    f"🗝️ <b>Key:</b> <code>{api_key[:12]}...</code>\n\n"
                    f"⚠️ Endpoint auto-detect হয়নি — default <code>/api/sms</code> দিয়ে চলবে।\n"
                    f"/panels দিয়ে status দেখো।",
                    call.message.chat.id,
                    call.message.message_id,
                    parse_mode="HTML",
                )
            except Exception:
                pass
            return

        if data == "apkforce_cancel":
            bot.answer_callback_query(call.id, "বাদ দেওয়া হয়েছে।")
            try:
                bot.edit_message_text(
                    "❌ API Key panel add বাতিল।\n/addpanel দিয়ে আবার চেষ্টা করো।",
                    call.message.chat.id,
                    call.message.message_id,
                    parse_mode="HTML",
                )
            except Exception:
                pass
            return
        # ─────────────────────────────────────────────────────────────────────

        if data == "v":
            uid = call.from_user.id

            grp_id = get_otp_group_id()
            grp_link = get_otp_group_link()
            ch2_link = get_channel2()
            ch2_ref = _extract_username(ch2_link)

            not_joined = []

            grp_ok = _check_member(grp_id, uid) if grp_id else None
            if grp_ok is False:
                not_joined.append(("🔥 OTP Group", grp_link))

            ch2_ok = _check_member(ch2_ref, uid) if ch2_ref else None
            if ch2_ok is False:
                not_joined.append(("📢 Main Channel", ch2_link))

            if not_joined:
                bot.answer_callback_query(call.id, "❌ Sob jagay join hao nai!", show_alert=False)
                lines = "❌ <b>Verify hote parcho na!</b>\n\n"
                lines += "⛔ Tumi ekhono nicher jagay join hao nai:\n\n"
                for name, _ in not_joined:
                    lines += f"  🚫 <b>{name}</b>\n"
                lines += "\n👇 Join kore <b>Verify Koro</b> te click koro:"
                err_markup = types.InlineKeyboardMarkup(row_width=1)
                for name, lnk in not_joined:
                    err_markup.add(types.InlineKeyboardButton(
                        f"👉 {name}-e JOIN KORO", url=lnk
                    ))
                err_markup.add(types.InlineKeyboardButton(
                    "🔄 Verify Koro", callback_data="v"
                ))
                try:
                    bot.edit_message_text(
                        lines,
                        call.message.chat.id,
                        call.message.message_id,
                        reply_markup=err_markup,
                        parse_mode="HTML",
                    )
                except Exception:
                    bot.send_message(
                        call.message.chat.id,
                        lines,
                        reply_markup=err_markup,
                        parse_mode="HTML",
                    )
            else:
                bot.delete_message(call.message.chat.id, call.message.message_id)
                vname = call.from_user.first_name or call.from_user.username or "User"
                bot.send_message(
                    call.message.chat.id,
                    get_template("verify_success").format(vname=vname, uid=uid),
                    reply_markup=main_menu(call.from_user.id),
                    parse_mode="HTML",
                )

        elif data == "back_to_services":
            show_services(call.message)

        elif data.startswith("s:"):
            svc = data.split(":")[1]
            markup = types.InlineKeyboardMarkup(row_width=2)
            btns = []
            if svc in stock:
                for cnt, nums in stock[svc].items():
                    if nums:
                        _, flag = get_country_details(nums[0])
                        btns.append(
                            types.InlineKeyboardButton(
                                f" {flag} {cnt}", callback_data=f"n:{svc}:{cnt}"
                            )
                        )
            if btns:
                markup.add(*btns)
            markup.add(
                types.InlineKeyboardButton("⬅️ 𝗕𝗮𝗰𝗸", callback_data="back_to_services")
            )
            bot.edit_message_text(
                f"🔥 <b>{svc.upper()} — COUNTRY</b> 🔥",
                call.message.chat.id,
                call.message.message_id,
                reply_markup=markup,
                parse_mode="HTML",
            )

        elif data.startswith("n:"):
            _, svc, scnt = data.split(":")
            if scnt in stock.get(svc, {}) and stock[svc][scnt]:
                num = stock[svc][scnt].pop(0)
                save_stock()
                c_name, flag = get_country_details(num)
                uid_n = call.from_user.id
                # Release any previously assigned number for this user — delete it permanently
                with user_map_lock:
                    old_nums = [k for k, v in user_map.items() if v == uid_n]
                    for old_clean in old_nums:
                        user_map.pop(old_clean, None)
                        assigned_time.pop(old_clean, None)
                if old_nums:
                    _save_user_map()
                    print(f"[N:] Deleted old number(s) {old_nums} for user {uid_n}")
                register_number(call.message.chat.id, num)
                display_num = num if num.startswith("+") else "+" + num
                init_kb = types.InlineKeyboardMarkup(row_width=2)
                init_kb.add(
                    types.InlineKeyboardButton("🔄 𝗚𝗲𝘁 𝗡𝗲𝘄 𝗡𝘂𝗺𝗯𝗲𝗿", callback_data=f"n:{svc}:{scnt}"),
                    types.InlineKeyboardButton("🌍 𝗖𝗵𝗮𝗻𝗴𝗲 𝗖𝗼𝘂𝗻𝘁𝗿𝘆", callback_data=f"s:{svc}"),
                )
                if get_otp_group_link():
                    init_kb.add(
                        types.InlineKeyboardButton("📢 𝗢𝗧𝗣 𝗚𝗿𝗼𝘂𝗽", url=get_otp_group_link()),
                    )
                res = get_template("number_assigned").format(
                    svc=svc.capitalize(), flag=flag, country=c_name, number=display_num
                )
                # Track service/country for this user so OTP message buttons work
                _user_last_svc[uid_n] = (svc, scnt)
                tracked_num_msg = _user_last_num_msg.get(uid_n)
                clicked_msg_id = call.message.message_id
                if tracked_num_msg and clicked_msg_id == tracked_num_msg:
                    # Clicked from "Number Assigned" message → delete it
                    try:
                        bot.delete_message(
                            chat_id=call.message.chat.id,
                            message_id=clicked_msg_id,
                        )
                    except Exception:
                        pass
                else:
                    # Clicked from OTP code message → keep it, only strip buttons
                    try:
                        bot.edit_message_reply_markup(
                            chat_id=call.message.chat.id,
                            message_id=clicked_msg_id,
                            reply_markup=None,
                        )
                    except Exception:
                        pass
                    # Delete any separately tracked "Number Assigned" message
                    if tracked_num_msg:
                        try:
                            bot.delete_message(
                                chat_id=call.message.chat.id,
                                message_id=tracked_num_msg,
                            )
                        except Exception:
                            pass
                # Send fresh number-assigned message
                new_msg = bot.send_message(
                    call.message.chat.id,
                    res,
                    reply_markup=init_kb,
                    parse_mode="HTML",
                )
                # Track this new message so OTP arrival can delete it
                _user_last_num_msg[uid_n] = new_msg.message_id
                _start_countdown(
                    call.message.chat.id,
                    new_msg.message_id,
                    svc, flag, c_name, display_num, scnt,
                )
            else:
                bot.answer_callback_query(call.id, " STOCK SHESH! ", show_alert=True)

        elif data == "clr_menu":
            if call.from_user.id not in ADMIN_IDS:
                return
            bot.edit_message_text(
                "🗑️🔥 <b>STOCK CLEAR PANEL</b> 🔥🗑️\n\n"
                " <b>Kon service-er stock clear korbe?</b>\n"
                "⬇️ Service choose koro:",
                call.message.chat.id,
                call.message.message_id,
                reply_markup=_clr_service_markup(),
                parse_mode="HTML",
            )

        elif data.startswith("clr_s:"):
            if call.from_user.id not in ADMIN_IDS:
                return
            svc = data[6:]
            markup = types.InlineKeyboardMarkup(row_width=1)
            svc_stock = stock.get(svc, {})
            has_any = False
            for cnt, nums in svc_stock.items():
                if nums:
                    has_any = True
                    _, flag = get_country_details(nums[0])
                    cb = f"clr_c:{svc}:{cnt}"
                    if len(cb.encode()) <= 64:
                        markup.add(
                            types.InlineKeyboardButton(
                                f"🗑️ {flag} {cnt}  ({len(nums)} টি)", callback_data=cb
                            )
                        )
            if not has_any:
                markup.add(
                    types.InlineKeyboardButton("⚠️ Stock nai!", callback_data="clr_menu")
                )
            markup.add(types.InlineKeyboardButton("⬅️ Back", callback_data="clr_menu"))
            bot.edit_message_text(
                f"🔥 <b>{svc.upper()} — Kon desh clear korbe?</b> 🔥\n\n"
                f"⬇️ Country choose koro:",
                call.message.chat.id,
                call.message.message_id,
                reply_markup=markup,
                parse_mode="HTML",
            )

        elif data.startswith("clr_c:"):
            if call.from_user.id not in ADMIN_IDS:
                return
            _, svc, cnt = data.split(":", 2)
            count = len(stock.get(svc, {}).get(cnt, []))
            _, flag = get_country_details(stock[svc][cnt][0]) if count else ("", "🌐")
            markup = types.InlineKeyboardMarkup(row_width=2)
            markup.add(
                types.InlineKeyboardButton(
                    "✅ Haa, Delete Koro", callback_data=f"clr_y:{svc}:{cnt}"
                ),
                types.InlineKeyboardButton("❌ Cancel", callback_data=f"clr_s:{svc}"),
            )
            bot.edit_message_text(
                f"⚠️ <b>CONFIRM DELETE</b> ⚠️\n\n"
                f"💬 <b>Service ▸▸</b>  {svc.upper()}\n"
                f"🌍 <b>Country ▸▸</b>  {flag} {cnt}\n"
                f"📱 <b>Numbers ▸▸</b>  {count} টি\n\n"
                f" Sure? Ei {count} টি number delete hoye jabe!",
                call.message.chat.id,
                call.message.message_id,
                reply_markup=markup,
                parse_mode="HTML",
            )

        elif data.startswith("clr_y:"):
            if call.from_user.id not in ADMIN_IDS:
                return
            _, svc, cnt = data.split(":", 2)
            removed = len(stock.get(svc, {}).get(cnt, []))
            if svc in stock and cnt in stock[svc]:
                del stock[svc][cnt]
                save_stock()
            markup = types.InlineKeyboardMarkup(row_width=2)
            markup.add(
                types.InlineKeyboardButton("🗑️ Aro Clear", callback_data=f"clr_s:{svc}"),
                types.InlineKeyboardButton("🔙 Services", callback_data="clr_menu"),
            )
            bot.edit_message_text(
                f"✅🔥 <b>DELETE COMPLETE!</b> 🔥✅\n\n"
                f"💬 <b>Service ▸▸</b>  {svc.upper()}\n"
                f"🌍 <b>Country ▸▸</b>  {cnt}\n"
                f"📱 <b>Deleted  ▸▸</b>  {removed} টি number\n\n"
                f"⚡ <i>Stock update hoyeche!</i>",
                call.message.chat.id,
                call.message.message_id,
                reply_markup=markup,
                parse_mode="HTML",
            )

        elif data == "clr_all":
            if call.from_user.id not in ADMIN_IDS:
                return
            total = sum(
                len(nums) for svc_d in stock.values() for nums in svc_d.values()
            )
            markup = types.InlineKeyboardMarkup(row_width=2)
            markup.add(
                types.InlineKeyboardButton(
                    " Haa, SOB Clear", callback_data="clr_allok"
                ),
                types.InlineKeyboardButton("❌ Cancel", callback_data="clr_menu"),
            )
            bot.edit_message_text(
                f"☠️⚠️ <b>CLEAR ALL CONFIRM</b> ⚠️☠️\n\n"
                f" Total <b>{total} টি</b> number delete hobe!\n"
                f"⚡ Sob service-er sob country mochhe jabe!\n\n"
                f"🔥 Sure? Eta undo kora jabe na!",
                call.message.chat.id,
                call.message.message_id,
                reply_markup=markup,
                parse_mode="HTML",
            )

        elif data == "clr_allok":
            if call.from_user.id not in ADMIN_IDS:
                return
            stock = {
                "whatsapp": {},
                "facebook": {},
                "telegram": {},
                "instagram": {},
                "pc clone": {},
                "binance": {},
            }
            save_stock()
            bot.edit_message_text(
                "🔥 <b>SOB STOCK CLEAR HOYECHE!</b> 🔥\n <i>Ekhon naya number add koro!</i> ",
                call.message.chat.id,
                call.message.message_id,
                parse_mode="HTML",
            )

        elif data.startswith("rmpanel:"):
            caller_uid = call.from_user.id
            if caller_uid not in ADMIN_IDS:
                return
            pid = data.split(":", 1)[1]
            target = next((p for p in _dynamic_panels if p["id"] == pid), None)
            if not target:
                bot.answer_callback_query(call.id, "❌ Panel pawa jaini!", show_alert=True)
            elif not is_super_admin(caller_uid) and target.get("admin_id") != caller_uid:
                bot.answer_callback_query(call.id, "❌ Ei panel tomar na!", show_alert=True)
            else:
                _dynamic_panels[:] = [p for p in _dynamic_panels if p["id"] != pid]
                save_dynamic_panels()
                with _stats_lock:
                    _panel_stats.pop(pid, None)
                _dynamic_sessions.pop(pid, None)
                _dynamic_locks.pop(pid, None)
                try:
                    bot.edit_message_text(
                        f"✅🔥 <b>Panel <code>{pid}</code> removed!</b>\n"
                        f"<i>Monitor thread will stop naturally.</i>",
                        call.message.chat.id,
                        call.message.message_id,
                        parse_mode="HTML",
                    )
                except Exception:
                    pass

        elif data.startswith("rmsvc:"):
            if call.from_user.id not in ADMIN_IDS:
                return
            key = data.split(":", 1)[1]
            before = len(_services)
            _services[:] = [s for s in _services if s["key"] != key]
            if len(_services) < before:
                save_services()
                bot.edit_message_text(
                    f"✅🔥 <b>Service <code>{key}</code> removed!</b>\n"
                    f"<i>Service menu theke hatiye dewa hoyeche.</i>",
                    call.message.chat.id,
                    call.message.message_id,
                    parse_mode="HTML",
                )
            else:
                bot.answer_callback_query(call.id, "❌ Service pawa jaini!", show_alert=True)

        elif data.startswith("aadur:"):
            if not is_super_admin(call.from_user.id):
                bot.answer_callback_query(call.id, "❌ Permission nei!", show_alert=True)
                return
            parts = data.split(":")
            new_uid = int(parts[1])
            months = int(parts[2])
            add_admin(new_uid, months=months)
            exp_ts = _admin_expiry.get(str(new_uid))
            exp_str = datetime.datetime.fromtimestamp(exp_ts).strftime("%d %b %Y") if exp_ts else "—"
            raw_n = user_names.get(str(new_uid), "")
            name_str = raw_n if isinstance(raw_n, str) else raw_n.get("first_name", str(new_uid))
            name_str = name_str or str(new_uid)
            try:
                bot.edit_message_text(
                    f"✅ <b>ADMIN ADDED!</b>\n\n"
                    f"👑 <b>New Admin:</b> {name_str} [<code>{new_uid}</code>]\n"
                    f"📅 <b>Meiad:</b> {months} Mash\n"
                    f"🗓️ <b>Expire Date:</b> {exp_str}\n\n"
                    f"<i>Ekhon theke ei user admin panel access pabe.</i>",
                    call.message.chat.id,
                    call.message.message_id,
                    parse_mode="HTML",
                )
            except Exception:
                pass
            bot.answer_callback_query(call.id, f"✅ Admin added ({months} mash)!", show_alert=False)
            try:
                bot.send_message(
                    new_uid,
                    f"🎉 <b>Congratulations! Tumi Admin hoyecho!</b>\n\n"
                    f"📅 <b>Admin Meiad:</b> {months} Mash\n"
                    f"🗓️ <b>Expire:</b> {exp_str}\n\n"
                    f"Admin panel access er jonno /admin command dao.",
                    parse_mode="HTML",
                )
            except Exception:
                pass

        elif data == "aadur_cancel":
            bot.answer_callback_query(call.id, "❌ Cancel kora hoyeche.")
            try:
                bot.delete_message(call.message.chat.id, call.message.message_id)
            except Exception:
                pass

        elif data.startswith("rmadmin:"):
            if not is_super_admin(call.from_user.id):
                bot.answer_callback_query(call.id, "❌ Shudhu Super Admin remove korte parbe!", show_alert=True)
                return
            target = int(data.split(":")[1])
            if remove_admin(target):
                raw_n = user_names.get(str(target), "")
                name = raw_n if isinstance(raw_n, str) else raw_n.get("first_name", str(target))
                name = name or str(target)
                bot.answer_callback_query(call.id, f"✅ {name} removed!", show_alert=False)
                try:
                    bot.edit_message_text(
                        f"✅ <b>ADMIN REMOVED!</b>\n\n"
                        f"🗑️ <b>Removed:</b> {name} [<code>{target}</code>]\n\n"
                        f"<i>Ekhon theke ei user admin access harabe.</i>",
                        call.message.chat.id,
                        call.message.message_id,
                        parse_mode="HTML",
                    )
                except Exception:
                    pass
            else:
                bot.answer_callback_query(call.id, "❌ Remove kora gelo na (Super Admin)!", show_alert=True)

        elif data.startswith("cfg_toggle:"):
            if call.from_user.id not in ADMIN_IDS:
                return
            parts = data.split(":")
            try:
                cid = int(parts[1])
                action = parts[2]
            except (IndexError, ValueError):
                bot.answer_callback_query(call.id, "❌ Error!", show_alert=True)
                return
            with _demo_lock:
                for cfg in _demo_configs:
                    if cfg["id"] == cid:
                        cfg["active"] = (action == "start")
                        cfg_name = cfg["name"]
                        break
                else:
                    bot.answer_callback_query(call.id, "❌ Config পাওয়া যায়নি!", show_alert=True)
                    return
            if action == "start":
                _demo_next_fire[cid] = 0
                status_msg = f"🟢 <b>{cfg_name} চালু হয়েছে!</b>"
            else:
                _demo_next_fire.pop(cid, None)
                status_msg = f"🔴 <b>{cfg_name} বন্ধ হয়েছে!</b>"
            bot.answer_callback_query(call.id, status_msg.replace("<b>", "").replace("</b>", ""), show_alert=False)
            try:
                bot.edit_message_text(
                    "⚡ <b>Config Start/Stop:</b>",
                    call.message.chat.id,
                    call.message.message_id,
                    reply_markup=demo_cfg_inline_markup(),
                    parse_mode="HTML",
                )
            except Exception:
                pass
            bot.send_message(
                call.message.chat.id,
                status_msg + "\n\n" + demo_status_text(),
                parse_mode="HTML",
            )

        elif data.startswith("rmcfg:"):
            if call.from_user.id not in ADMIN_IDS:
                return
            try:
                cid = int(data.split(":", 1)[1])
            except ValueError:
                bot.answer_callback_query(call.id, "❌ Invalid config!", show_alert=True)
                return
            with _demo_lock:
                before = len(_demo_configs)
                _demo_configs[:] = [c for c in _demo_configs if c["id"] != cid]
                removed = before > len(_demo_configs)
            if removed:
                _demo_next_fire.pop(cid, None)
                try:
                    bot.edit_message_text(
                        f"✅🔥 <b>Config মুছে গেছে!</b>\n\n" + demo_status_text(),
                        call.message.chat.id,
                        call.message.message_id,
                        parse_mode="HTML",
                    )
                except Exception:
                    pass
            else:
                bot.answer_callback_query(call.id, "❌ Config পাওয়া যায়নি!", show_alert=True)

        elif data.startswith("editmsg:"):
            if call.from_user.id not in ADMIN_IDS:
                return
            key = data.split(":", 1)[1]
            if key in _TEMPLATE_LABELS:
                _ask_new_template(call, key)
            else:
                bot.answer_callback_query(call.id, "❌ Unknown template!", show_alert=True)

        elif data.startswith("editmsg_reset:"):
            if call.from_user.id not in ADMIN_IDS:
                return
            key = data.split(":", 1)[1]
            if key in _DEFAULT_TEMPLATES:
                _templates[key] = _DEFAULT_TEMPLATES[key]
                save_templates()
                bot.answer_callback_query(call.id, f"✅ '{key}' Default এ Reset হয়েছে!", show_alert=True)
            else:
                bot.answer_callback_query(call.id, "❌ Unknown template!", show_alert=True)

        elif data == "editmsg_reset_all":
            if call.from_user.id not in ADMIN_IDS:
                return
            _templates.update(_DEFAULT_TEMPLATES)
            save_templates()
            try:
                bot.edit_message_text(
                    "✅🔥 <b>সব মেসেজ Default এ Reset হয়েছে!</b>\n\n"
                    "<i>এখন থেকে সব মেসেজ Default ফরমেটে যাবে।</i>",
                    call.message.chat.id,
                    call.message.message_id,
                    parse_mode="HTML",
                )
            except Exception:
                pass

        elif data == "grp_info":
            if call.from_user.id not in ADMIN_IDS:
                return
            _show_settings_inline(call)

        elif data == "set_autodel":
            if call.from_user.id not in ADMIN_IDS:
                return
            cur = _group_settings.get("auto_delete", True)
            _group_settings["auto_delete"] = not cur
            save_group_settings()
            bot.answer_callback_query(
                call.id,
                "✅ Auto Delete: " + ("🟢 ON" if not cur else "🔴 OFF"),
                show_alert=False,
            )
            _show_settings_inline(call)

        elif data == "toggle_grp_send":
            if call.from_user.id not in ADMIN_IDS:
                return
            cur = _group_settings.get("group_otp_send", True)
            _group_settings["group_otp_send"] = not cur
            save_group_settings()
            new_state = not cur
            bot.answer_callback_query(
                call.id,
                "✅ Group OTP Send: " + ("🟢 ON — Group-e OTP jabey" if new_state else "🔴 OFF — Shudhu Inbox-e jabey"),
                show_alert=True,
            )
            _show_settings_inline(call)

        elif data == "set_channel2":
            if call.from_user.id not in ADMIN_IDS:
                return
            bot.answer_callback_query(call.id)
            msg = bot.send_message(
                call.message.chat.id,
                "📢 <b>Notun Join Channel link dao:</b>\n\n"
                "<i>Example: https://t.me/aR_OTP_rcv</i>",
                reply_markup=_back_admin_kb(),
                parse_mode="HTML",
            )
            bot.register_next_step_handler(msg, _sett_get_channel2)

        elif data == "set_botlink":
            if call.from_user.id not in ADMIN_IDS:
                return
            bot.answer_callback_query(call.id)
            msg = bot.send_message(
                call.message.chat.id,
                "🤖 <b>Notun Bot link dao:</b>\n\n"
                "<i>Example: https://t.me/ar_otp_bot</i>",
                reply_markup=_back_admin_kb(),
                parse_mode="HTML",
            )
            bot.register_next_step_handler(msg, _sett_get_botlink)

        elif data == "grp_setlink":
            if call.from_user.id not in ADMIN_IDS:
                return
            bot.answer_callback_query(call.id)
            msg = bot.send_message(
                call.message.chat.id,
                "🔗 <b>Notun OTP Group Link dao:</b>\n\n"
                "<i>Example: https://t.me/aR_OTP_rcv</i>",
                reply_markup=_back_admin_kb(),
                parse_mode="HTML",
            )
            bot.register_next_step_handler(msg, _grp_get_link)

        elif data == "grp_setid":
            if call.from_user.id not in ADMIN_IDS:
                return
            bot.answer_callback_query(call.id)
            msg = bot.send_message(
                call.message.chat.id,
                "🆔 <b>Notun OTP Group Chat ID dao:</b>\n\n"
                "<i>Example: -1001234567890</i>\n"
                "⚠️ Negative number dite hobe (group ID always negative)",
                reply_markup=_back_admin_kb(),
                parse_mode="HTML",
            )
            bot.register_next_step_handler(msg, _grp_get_id)

        elif data == "grp_remove":
            if call.from_user.id not in ADMIN_IDS:
                return
            markup = types.InlineKeyboardMarkup(row_width=2)
            markup.add(
                types.InlineKeyboardButton("✅ Haa, Remove", callback_data="grp_removeok"),
                types.InlineKeyboardButton("❌ Cancel", callback_data="grp_info"),
            )
            bot.answer_callback_query(call.id)
            bot.edit_message_text(
                "⚠️ <b>CONFIRM GROUP REMOVE</b> ⚠️\n\n"
                "OTP Group setting reset hobe!\n"
                "Group-e aro OTP pathano bondho hobe.\n\n"
                "Sure?",
                call.message.chat.id,
                call.message.message_id,
                reply_markup=markup,
                parse_mode="HTML",
            )

        elif data.startswith("v2svc:"):
            sid = data.split(":", 1)[1]
            services = _fastx_liveaccess()
            svc_data = next((s for s in services if s.get("sid") == sid), None)
            if not svc_data:
                bot.answer_callback_query(call.id, "❌ Service পাওয়া যায়নি!", show_alert=True)
                return
            ranges = svc_data.get("ranges", [])
            markup = types.InlineKeyboardMarkup(row_width=2)
            rng_btns = []
            for rng in ranges:
                prefix = rng.rstrip("X")
                c_name, flag = get_country_details(prefix)
                short = c_name.split()[0] if c_name and c_name != "Unknown" else ""
                label = f"{flag} {short} | {rng}" if short else f"{flag} {rng}"
                rng_btns.append(types.InlineKeyboardButton(
                    label, callback_data=f"v2rng:{prefix}:{sid}"
                ))
            if rng_btns:
                markup.add(*rng_btns)
            markup.add(types.InlineKeyboardButton("⬅️ Back", callback_data="v2back"))
            bot.edit_message_text(
                f"📡 <b>V2 — {_v2_svc_emoji(sid)} {sid}</b>\n\n"
                f"🔢 <b>Range select koro:</b>\n"
                f"<i>এই range-এ live OTP আসছে — click করলে number পাবে</i>",
                call.message.chat.id, call.message.message_id,
                reply_markup=markup, parse_mode="HTML"
            )
            bot.answer_callback_query(call.id)

        elif data.startswith("v2rng:"):
            parts = data.split(":")
            prefix = parts[1] if len(parts) > 1 else ""
            sid = parts[2] if len(parts) > 2 else "?"
            bot.answer_callback_query(call.id, "⏳ Number নিচ্ছি...", show_alert=False)
            num = _fastx_getnum(prefix)
            uid_v2 = call.from_user.id
            if num:
                with user_map_lock:
                    old_nums = [k for k, v in user_map.items() if v == uid_v2]
                    for old_clean in old_nums:
                        user_map.pop(old_clean, None)
                        assigned_time.pop(old_clean, None)
                if old_nums:
                    _save_user_map()
                register_number(uid_v2, num)
                c_name, flag = get_country_details(num)
                display_num = num if num.startswith("+") else "+" + num
                _user_last_svc[uid_v2] = (sid.lower(), c_name)
                refresh_kb = types.InlineKeyboardMarkup(row_width=2)
                refresh_kb.add(
                    types.InlineKeyboardButton("🔄 New Number", callback_data=f"v2rng:{prefix}:{sid}"),
                    types.InlineKeyboardButton("⬅️ Back", callback_data="v2back"),
                )
                if get_otp_group_link():
                    refresh_kb.add(types.InlineKeyboardButton("📢 OTP Group", url=get_otp_group_link()))
                bot.edit_message_text(
                    f"✅ <b>V2 Number পেয়েছ!</b>\n\n"
                    f"📱 <b>Service:</b> {_v2_svc_emoji(sid)} <b>{sid}</b>\n"
                    f"📞 <b>Number:</b> <code>{display_num}</code>\n"
                    f"🌍 <b>Country:</b> {flag} {c_name}\n\n"
                    f"⏳ OTP আসলে এখানে পাঠানো হবে।",
                    call.message.chat.id, call.message.message_id,
                    reply_markup=refresh_kb, parse_mode="HTML"
                )
            else:
                bot.answer_callback_query(call.id, "❌ Number পাওয়া যায়নি! পরে চেষ্টা করো।", show_alert=True)

        elif data == "v2back":
            services = _fastx_liveaccess()
            markup, has_btns = _v2_build_console_markup(services)
            if has_btns:
                bot.edit_message_text(
                    "📡 <b>V2 LIVE CONSOLE</b>\n"
                    "⚡━━━━━━━━━━━━━━⚡\n\n"
                    "🔴 <b>Live service select koro:</b>\n"
                    "<i>এই service-গুলোতে এখন OTP আসছে</i>\n\n"
                    "⚡━━━━━━━━━━━━━━⚡",
                    call.message.chat.id, call.message.message_id,
                    reply_markup=markup, parse_mode="HTML"
                )
            else:
                bot.answer_callback_query(call.id, "❌ No live services right now.", show_alert=True)
            bot.answer_callback_query(call.id)

        elif data == "grp_removeok":
            if call.from_user.id not in ADMIN_IDS:
                return
            _group_settings["otp_group_id"] = None
            _group_settings["otp_group_link"] = ""
            save_group_settings()
            bot.answer_callback_query(call.id, "✅ Group removed!")
            _show_settings_inline(call)

        elif data == "set_group_tag":
            if call.from_user.id not in ADMIN_IDS:
                return
            bot.answer_callback_query(call.id)
            cur_tag = _group_settings.get("group_tag", "BOT")
            msg = bot.send_message(
                call.message.chat.id,
                f"🌸 <b>Number Tag Set/Change</b>\n\n"
                f"🔹 <b>Bortoman Tag:</b> <code>{cur_tag}</code>\n"
                f"📱 Preview: <b>245⚜{cur_tag}⚜5660</b>\n\n"
                f"Notun tag dao (shudhu text, kono emoji na):\n"
                f"<i>Example: ATIK, BOT, OTP, KING</i>",
                reply_markup=_back_admin_kb(),
                parse_mode="HTML",
            )
            bot.register_next_step_handler(msg, _sett_get_group_tag)

    except Exception as e:
        print(f"Callback Error: {e}")


# ── Excel / CSV helpers ───────────────────────────────────────────────────────

VALID_SERVICES = [
    "facebook",
    "instagram",
    "whatsapp",
    "telegram",
    "binance",
    "pc clone",
]


def _parse_spreadsheet(data: bytes, filename: str):
    """
    Parse Excel (.xlsx / .xls) or CSV file.
    Returns:
      - (rows, mode)
        mode='two_col' → rows = list of (service, number)
        mode='one_col' → rows = list of number strings
    Accepts header rows with 'service'/'number' labels.
    Falls back: 2-column files = service+number, 1-column = numbers only.
    """
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    raw_rows = []

    if ext == "csv":
        text = data.decode("utf-8", errors="replace")
        reader = csv.reader(io.StringIO(text))
        for row in reader:
            cleaned = [c.strip() for c in row if c.strip()]
            if cleaned:
                raw_rows.append(cleaned)
    elif ext == "xlsx":
        wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        ws = wb.active
        for row in ws.iter_rows(values_only=True):
            cleaned = [str(c).strip() for c in row if c is not None and str(c).strip()]
            if cleaned:
                raw_rows.append(cleaned)
    elif ext == "xls":
        wb = xlrd.open_workbook(file_contents=data)
        ws = wb.sheet_by_index(0)
        for ri in range(ws.nrows):
            cleaned = [
                str(ws.cell_value(ri, ci)).strip()
                for ci in range(ws.ncols)
                if str(ws.cell_value(ri, ci)).strip()
            ]
            if cleaned:
                raw_rows.append(cleaned)
    else:
        return [], "unknown"

    if not raw_rows:
        return [], "empty"

    # Detect header row
    start = 0
    first = [c.lower() for c in raw_rows[0]]
    if any(h in first for h in ("service", "number", "phone", "mobile")):
        start = 1

    data_rows = raw_rows[start:]
    if not data_rows:
        return [], "empty"

    # Detect mode by column count of the majority of rows
    two_col_count = sum(1 for r in data_rows if len(r) >= 2)
    one_col_count = len(data_rows) - two_col_count

    if two_col_count > one_col_count:
        result = []
        for r in data_rows:
            if len(r) < 2:
                continue
            col0, col1 = r[0], r[1]
            # Determine which column is service and which is number
            col0_is_num = re.match(r"^\+?\d{6,15}$", re.sub(r"\s", "", col0))
            col1_is_num = re.match(r"^\+?\d{6,15}$", re.sub(r"\s", "", col1))
            if col0_is_num and not col1_is_num:
                svc = col1.lower().strip()
                num = re.sub(r"\D", "", col0)
            elif col1_is_num and not col0_is_num:
                svc = col0.lower().strip()
                num = re.sub(r"\D", "", col1)
            else:
                svc = col0.lower().strip()
                num = re.sub(r"\D", "", col1)
            if num and len(num) >= 7:
                result.append((svc, num))
        return result, "two_col"
    else:
        result = []
        for r in data_rows:
            num = re.sub(r"\D", "", r[0])
            if len(num) >= 7:
                result.append(num)
        return result, "one_col"


def _add_numbers_bulk(svc: str, numbers: list):
    """Add a list of number strings to stock[svc]. Returns (added, skipped)."""
    added, skipped = 0, 0
    svc = svc.lower().strip()
    if svc not in stock:
        return 0, len(numbers)
    for num in numbers:
        num = re.sub(r"\D", "", str(num))
        if not num:
            skipped += 1
            continue
        c_name, _ = get_country_details(num)
        if c_name == "Unknown":
            skipped += 1
            continue
        if c_name not in stock[svc]:
            stock[svc][c_name] = []
        stock[svc][c_name].append(num)
        added += 1
    if added:
        save_stock()
    return added, skipped


def _service_select_markup():
    m = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
    m.add("Facebook", "Instagram", "WhatsApp", "Telegram", "Binance", "PC Clone")
    return m


@bot.message_handler(content_types=["document"])
def document_handler(message):
    uid = message.from_user.id
    if uid not in ADMIN_IDS:
        return
    register_user(message.chat.id)

    doc = message.document
    name = doc.file_name or ""
    ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""

    if ext not in ("xlsx", "xls", "csv"):
        bot.send_message(
            message.chat.id,
            "❌ <b>Unsupported file!</b>\n\n"
            "📎 Supported formats:\n"
            "  • <b>.xlsx</b> — Excel (new)\n"
            "  • <b>.xls</b>  — Excel (old)\n"
            "  • <b>.csv</b>  — CSV\n\n"
            "💡 File pathao abar!",
            parse_mode="HTML",
        )
        return

    wait = bot.send_message(
        message.chat.id, f"⏳🔥 <b>{name}</b> parse korchi...", parse_mode="HTML"
    )

    try:
        file_info = bot.get_file(doc.file_id)
        raw = bot.download_file(file_info.file_path)
    except Exception as e:
        bot.edit_message_text(
            f"❌ File download hoyni: {e}",
            message.chat.id,
            wait.message_id,
            parse_mode="HTML",
        )
        return

    rows, mode = _parse_spreadsheet(raw, name)

    try:
        bot.delete_message(message.chat.id, wait.message_id)
    except Exception:
        pass

    if mode in ("unknown", "empty") or not rows:
        bot.send_message(
            message.chat.id,
            "⚠️ <b>File-e kono data paini!</b> ⚠️\n\n"
            "📋 <b>Supported formats:</b>\n"
            "  • <b>2-column:</b>  Service | Number\n"
            "  • <b>1-column:</b>  Number only (service pore dao)\n\n"
            "💡 Sample format:\n"
            "<code>facebook  | 8801700123456\n"
            "whatsapp  | 8801800234567\n"
            "telegram  | 251912345678</code>",
            parse_mode="HTML",
        )
        return

    if mode == "two_col":
        # Group by service and add directly
        service_map = {}
        for svc, num in rows:
            service_map.setdefault(svc, []).append(num)

        total_added, total_skipped = 0, 0
        report_lines = ""
        for svc, nums in service_map.items():
            added, skipped = _add_numbers_bulk(svc, nums)
            total_added += added
            total_skipped += skipped
            icon = "✅" if added else "⚠️"
            report_lines += f"{icon} <b>{svc.upper()}</b>: +{added} added"
            if skipped:
                report_lines += f"  (⚠️ {skipped} skip)"
            report_lines += "\n"

        bot.send_message(
            message.chat.id,
            f"📊🔥 <b>EXCEL IMPORT DONE!</b> 🔥📊\n"
            f"⚡━━━━━━━━━━━━━━━━⚡\n\n"
            f"📎 <b>File:</b> <code>{name}</code>\n"
            f"📋 <b>Rows parsed:</b> {len(rows)}\n\n"
            f"{report_lines}\n"
            f"✅ <b>Total added:</b> {total_added}\n"
            f"⚠️ <b>Skipped:</b> {total_skipped}\n\n"
            f"⚡━━━━━━━━━━━━━━━━⚡\n"
            f"💡 /panels diye stock check koro.",
            reply_markup=main_menu(uid),
            parse_mode="HTML",
        )

    else:
        # one_col: ask which service
        _pending_excel[uid] = {"numbers": rows, "filename": name}
        bot.send_message(
            message.chat.id,
            f"📂🔥 <b>FILE LOADED!</b> 🔥📂\n"
            f"⚡━━━━━━━━━━━━━━━━⚡\n\n"
            f"📎 <b>File:</b> <code>{name}</code>\n"
            f"📱 <b>Numbers found:</b> {len(rows)}\n\n"
            f" <b>Kon service-e add korbo?</b>\n"
            f"⬇️ Choose koro:",
            reply_markup=_service_select_markup(),
            parse_mode="HTML",
        )
        msg = bot.send_message(
            message.chat.id, "⬇️ Service type koro:", parse_mode="HTML"
        )
        bot.register_next_step_handler(msg, _excel_pick_service)


def _excel_pick_service(message):
    uid = message.from_user.id
    if uid not in ADMIN_IDS:
        return
    if _intercept_menu_btn(message):
        _pending_excel.pop(uid, None)
        return
    svc_raw = (message.text or "").strip().lower()
    # normalise common aliases
    svc_map = {
        "facebook": "facebook",
        "fb": "facebook",
        "instagram": "instagram",
        "ig": "instagram",
        "whatsapp": "whatsapp",
        "wa": "whatsapp",
        "telegram": "telegram",
        "tg": "telegram",
        "binance": "binance",
        "bnb": "binance",
        "pc clone": "pc clone",
        "pc": "pc clone",
        "clone": "pc clone",
    }
    svc = svc_map.get(svc_raw)
    if svc is None:
        # try direct match
        for s in VALID_SERVICES:
            if svc_raw == s:
                svc = s
                break
    if svc is None:
        msg = bot.send_message(
            message.chat.id,
            "❌ Valid service choose koro:\n"
            "<code>Facebook / Instagram / WhatsApp / Telegram / Binance / PC Clone</code>",
            reply_markup=_service_select_markup(),
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, _excel_pick_service)
        return

    pending = _pending_excel.pop(uid, None)
    if not pending:
        bot.send_message(
            message.chat.id,
            "⚠️ Session expired. File abar pathao.",
            reply_markup=main_menu(uid),
        )
        return

    numbers = pending["numbers"]
    filename = pending["filename"]
    added, skipped = _add_numbers_bulk(svc, numbers)

    bot.send_message(
        message.chat.id,
        f"📊🔥 <b>EXCEL IMPORT DONE!</b> 🔥📊\n"
        f"⚡━━━━━━━━━━━━━━━━⚡\n\n"
        f"📎 <b>File:</b>     <code>{filename}</code>\n"
        f"💬 <b>Service:</b>  <b>{svc.upper()}</b>\n"
        f"📱 <b>Parsed:</b>   {len(numbers)}\n\n"
        f"✅ <b>Added:</b>    {added}\n"
        f"⚠️ <b>Skipped:</b>  {skipped}\n\n"
        f"⚡━━━━━━━━━━━━━━━━⚡\n"
        f"💡 /panels diye stock check koro.",
        reply_markup=main_menu(uid),
        parse_mode="HTML",
    )


@bot.message_handler(func=lambda m: True)
def text_handler(message):
    global stock
    uid = message.from_user.id
    txt = message.text
    register_user(message.chat.id)

    if txt in ("☎️ 𝗩𝟭 𝗡𝗨𝗠𝗕𝗔𝗥 ☎️", "☎️ 𝗡𝗨𝗠𝗕𝗔𝗥 ☎️"):
        show_services(message)

    elif txt == "🔄 𝗩𝟮 𝗦𝗪𝗜𝗧𝗖𝗛":
        _v2_users.add(uid)
        _save_v2_users()
        bot.send_message(
            message.chat.id,
            "🔄 <b>V2 SWITCH</b>\n"
            "⚡━━━━━━━━━━━━━━⚡\n\n"
            "কোন mode চাও select koro:\n\n"
            "🔴 <b>LIVE RANGE</b> — Panel থেকে live OTP range দেখাবে\n"
            "⌨️ <b>CUSTOM RANGE</b> — নিজে range দাও, matching number পাবে\n\n"
            "⚡━━━━━━━━━━━━━━⚡",
            reply_markup=v2_switch_menu(),
            parse_mode="HTML",
        )

    elif txt == "🔴 𝗟𝗜𝗩𝗘 𝗥𝗔𝗡𝗚𝗘":
        _v2_show_console(message.chat.id)

    elif txt == "⌨️ 𝗖𝗨𝗦𝗧𝗢𝗠 𝗥𝗔𝗡𝗚𝗘":
        cancel_markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        cancel_markup.add(types.KeyboardButton("❌ Cancel"))
        msg = bot.send_message(
            message.chat.id,
            "⌨️ <b>CUSTOM RANGE</b>\n"
            "⚡━━━━━━━━━━━━━━⚡\n\n"
            "📲 কোন range/prefix চাও লিখো:\n"
            "<i>উদাহরণ: <code>8801</code>, <code>44</code>, <code>33</code></i>\n\n"
            "⚡━━━━━━━━━━━━━━⚡",
            reply_markup=cancel_markup,
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, _v2_custom_range_step)

    elif txt == "🔙 𝗩𝟭 𝗦𝗪𝗜𝗧𝗖𝗛":
        _v2_users.discard(uid)
        _save_v2_users()
        mname = message.from_user.first_name or message.from_user.username or "User"
        bot.send_message(
            message.chat.id,
            f"╔═════════════════════╗\n"
            f"      USER MENU-te WELCOME!\n"
            f"   👋 <b>{mname}</b>, ki korte chao?\n"
            f"╚═════════════════════╝",
            reply_markup=main_menu(uid),
            parse_mode="HTML",
        )

    elif txt in _get_svc_map():
        svc = _get_svc_map()[txt]
        show_countries(message.chat.id, svc)

    elif txt in ("🔙 Admin Menu", "🔙 Admin Panel", "🔙 𝗔𝗗𝗠𝗜𝗡 𝗣𝗔𝗡𝗘𝗟") and uid in ADMIN_IDS:
        _go_admin_panel(message)

    elif txt == "🔙 Main Menu":
        mname = message.from_user.first_name or message.from_user.username or "User"
        bot.send_message(
            message.chat.id,
            f"╔═════════════════════╗\n"
            f"      USER MENU-te WELCOME!\n"
            f"   👋 <b>{mname}</b>, ki korte chao?\n"
            f"╚═════════════════════╝",
            reply_markup=main_menu(uid),
            parse_mode="HTML",
        )

    elif txt == "📞 𝗦𝗔𝗣𝗢𝗥𝗧":
        markup = types.InlineKeyboardMarkup()
        markup.add(
            types.InlineKeyboardButton("📩 Support Team", url="https://t.me/Tom_9805")
        )
        bot.send_message(
            message.chat.id,
            "📞 <b>SUPPORT</b> 📞\n"
            "⚡━━━━━━━━━━━━━━⚡\n\n"
            "Kono somossa hole nicher button e click koro!\n\n"
            "⚡━━━━━━━━━━━━━━⚡",
            reply_markup=markup,
            parse_mode="HTML",
        )

    elif txt == "📊 𝗦𝗧𝗢𝗖𝗞":
        report = "🔥 <b>LIVE STOCK REPORT</b> 🔥\n⚡━━━━━━━━━━━━⚡\n\n"
        for s, d in stock.items():
            total = sum(len(v) for v in d.values())
            report += f" <b>{s.upper()}</b>: {total} টি \n"
        report += "\n⚡━━━━━━━━━━━━⚡\n🤖 <b>AR OTP BOT</b> 🔥"
        bot.send_message(message.chat.id, report, parse_mode="HTML")

    elif txt == "⚙️ 𝗔𝗗𝗠𝗜𝗡 𝗣𝗔𝗡𝗘𝗟 ⚙️" and uid in ADMIN_IDS:
        _go_admin_panel(message)

    elif txt == "🔥📢 𝗕𝗿𝗼𝗮𝗱𝗰𝗮𝘀𝘁" and uid in ADMIN_IDS:
        msg = bot.send_message(
            message.chat.id,
            "✍️ <b>Broadcast content পাঠাও:</b> \n\n"
            "📝 Text, 🖼️ Photo, 🎥 Video, or 🎭 Sticker (with optional caption) — সব accept হবে!\n\n"
            "🔙 Back jete <b>Admin Panel</b> button press koro.",
            reply_markup=_back_admin_kb(),
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, do_broadcast)

    elif txt == "⚡👥 𝗨𝘀𝗲𝗿 𝗖𝗼𝘂𝗻𝘁" and uid in ADMIN_IDS:
        bot.send_message(
            message.chat.id,
            f" <b>TOTAL USERS</b> \n\n⚡ <b>{len(users)}</b> জন আছে! 🔥",
            parse_mode="HTML",
        )

    elif txt == "📋👥 𝗨𝘀𝗲𝗿 𝗟𝗶𝘀𝘁" and uid in ADMIN_IDS:
        all_ids = list(users)
        total = len(all_ids)
        if total == 0:
            bot.send_message(message.chat.id, "📋 No users yet.", parse_mode="HTML")
        else:
            bot.send_message(
                message.chat.id, "⏳ Loading user names...", parse_mode="HTML"
            )
            updated = False
            for user_id in all_ids:
                key = str(user_id)
                existing = user_names.get(key, "")
                if existing and not existing.strip().lstrip("-").isdigit():
                    continue
                try:
                    chat_info = bot.get_chat(user_id)
                    full = f"{chat_info.first_name or ''} {chat_info.last_name or ''}".strip()
                    uname = chat_info.username or ""
                    if full and uname:
                        display = f"{full} (@{uname})"
                    elif full:
                        display = full
                    elif uname:
                        display = f"@{uname}"
                    else:
                        display = None
                    if display:
                        user_names[key] = display
                        updated = True
                except Exception:
                    pass
            if updated:
                save_json(USER_NAMES_FILE, user_names)

            PAGE = 50
            chunks = [all_ids[i : i + PAGE] for i in range(0, total, PAGE)]
            for idx, chunk in enumerate(chunks):
                lines = (
                    f"📋👥 <b>USER LIST</b> 👥📋\n"
                    f"⚡━━━━━━━━━━━━━━━━⚡\n"
                    f"📊 Total: <b>{total}</b> users"
                    + (f"  |  Page {idx + 1}/{len(chunks)}" if len(chunks) > 1 else "")
                    + "\n⚡━━━━━━━━━━━━━━━━⚡\n\n"
                )
                for i, user_id in enumerate(chunk, start=idx * PAGE + 1):
                    name = user_names.get(str(user_id), "—")
                    lines += f"{i}. 🆔 <code>{user_id}</code>\n    👤 {name}\n\n"
                bot.send_message(message.chat.id, lines, parse_mode="HTML")

    elif txt == "📈 𝗢𝗧𝗣 𝗦𝘁𝗮𝘁𝘀" and uid in ADMIN_IDS:
        with otp_stats_lock:
            stats_copy = dict(otp_stats)
        if not stats_copy:
            bot.send_message(
                message.chat.id,
                "📈 <b>OTP STATS</b>\n\n"
                "⚠️ এখনো কোনো OTP ডেলিভার হয়নি।",
                parse_mode="HTML",
            )
        else:
            sorted_stats = sorted(stats_copy.items(), key=lambda x: x[1], reverse=True)
            total_otps = sum(stats_copy.values())
            PAGE = 30
            chunks = [sorted_stats[i:i+PAGE] for i in range(0, len(sorted_stats), PAGE)]
            for idx, chunk in enumerate(chunks):
                lines = (
                    f"📈 <b>OTP STATS</b>\n"
                    f"⚡━━━━━━━━━━━━━━━━⚡\n"
                    f"📊 মোট OTP ডেলিভার: <b>{total_otps}</b> টি"
                    + (f"  |  Page {idx+1}/{len(chunks)}" if len(chunks) > 1 else "")
                    + f"\n👥 মোট ইউজার: <b>{len(sorted_stats)}</b> জন\n"
                    f"⚡━━━━━━━━━━━━━━━━⚡\n\n"
                )
                for rank, (user_id, count) in enumerate(chunk, start=idx*PAGE+1):
                    name = user_names.get(str(user_id), "")
                    if not name or str(name).strip().lstrip("-").isdigit():
                        try:
                            chat_info = bot.get_chat(int(user_id))
                            full = f"{chat_info.first_name or ''} {chat_info.last_name or ''}".strip()
                            uname = chat_info.username or ""
                            name = f"{full} (@{uname})" if full and uname else (full or f"@{uname}" if uname else str(user_id))
                            user_names[str(user_id)] = name
                        except Exception:
                            name = str(user_id)
                    medal = "🥇" if rank == 1 else "🥈" if rank == 2 else "🥉" if rank == 3 else f"{rank}."
                    lines += f"{medal} <code>{user_id}</code> — <b>{count}</b> টি OTP\n    👤 {name}\n\n"
                bot.send_message(message.chat.id, lines, parse_mode="HTML")

    elif txt == "➕ 𝗡𝘂𝗺𝗯𝗮𝗿 𝗔𝗱𝗱" and uid in ADMIN_IDS:
        m = types.ReplyKeyboardMarkup(resize_keyboard=True)
        m.add("facebook", "instagram", "whatsapp", "telegram", "binance", "pc clone")
        m.add("❌ Cancel")
        msg = bot.send_message(
            message.chat.id,
            "🔥 <b>Service choose koro:</b> 🔥",
            reply_markup=m,
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, process_auto_add)

    elif txt == "🗑️ 𝗦𝗼𝗯 𝗖𝗹𝗲𝗮𝗿" and uid in ADMIN_IDS:
        bot.send_message(
            message.chat.id,
            "🗑️🔥 <b>STOCK CLEAR PANEL</b> 🔥🗑️\n\n"
            " <b>Kon service-er stock clear korbe?</b>\n"
            "⬇️ Service choose koro:",
            reply_markup=_clr_service_markup(),
            parse_mode="HTML",
        )

    elif txt == "🎭 𝗗𝗘𝗠𝗢 𝗢𝗧𝗣" and uid in ADMIN_IDS:
        bot.send_message(
            message.chat.id,
            demo_status_text(),
            reply_markup=demo_menu_markup(),
            parse_mode="HTML",
        )
        with _demo_lock:
            has_configs = len(_demo_configs) > 0
        if has_configs:
            bot.send_message(
                message.chat.id,
                "⚡ <b>Config Start/Stop:</b>",
                reply_markup=demo_cfg_inline_markup(),
                parse_mode="HTML",
            )

    elif txt == "➕ 𝗔𝗱𝗱 𝗣𝗮𝗻𝗲𝗹" and uid in ADMIN_IDS:
        _show_addpanel_type_select(message.chat.id, uid)

    elif txt == "➕ 𝗔𝗱𝗱 𝗦𝗲𝗿𝘃𝗶𝗰𝗲" and uid in ADMIN_IDS:
        _addservice_state[uid] = {}
        msg = bot.send_message(
            message.chat.id,
            "📋🔥 <b>ADD NEW SERVICE</b> 🔥📋\n\n"
            "🏷️ <b>Step 1/2:</b> Button-e ki lekha thakbe?\n"
            "<i>Example: Telegram 🔵, Binance 💛, TikTok 🎵</i>",
            reply_markup=_back_admin_kb(),
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, _svc_get_label)

    elif txt == "🗑️ 𝗥𝗲𝗺𝗼𝘃𝗲 𝗦𝗲𝗿𝘃𝗶𝗰𝗲" and uid in ADMIN_IDS:
        if not _services:
            bot.send_message(message.chat.id, "📋 Kono service nai!", parse_mode="HTML")
        else:
            markup = types.InlineKeyboardMarkup(row_width=1)
            for s in _services:
                markup.add(types.InlineKeyboardButton(
                    f"🗑️ {s['label']}  [{s['key']}]",
                    callback_data=f"rmsvc:{s['key']}",
                ))
            bot.send_message(
                message.chat.id,
                "🗑️🔥 <b>REMOVE SERVICE</b>\n\nKon service remove korbe?",
                reply_markup=markup,
                parse_mode="HTML",
            )

    elif txt == "🗑️ 𝗥𝗲𝗺𝗼𝘃𝗲 𝗣𝗮𝗻𝗲𝗹" and uid in ADMIN_IDS:
        if not _dynamic_panels:
            bot.send_message(
                message.chat.id,
                "📋 <b>Kono dynamic panel nai!</b>\n💡 Add Panel button diye add koro.",
                parse_mode="HTML",
            )
        else:
            markup = types.InlineKeyboardMarkup(row_width=1)
            for p in _dynamic_panels:
                pid = p["id"]
                with _stats_lock:
                    s = _panel_stats.get(pid, {})
                st = s.get("status", "⏳")
                markup.add(
                    types.InlineKeyboardButton(
                        f"{st} {p.get('username','?')} — {p.get('host','?')}",
                        callback_data=f"rmpanel:{pid}",
                    )
                )
            bot.send_message(
                message.chat.id,
                "🗑️🔥 <b>REMOVE PANEL</b>\n\nKon panel remove korbe?",
                reply_markup=markup,
                parse_mode="HTML",
            )


    elif txt == "➕ 𝗖𝗼𝗻𝗳𝗶𝗴 𝗬𝗼𝗴 𝗞𝗼𝗿𝗼" and uid in ADMIN_IDS:
        _demo_cfg_temp[uid] = {}
        msg = bot.send_message(
            message.chat.id,
            "📱 <b>Phone number(s) dao:</b>\n\n"
            "• Ekta number: <code>8801700123456</code>\n"
            "• Multiple (newline or comma):\n"
            "<code>8801700123456\n251912345678\n2348012345678</code>\n\n"
            "⚠️ Full country code including number lagbe!",
            reply_markup=_back_admin_kb(),
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, _demo_cfg_number)

    elif txt == "🗑️ 𝗖𝗼𝗻𝗳𝗶𝗴 𝗠𝘂𝗰𝗵𝗼" and uid in ADMIN_IDS:
        with _demo_lock:
            configs = list(_demo_configs)
        if not configs:
            bot.send_message(
                message.chat.id,
                "📋 <b>Kono config nai!</b>",
                reply_markup=demo_menu_markup(),
                parse_mode="HTML",
            )
        else:
            markup = types.InlineKeyboardMarkup(row_width=1)
            for cfg in configs:
                svcs = ", ".join(cfg.get("services") or ["?"])
                markup.add(types.InlineKeyboardButton(
                    f"🗑️ {cfg['name']}  [{svcs}  |  {cfg['interval']}s]",
                    callback_data=f"rmcfg:{cfg['id']}",
                ))
            bot.send_message(
                message.chat.id,
                "🗑️🔥 <b>Config মুছো</b>\n\nকোন config মুছবে?",
                reply_markup=markup,
                parse_mode="HTML",
            )

    elif txt == "📊 𝗣𝗮𝗻𝗲𝗹𝘀" and uid in ADMIN_IDS:
        panels_cmd(message)

    elif txt == "📤 𝗣𝘂𝗿𝗮𝗻𝗼 𝗢𝗧𝗣 𝗚𝗿𝘂𝗽𝗲 𝗦𝗲𝗻𝗱" and uid in ADMIN_IDS:
        _resend_old_otps(message)

    elif txt == "🛑 𝗣𝘂𝗿𝗮𝗻𝗼 𝗢𝗧𝗣 𝗕𝗼𝗻𝗱𝗵𝗼" and uid in ADMIN_IDS:
        global _resend_stop, _resend_running
        _resend_stop = True
        if _resend_running:
            bot.send_message(message.chat.id,
                "🛑 <b>Resend stop signal pathano hoise!</b>\n"
                "<i>Chal OTP send hoye gele bondho hobe.</i>",
                parse_mode="HTML")
        else:
            bot.send_message(message.chat.id,
                "ℹ️ <b>Kono resend cholthechhilo na.</b>",
                parse_mode="HTML")

    elif txt == "🔍 𝗧𝗲𝘀𝘁 𝗣𝗮𝗻𝗲𝗹" and uid in ADMIN_IDS:
        _testpanel_state[uid] = {"step": "url", "data": {}}
        msg = bot.send_message(
            message.chat.id,
            "🔍🔥 <b>TEST PANEL</b> 🔥🔍\n\n"
            "Panel-er jekono URL pathao — test korbo, <b>save korbo na</b>.\n\n"
            "✅ <b>Jekono format:</b>\n"
            "• <code>http://1.2.3.4/konekta/agent/SMSCDRReports</code>\n"
            "• <code>http://1.2.3.4/ints</code>\n"
            "• <code>https://truesms.net</code>",
            reply_markup=_back_admin_kb(),
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, _tp_get_url)

    elif txt == "👑 𝗔𝗱𝗱 𝗔𝗱𝗺𝗶𝗻" and uid in ADMIN_IDS:
        if not is_super_admin(uid):
            bot.send_message(message.chat.id, "❌ <b>Shudhu Super Admin notun admin add korte parbe!</b>", parse_mode="HTML")
            return
        msg = bot.send_message(
            message.chat.id,
            "👑 <b>New Admin add</b>\n\n"
            "Notun admin-er Telegram <b>User ID</b> dao:\n"
            "<i>Example: 123456789</i>\n\n"
            "💡 User ID jante hole sei user-ke @userinfobot-e forward koro.",
            reply_markup=_back_admin_kb(),
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, _admin_add_get_id)

    elif txt == "🗑️ 𝗥𝗲𝗺𝗼𝘃𝗲 𝗔𝗱𝗺𝗶𝗻" and uid in ADMIN_IDS:
        if not is_super_admin(uid):
            bot.send_message(message.chat.id, "❌ <b>Shudhu Super Admin admin remove korte parbe!</b>", parse_mode="HTML")
            return
        _show_remove_admin(message)

    elif txt == "📞 𝗦𝘂𝗽𝗽𝗼𝗿𝘁 𝗜𝗗" and uid in ADMIN_IDS:
        if not is_super_admin(uid):
            bot.send_message(message.chat.id, "❌ <b>Shudhu Super Admin Support ID set korte parbe!</b>", parse_mode="HTML")
            return
        cur = _group_settings.get("support_id", "") or "❌ Set hoy nai"
        msg = bot.send_message(
            message.chat.id,
            f"📞 <b>Support ID Set/Change</b>\n\n"
            f"🔹 <b>Bortoman Support ID:</b> {cur}\n\n"
            f"Notun Support Telegram ID dao\n"
            f"<i>(User ID, username ba t.me link — jokono ekta)</i>\n\n"
            f"Example: <code>@support_user</code> ba <code>123456789</code>",
            reply_markup=_back_admin_kb(),
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, _sett_get_support_id)

    elif txt == "⚙️ 𝗦𝗲𝘁𝘁𝗶𝗻𝗴𝘀" and uid in ADMIN_IDS:
        _show_settings(message)

    elif txt == "✏️ 𝗘𝗱𝗶𝘁 𝗠𝗲𝘀𝘀𝗮𝗴𝗲𝘀" and uid in ADMIN_IDS:
        _show_edit_messages_menu(message)

    elif txt == "📡 𝗩𝟮 𝗠𝗲𝘀𝘀𝗮𝗴𝗲 𝗙𝗼𝗿𝗺𝗮𝘁" and uid in ADMIN_IDS:
        import html as _html
        current = get_template("otp_dm_v2")
        current_esc = _html.escape(current[:600])
        vars_hint = _TEMPLATE_VARS.get("otp_dm_v2", "")
        markup = types.InlineKeyboardMarkup(row_width=1)
        markup.add(types.InlineKeyboardButton("✏️ Edit V2 DM Format", callback_data="editmsg:otp_dm_v2"))
        markup.add(types.InlineKeyboardButton("🔄 Default এ Reset", callback_data="editmsg_reset:otp_dm_v2"))
        bot.send_message(
            message.chat.id,
            "📡 <b>V2 Message Format</b>\n"
            "⚡━━━━━━━━━━━━━━⚡\n\n"
            "📌 <b>ব্যবহারযোগ্য ভেরিয়েবল:</b>\n"
            f"<code>{vars_hint}</code>\n\n"
            "📄 <b>বর্তমান V2 DM Format:</b>\n"
            f"<code>{current_esc}</code>\n\n"
            "⚡━━━━━━━━━━━━━━⚡\n"
            "<i>ℹ️ V2 mode-এ Get New Number ও Change Country বাটন দেখায় না।</i>",
            reply_markup=markup,
            parse_mode="HTML",
        )

    elif txt in ("🔙 𝗔𝗗𝗠𝗜𝗡 𝗣𝗔𝗡𝗘𝗟", "🔙 Admin Panel") and uid in ADMIN_IDS:
        _go_admin_panel(message)

    elif txt == "⬅️🔙 𝗨𝘀𝗲𝗿 𝗠𝗲𝗻𝘂":
        mname = message.from_user.first_name or message.from_user.username or "User"
        bot.send_message(
            message.chat.id,
            f"╔═════════════════════╗\n"
            f"      USER MENU-te WELCOME!\n"
            f"   👋 <b>{mname}</b>, ki korte chao?\n"
            f"╚═════════════════════╝",
            reply_markup=main_menu(uid),
            parse_mode="HTML",
        )


# ── Demo OTP config step handlers ─────────────────────────────────────────────


def _demo_cfg_number(message):
    if _is_back(message.text):
        _go_admin_panel(message)
        return
    if _intercept_menu_btn(message):
        return
    raw_lines = re.split(r"[\n,]+", message.text or "")
    candidates = [re.sub(r"\D", "", ln) for ln in raw_lines if re.sub(r"\D", "", ln)]
    if not candidates:
        msg = bot.send_message(
            message.chat.id,
            "❌ Kono number paini. Ekta ba multiple number dao:",
            reply_markup=_back_admin_kb(),
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, _demo_cfg_number)
        return
    valid, invalid = [], []
    result_lines = ""
    for num in candidates:
        if len(num) < 7:
            invalid.append(num)
            continue
        c_name, flag = get_country_details(num)
        if c_name == "Unknown":
            invalid.append(num)
        else:
            valid.append(num)
            result_lines += f"  ✅ <code>{num}</code>  {flag} {c_name}\n"
    if not valid:
        msg = bot.send_message(
            message.chat.id,
            f"⚠️ <b>Kono valid number paini!</b>\n\n"
            f"Full international number dao (country code including):\n"
            f"🇧🇩 Bangladesh → <code>8801700123456</code>\n"
            f"🇪🇹 Ethiopia   → <code>251912345678</code>\n"
            f"🇳🇬 Nigeria    → <code>2348012345678</code>\n\n"
            f"Aro ekbar try koro:",
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, _demo_cfg_number)
        return
    uid = message.from_user.id
    _demo_cfg_temp.setdefault(uid, {})["numbers"] = valid
    SHOW_MAX = 10
    shown = result_lines.split("\n")[:SHOW_MAX]
    preview = "\n".join(shown)
    if len(valid) > SHOW_MAX:
        preview += f"\n  ... +{len(valid) - SHOW_MAX} more"
    feedback = f"✅ <b>{len(valid)} টি number set hoiche:</b>\n{preview}\n"
    if invalid:
        inv_preview = invalid[:5]
        feedback += (
            f"\n⚠️ Skip (invalid): {', '.join(f'<code>{x}</code>' for x in inv_preview)}"
        )
        if len(invalid) > 5:
            feedback += f" +{len(invalid) - 5} more"
        feedback += "\n"
    svc_markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    svc_markup.add("4", "5", "6", "7", "8")
    svc_markup.add("🔙 Admin Panel")
    msg = bot.send_message(
        message.chat.id,
        feedback + "\n🔢 <b>OTP digit count choose koro:</b>",
        reply_markup=svc_markup,
        parse_mode="HTML",
    )
    bot.register_next_step_handler(msg, _demo_cfg_digits)


def _demo_cfg_digits(message):
    if _is_back(message.text):
        _go_admin_panel(message)
        return
    if _intercept_menu_btn(message):
        return
    try:
        d = int(message.text.strip())
        if d < 4 or d > 8:
            raise ValueError
    except ValueError:
        svc_markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        svc_markup.add("4", "5", "6", "7", "8")
        svc_markup.add("🔙 Admin Panel")
        msg = bot.send_message(message.chat.id, "❌ 4 theke 8 er modhye number dao:", reply_markup=svc_markup)
        bot.register_next_step_handler(msg, _demo_cfg_digits)
        return
    uid = message.from_user.id
    _demo_cfg_temp.setdefault(uid, {})["digits"] = d
    _demo_svc_state[uid] = []
    _demo_cfg_service_ask(message)


def _demo_cfg_service_ask(message):
    uid = message.from_user.id
    current = _demo_svc_state.get(uid, [])
    svc_markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=3)
    svc_markup.add("Facebook", "Instagram", "WhatsApp")
    svc_markup.add("Telegram", "PC Clone", "Twitter")
    svc_markup.add("Tiktok", "Snapchat", "Gmail")
    if current:
        svc_markup.add("✅ হয়েছে (Done)")
    svc_markup.add("🔙 Admin Panel")
    if current:
        svc_list = "\n".join(f"  ✅ {s}" for s in current)
        prompt = (
            f"✅ <b>Selected services ({len(current)}):</b>\n{svc_list}\n\n"
            f"➕ <b>আরো service যোগ করো</b> অথবা <b>✅ হয়েছে</b> চাপো:"
        )
    else:
        prompt = (
            "💬 <b>Service choose koro</b>\n\n"
            "<i>একাধিক service add করা যাবে — সব শেষে '✅ হয়েছে' চাপো।</i>"
        )
    msg = bot.send_message(message.chat.id, prompt, reply_markup=svc_markup, parse_mode="HTML")
    bot.register_next_step_handler(msg, _demo_cfg_service_multi)


def _demo_cfg_service_multi(message):
    if _is_back(message.text):
        _go_admin_panel(message)
        return
    if _intercept_menu_btn(message):
        return
    uid = message.from_user.id
    txt = (message.text or "").strip()

    if txt in ("✅ হয়েছে (Done)", "✅ হয়েছে"):
        svcs = _demo_svc_state.get(uid, [])
        if not svcs:
            bot.send_message(
                message.chat.id,
                "⚠️ <b>কমপক্ষে একটা service select করো!</b>",
                parse_mode="HTML",
            )
            _demo_cfg_service_ask(message)
            return
        uid2 = message.from_user.id
        _demo_cfg_temp.setdefault(uid2, {})["services"] = svcs
        intvl_markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=4)
        intvl_markup.add("5", "10", "15", "30", "60", "120", "300")
        intvl_markup.add("🔙 Admin Panel")
        svc_list = ", ".join(svcs)
        msg = bot.send_message(
            message.chat.id,
            f"✅ <b>Services set:</b> {svc_list}\n\n⏱️ <b>Interval (seconds) dao:</b>",
            reply_markup=intvl_markup,
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, _demo_cfg_interval)
        return

    if not txt:
        _demo_cfg_service_ask(message)
        return

    current = _demo_svc_state.setdefault(uid, [])
    if txt in current:
        bot.send_message(
            message.chat.id,
            f"⚠️ <b>{txt}</b> already added আছে! আরো যোগ করো বা <b>✅ হয়েছে</b> চাপো।",
            parse_mode="HTML",
        )
    else:
        current.append(txt)
    _demo_cfg_service_ask(message)


def _demo_cfg_interval(message):
    if _is_back(message.text):
        _go_admin_panel(message)
        return
    if _intercept_menu_btn(message):
        return
    try:
        iv = int(message.text.strip())
        if iv < 5:
            raise ValueError
    except ValueError:
        intvl_markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=4)
        intvl_markup.add("5", "10", "15", "30", "60", "120", "300")
        intvl_markup.add("🔙 Admin Panel")
        msg = bot.send_message(message.chat.id, "❌ Minimum 5 second. Aro dao:", reply_markup=intvl_markup)
        bot.register_next_step_handler(msg, _demo_cfg_interval)
        return
    global _demo_cfg_id_counter
    uid = message.from_user.id
    tmp = _demo_cfg_temp.pop(uid, {})
    numbers = tmp.get("numbers", ["8801700000000"])
    digits = tmp.get("digits", 6)
    services = tmp.get("services", ["Facebook"])
    with _demo_lock:
        _demo_cfg_id_counter += 1
        cid = _demo_cfg_id_counter
        cfg_name = f"Config {cid}"
        _demo_configs.clear()
        _demo_next_fire.clear()
        _demo_configs.append({
            "id": cid,
            "name": cfg_name,
            "active": True,
            "numbers": numbers,
            "digits": digits,
            "services": services,
            "interval": iv,
        })
    svcs_str = ", ".join(services)
    bot.send_message(
        message.chat.id,
        f"✅🔥 <b>{cfg_name} যোগ হয়েছে!</b>\n\n"
        f"  📱 Numbers: {len(numbers)} টি\n"
        f"  🔢 Digits: {digits}\n"
        f"  💬 Services: {svcs_str}\n"
        f"  ⏱️ Interval: {iv}s\n\n"
        + demo_status_text(),
        reply_markup=demo_menu_markup(),
        parse_mode="HTML",
    )


def make_broadcast_msg(text):
    return get_template("broadcast").format(text=text)


def do_broadcast(message):
    if _is_back(message.text):
        _go_admin_panel(message)
        return
    if _intercept_menu_btn(message):
        return
    has_text = bool(message.text)
    has_photo = bool(message.photo)
    has_video = bool(message.video)
    has_sticker = bool(message.sticker)
    has_animation = bool(message.animation)
    has_audio = bool(message.audio)
    has_voice = bool(message.voice)
    has_document = bool(message.document)
    has_video_note = bool(message.video_note)

    if not any(
        [
            has_text,
            has_photo,
            has_video,
            has_sticker,
            has_animation,
            has_audio,
            has_voice,
            has_document,
            has_video_note,
        ]
    ):
        bot.send_message(
            message.chat.id,
            "⚠️ <b>Kono content paoa jaini!</b> ⚠️\n"
            "Text, Photo, Video, GIF, Audio, Voice, Document ba Sticker pathao.",
            parse_mode="HTML",
        )
        return

    cap = (
        lambda m: make_broadcast_msg(m.caption) if m.caption else make_broadcast_msg("")
    )

    bot.send_message(
        message.chat.id,
        f"⏳🔥 <b>{len(users)} জনকে পাঠানো হচ্ছে...</b> 🔥⏳",
        parse_mode="HTML",
    )

    success, fail = 0, 0
    for uid in list(users):
        try:
            if has_photo:
                bot.send_photo(
                    uid,
                    message.photo[-1].file_id,
                    caption=cap(message),
                    parse_mode="HTML",
                )
            elif has_animation:
                bot.send_animation(
                    uid,
                    message.animation.file_id,
                    caption=cap(message),
                    parse_mode="HTML",
                )
            elif has_video:
                bot.send_video(
                    uid, message.video.file_id, caption=cap(message), parse_mode="HTML"
                )
            elif has_video_note:
                bot.send_video_note(uid, message.video_note.file_id)
            elif has_sticker:
                bot.send_sticker(uid, message.sticker.file_id)
            elif has_audio:
                bot.send_audio(
                    uid, message.audio.file_id, caption=cap(message), parse_mode="HTML"
                )
            elif has_voice:
                bot.send_voice(
                    uid, message.voice.file_id, caption=cap(message), parse_mode="HTML"
                )
            elif has_document:
                bot.send_document(
                    uid,
                    message.document.file_id,
                    caption=cap(message),
                    parse_mode="HTML",
                )
            else:
                bot.send_message(
                    uid, make_broadcast_msg(message.text), parse_mode="HTML"
                )
            success += 1
        except Exception:
            fail += 1

    bot.send_message(
        message.chat.id,
        f" <b>BROADCAST COMPLETE!</b> \n\n"
        f"✅ <b>𝗦𝗼𝗳𝗼𝗹:</b> {success} জন 🔥\n"
        f"❌ <b>𝗕𝗮𝗿𝘁𝗵𝗼:</b> {fail} জন ",
        parse_mode="HTML",
    )
    _go_admin_panel(message)


_pending_add = {}


def _start_countdown(chat_id, msg_id, svc, flag, c_name, display_num, scnt):
    if chat_id in _countdowns:
        _countdowns[chat_id].set()
    cancel = threading.Event()
    _countdowns[chat_id] = cancel

    def _make_kb():
        kb = types.InlineKeyboardMarkup(row_width=2)
        kb.add(
            types.InlineKeyboardButton("🔄 New Number", callback_data=f"n:{svc}:{scnt}"),
            types.InlineKeyboardButton("🌍 Change Country", callback_data=f"s:{svc}"),
        )
        if get_otp_group_link():
            kb.add(types.InlineKeyboardButton("📢 OTP Group", url=get_otp_group_link()))
        return kb

    def run():
        TICK = 5            # update every 5s
        DURATION = 600      # 10 minutes
        deadline = time.time() + DURATION
        current_msg_id = [msg_id]  # list so inner scope can mutate

        def _parse_retry_after(err_str):
            try:
                return int(re.search(r"retry after (\d+)", err_str).group(1))
            except Exception:
                return 60

        def try_update(text):
            """Try edit, fall back to send+delete.
            Returns: True=ok, False=skip tick, None=stop, int=rate-limited (seconds to wait)."""
            # 1. try edit
            try:
                bot.edit_message_text(
                    text, chat_id, current_msg_id[0],
                    reply_markup=_make_kb(), parse_mode="HTML",
                )
                return True
            except Exception as e:
                err = str(e)
                if "message is not modified" in err:
                    return True
                if "message to edit not found" in err or "MESSAGE_ID_INVALID" in err:
                    return None
                if "429" in err or "Too Many Requests" in err:
                    return _parse_retry_after(err)  # int → caller will wait

            # 2. edit failed (non-429) — try send+delete
            try:
                sent = bot.send_message(
                    chat_id, text,
                    reply_markup=_make_kb(), parse_mode="HTML",
                )
                try:
                    bot.delete_message(chat_id, current_msg_id[0])
                except Exception:
                    pass
                current_msg_id[0] = sent.message_id
                return True
            except Exception as e2:
                err2 = str(e2)
                if "429" in err2 or "Too Many Requests" in err2:
                    return _parse_retry_after(err2)
                print(f"[COUNTDOWN] tick failed: {e2}")
                return False

        while not cancel.is_set():
            remaining = int(deadline - time.time())
            if remaining <= 0:
                deadline = time.time() + DURATION
                remaining = DURATION

            mins = remaining // 60
            secs = remaining % 60
            text = (
                f"✅ <b>Number Assigned Successfully !</b>\n\n"
                f"🔧 <b>Platform :</b> {svc.capitalize()}\n"
                f"🌍 <b>Country :</b> {flag} {c_name}\n\n"
                f"📞 <b>Number :</b> <code>{display_num}</code>\n\n"
                f"⏱ <b>Auto code fetch :</b> {mins:02d}:{secs:02d}s"
            )
            result = try_update(text)
            if result is None:
                break  # message gone, stop
            elif type(result) is int:
                # rate-limited — wait the full retry_after, then resume
                wait = min(result, 3600)
                print(f"[COUNTDOWN] Rate limited for {wait}s, pausing timer for {chat_id}")
                cancel.wait(wait)
            else:
                cancel.wait(TICK)

    threading.Thread(target=run, daemon=True).start()


def _settings_text(uid=None):
    """Per-admin settings. If uid given, show that admin's own settings."""
    grp_link = get_admin_setting(uid, "otp_group_link") if uid else _group_settings.get("otp_group_link", "")
    grp_id = get_admin_setting(uid, "otp_group_id") if uid else _group_settings.get("otp_group_id")
    ch2 = get_admin_setting(uid, "channel2") if uid else _group_settings.get("channel2", "")
    bot_lnk = get_admin_setting(uid, "bot_link") if uid else _group_settings.get("bot_link", "")
    auto_del = _group_settings.get("auto_delete", True)
    del_secs = _group_settings.get("auto_delete_seconds", 3600)
    grp_send = _group_settings.get("group_otp_send", True)
    grp_tag = _group_settings.get("group_tag", "BOT")
    id_str = f"<code>{grp_id}</code>" if grp_id else "❌ Set hoy nai"
    link_str = grp_link or "❌ Set hoy nai"
    auto_str = f"🟢 ON ({del_secs // 60} min)" if auto_del else "🔴 OFF"
    grp_send_str = "🟢 ON (Group-e OTP jabey)" if grp_send else "🔴 OFF (Shudhu Inbox-e jabey)"
    ch2_str = ch2 or "❌ Set hoy nai"
    bot_str = bot_lnk or "❌ Set hoy nai"
    return (
        "⚙️ <b>BOT SETTINGS</b> ⚙️\n"
        "⚡━━━━━━━━━━━━━━━━⚡\n\n"
        "📡 <b>OTP GROUP</b>\n"
        f"🔗 Link: {link_str}\n"
        f"🆔 Chat ID: {id_str}\n"
        f"⏱️ Auto Delete: {auto_str}\n"
        f"📤 Group OTP Send: {grp_send_str}\n"
        f"⚜ Number Tag: <b>{grp_tag}</b> (245⚜{grp_tag}⚜5660)\n\n"
        "📢 <b>LINKS</b>\n"
        f"📢 Join Channel: {ch2_str}\n"
        f"🤖 Bot Link: {bot_str}\n\n"
        "⚡━━━━━━━━━━━━━━━━⚡\n"
        "⬇️ Ki change korte chao?"
    )


def _settings_markup():
    auto_del = _group_settings.get("auto_delete", True)
    grp_send = _group_settings.get("group_otp_send", True)
    grp_tag = _group_settings.get("group_tag", "BOT")
    auto_label = "⏱️ Auto Delete: 🟢 ON" if auto_del else "⏱️ Auto Delete: 🔴 OFF"
    grp_send_label = "📤 Group Send: 🟢 ON" if grp_send else "📤 Group Send: 🔴 OFF"
    markup = types.InlineKeyboardMarkup(row_width=2)
    markup.add(
        types.InlineKeyboardButton("🔗 Group Link", callback_data="grp_setlink"),
        types.InlineKeyboardButton("🆔 Group Chat ID", callback_data="grp_setid"),
    )
    markup.add(
        types.InlineKeyboardButton(auto_label, callback_data="set_autodel"),
        types.InlineKeyboardButton("🗑️ Remove Group", callback_data="grp_remove"),
    )
    markup.add(
        types.InlineKeyboardButton(grp_send_label, callback_data="toggle_grp_send"),
    )
    markup.add(
        types.InlineKeyboardButton(f"⚜ Number Tag: {grp_tag}", callback_data="set_group_tag"),
    )
    markup.add(
        types.InlineKeyboardButton("📢 Join Channel", callback_data="set_channel2"),
        types.InlineKeyboardButton("🤖 Bot Link", callback_data="set_botlink"),
    )
    return markup


def _show_settings(message):
    bot.send_message(
        message.chat.id,
        _settings_text(message.from_user.id),
        reply_markup=_settings_markup(),
        parse_mode="HTML",
    )


def _show_settings_inline(call):
    try:
        bot.edit_message_text(
            _settings_text(call.from_user.id),
            call.message.chat.id,
            call.message.message_id,
            reply_markup=_settings_markup(),
            parse_mode="HTML",
        )
    except Exception:
        pass


def _show_group_settings(message):
    _show_settings(message)


def _show_group_settings_inline(call):
    _show_settings_inline(call)


def _grp_get_link(message):
    uid = message.from_user.id
    if uid not in ADMIN_IDS:
        return
    if _is_back(message.text):
        _go_admin_panel(message)
        return
    if _intercept_menu_btn(message):
        return
    link = (message.text or "").strip()
    if not link.startswith("https://t.me/") and not link.startswith("http://"):
        msg = bot.send_message(
            message.chat.id,
            "❌ Valid Telegram link dao:\n<i>Example: https://t.me/aR_OTP_rcv</i>",
            reply_markup=_back_admin_kb(),
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, _grp_get_link)
        return
    _admin_settings.setdefault(str(uid), {})["otp_group_link"] = link
    save_admin_settings()
    _group_settings["otp_group_link"] = link
    save_group_settings()
    _go_admin_panel(
        message,
        f"✅🔥 <b>GROUP LINK UPDATED!</b>\n\n"
        f"🔗 <b>Notun Link:</b> {link}",
    )


def _grp_get_id(message):
    uid = message.from_user.id
    if uid not in ADMIN_IDS:
        return
    if _is_back(message.text):
        _go_admin_panel(message)
        return
    if _intercept_menu_btn(message):
        return
    raw = (message.text or "").strip()
    try:
        gid = int(raw)
    except ValueError:
        msg = bot.send_message(
            message.chat.id,
            "❌ Valid Chat ID dao (number):\n<i>Example: -1001234567890</i>",
            reply_markup=_back_admin_kb(),
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, _grp_get_id)
        return
    # Save per-admin; also update global if super admin
    _admin_settings.setdefault(str(uid), {})["otp_group_id"] = gid
    save_admin_settings()
    if is_super_admin(uid):
        _group_settings["otp_group_id"] = gid
        save_group_settings()
    _go_admin_panel(
        message,
        f"✅🔥 <b>GROUP CHAT ID UPDATED!</b>\n\n"
        f"🆔 <b>Notun Chat ID:</b> <code>{gid}</code>\n\n"
        f"<i>Shudhu tomar settings-e update hoyeche.</i>",
    )


def _sett_get_channel2(message):
    uid = message.from_user.id
    if uid not in ADMIN_IDS:
        return
    if _is_back(message.text):
        _go_admin_panel(message)
        return
    if _intercept_menu_btn(message):
        return
    link = (message.text or "").strip()
    if not link.startswith("https://") and not link.startswith("http://"):
        msg = bot.send_message(
            message.chat.id,
            "❌ Valid link dao:\n<i>Example: https://t.me/aR_OTP_rcv</i>",
            reply_markup=_back_admin_kb(),
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, _sett_get_channel2)
        return
    _admin_settings.setdefault(str(uid), {})["channel2"] = link
    save_admin_settings()
    _group_settings["channel2"] = link
    save_group_settings()
    _go_admin_panel(
        message,
        f"✅ <b>JOIN CHANNEL UPDATED!</b>\n\n"
        f"📢 <b>Notun Link:</b> {link}",
    )


def _sett_get_botlink(message):
    uid = message.from_user.id
    if uid not in ADMIN_IDS:
        return
    if _is_back(message.text):
        _go_admin_panel(message)
        return
    if _intercept_menu_btn(message):
        return
    link = (message.text or "").strip()
    if not link.startswith("https://") and not link.startswith("http://"):
        msg = bot.send_message(
            message.chat.id,
            "❌ Valid link dao:\n<i>Example: https://t.me/ar_otp_bot</i>",
            reply_markup=_back_admin_kb(),
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, _sett_get_botlink)
        return
    _admin_settings.setdefault(str(uid), {})["bot_link"] = link
    save_admin_settings()
    _group_settings["bot_link"] = link
    save_group_settings()
    _go_admin_panel(
        message,
        f"✅ <b>BOT LINK UPDATED!</b>\n\n"
        f"🤖 <b>Notun Link:</b> {link}",
    )


def _sett_get_support_id(message):
    uid = message.from_user.id
    if uid not in ADMIN_IDS:
        return
    if _is_back(message.text):
        _go_admin_panel(message)
        return
    if _intercept_menu_btn(message):
        return
    val = (message.text or "").strip()
    if not val:
        msg = bot.send_message(
            message.chat.id,
            "❌ Valid Support ID dao:\n<i>Example: @support_user ba 123456789</i>",
            reply_markup=_back_admin_kb(),
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, _sett_get_support_id)
        return
    _group_settings["support_id"] = val
    save_group_settings()
    _go_admin_panel(
        message,
        f"✅ <b>SUPPORT ID UPDATED!</b>\n\n"
        f"📞 <b>Notun Support ID:</b> <code>{val}</code>",
    )


def _sett_get_group_tag(message):
    uid = message.from_user.id
    if uid not in ADMIN_IDS:
        return
    if _is_back(message.text):
        _go_admin_panel(message)
        return
    if _intercept_menu_btn(message):
        return
    val = (message.text or "").strip().upper()
    if not val or len(val) > 20:
        msg = bot.send_message(
            message.chat.id,
            "❌ Valid tag dao (max 20 char):\n<i>Example: ATIK, BOT, KING</i>",
            reply_markup=_back_admin_kb(),
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, _sett_get_group_tag)
        return
    _group_settings["group_tag"] = val
    save_group_settings()
    _go_admin_panel(
        message,
        f"✅ <b>NUMBER TAG UPDATED!</b>\n\n"
        f"⚜ <b>Notun Tag:</b> <code>{val}</code>\n"
        f"📱 Preview: <b>245⚜{val}⚜5660</b>\n\n"
        f"<i>Ekhon theke group-e number ei format-e dekhabe!</i>",
    )


_pending_admin_uid = {}  # {requester_uid: new_uid}


def _admin_add_get_id(message):
    uid = message.from_user.id
    if uid not in ADMIN_IDS:
        return
    if _is_back(message.text):
        _go_admin_panel(message)
        return
    if _intercept_menu_btn(message):
        return
    raw = (message.text or "").strip()
    try:
        new_uid = int(raw)
    except ValueError:
        msg = bot.send_message(
            message.chat.id,
            "❌ Valid Telegram User ID dao (shudhu number):\n<i>Example: 123456789</i>",
            reply_markup=_back_admin_kb(),
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, _admin_add_get_id)
        return
    if new_uid in SUPER_ADMIN_IDS:
        _go_admin_panel(message, "⚠️ <b>Ei user already Super Admin ache!</b>")
        return
    _pending_admin_uid[uid] = new_uid
    dur_kb = types.InlineKeyboardMarkup(row_width=3)
    dur_kb.add(
        types.InlineKeyboardButton("1 Mash", callback_data=f"aadur:{new_uid}:1"),
        types.InlineKeyboardButton("2 Mash", callback_data=f"aadur:{new_uid}:2"),
        types.InlineKeyboardButton("3 Mash", callback_data=f"aadur:{new_uid}:3"),
    )
    dur_kb.add(
        types.InlineKeyboardButton("❌ Cancel", callback_data="aadur_cancel"),
    )
    raw_n = user_names.get(str(new_uid), "")
    name_str = raw_n if isinstance(raw_n, str) else raw_n.get("first_name", str(new_uid))
    name_str = name_str or str(new_uid)
    bot.send_message(
        message.chat.id,
        f"👑 <b>Admin Duration Select koro</b>\n\n"
        f"🔹 <b>User:</b> {name_str} [<code>{new_uid}</code>]\n\n"
        f"Koto mash admin thakbe?",
        reply_markup=dur_kb,
        parse_mode="HTML",
    )


def _show_remove_admin(message):
    removable = [a for a in ADMIN_IDS if a not in SUPER_ADMIN_IDS]
    if not removable:
        bot.send_message(
            message.chat.id,
            "ℹ️ <b>Remove korar moto kono extra admin nei.</b>\n\n"
            "<i>Super Admin remove kora jabe na.</i>",
            reply_markup=_back_admin_kb(),
            parse_mode="HTML",
        )
        return
    markup = types.InlineKeyboardMarkup(row_width=1)
    for aid in removable:
        raw = user_names.get(str(aid), "")
        if isinstance(raw, dict):
            name = raw.get("first_name", "") or str(aid)
        else:
            name = raw or str(aid)
        markup.add(types.InlineKeyboardButton(
            f"🗑️ {name} [{aid}]", callback_data=f"rmadmin:{aid}"
        ))
    bot.send_message(
        message.chat.id,
        "🗑️ <b>Remove Admin</b>\n\n"
        "⚡━━━━━━━━━━━━━━━━⚡\n"
        "Niche theke admin select koro:\n\n"
        "<i>⚠️ Super Admin remove kora jabe na.</i>",
        reply_markup=markup,
        parse_mode="HTML",
    )


_admin_panel_last: dict[int, float] = {}
_admin_panel_lock = threading.Lock()

# ── Resend stop flag ───────────────────────────────────────────────────────────
_resend_running = False
_resend_stop    = False


def _resend_old_otps(message):
    """Fetch today's real OTPs from ALL panels and forward to group (max 50 total)."""
    global _resend_running, _resend_stop
    uid = message.from_user.id
    cid = message.chat.id
    grp = get_admin_setting(uid, "otp_group_id", None) or get_otp_group_id()

    if not grp:
        bot.send_message(cid,
            "❌ <b>Group set kora nei!</b>\nSettings → OTP Group set koro.",
            parse_mode="HTML")
        return

    if _resend_running:
        bot.send_message(cid,
            "⚠️ <b>Resend already cholche!</b>\n"
            "🛑 <b>Purano OTP Bondho</b> button diye age bondo koro.",
            parse_mode="HTML")
        return

    wait_msg = bot.send_message(
        cid,
        "⏳ <b>Sob panel theke ajer OTP fetch korchi...</b>\n"
        "<i>(Shudhu real SMS OTP — fake/range data bade jabe)</i>",
        parse_mode="HTML",
    )

    _resend_running = True
    _resend_stop    = False

    def _do_resend():
        global _resend_running, _resend_stop
        all_found = {}

        static_fetchers = [
            ("P1", fetch_panel1), ("P2", fetch_panel2),
            ("P3", fetch_panel3), ("P4", fetch_panel4),
            ("P5", fetch_panel5), ("P6", fetch_panel6),
        ]
        for pid, fetcher in static_fetchers:
            if _resend_stop:
                break
            try:
                result = fetcher()
                all_found.update(result)
                print(f"[RESEND] {pid}: {len(result)} real OTPs")
            except Exception as e:
                print(f"[RESEND] {pid} error: {e}")

        for panel in list(_dynamic_panels):
            if _resend_stop:
                break
            try:
                result = _universal_fetch(panel)
                all_found.update(result)
                print(f"[RESEND] {panel['id']}: {len(result)} real OTPs")
            except Exception as e:
                print(f"[RESEND] {panel['id']} error: {e}")

        try:
            bot.delete_message(cid, wait_msg.message_id)
        except Exception:
            pass

        if _resend_stop:
            bot.send_message(cid, "🛑 <b>Resend bondho kora hoise!</b>", parse_mode="HTML")
            _resend_running = False
            return

        if not all_found:
            bot.send_message(
                cid,
                "⚠️ <b>Kono real OTP panel e nai!</b>\n"
                "<i>Fake/range data baad dewa hoise. Shudhu real SMS OTP count hoise.</i>",
                parse_mode="HTML",
            )
            _resend_running = False
            return

        MAX_SEND = 50
        items  = list(all_found.values())[:MAX_SEND]
        total  = len(all_found)
        sent   = 0
        failed = 0

        bot.send_message(
            cid,
            f"📤 <b>{total}টা real OTP paoa gese!</b>\n"
            f"<i>Max {MAX_SEND}টা pathano hobe.</i>",
            parse_mode="HTML",
        )

        for number, otp_val, sms_txt, service in items:
            if _resend_stop:
                break
            try:
                send_otp_message(grp, otp_val, number, "—", service, sms_txt or "")
                sent += 1
                time.sleep(0.4)
            except Exception as e:
                failed += 1
                print(f"[RESEND] Send error {number}: {e}")

        _resend_running = False
        status_icon = "🛑 Bondho kora hoise!" if _resend_stop else "✅ Shesh!"
        bot.send_message(
            cid,
            f"{status_icon}\n\n"
            f"📊 <b>Mোট real OTP:</b> {total}\n"
            f"📤 <b>Sent:</b> {sent}\n"
            f"❌ <b>Failed:</b> {failed}",
            parse_mode="HTML",
        )

    threading.Thread(target=_do_resend, daemon=True).start()


def _go_admin_panel(message, text="🔥 <b>ADMIN PANEL</b>"):
    uid = message.from_user.id
    chat_id = message.chat.id
    now = time.time()
    with _admin_panel_lock:
        if now - _admin_panel_last.get(chat_id, 0) < 2.0:
            return
        _admin_panel_last[chat_id] = now
    m_admin = types.ReplyKeyboardMarkup(resize_keyboard=True)
    m_admin.add("➕ 𝗡𝘂𝗺𝗯𝗮𝗿 𝗔𝗱𝗱", "🗑️ 𝗦𝗼𝗯 𝗖𝗹𝗲𝗮𝗿")
    m_admin.add("🔥📢 𝗕𝗿𝗼𝗮𝗱𝗰𝗮𝘀𝘁", "⚡👥 𝗨𝘀𝗲𝗿 𝗖𝗼𝘂𝗻𝘁")
    m_admin.add("📋👥 𝗨𝘀𝗲𝗿 𝗟𝗶𝘀𝘁", "📈 𝗢𝗧𝗣 𝗦𝘁𝗮𝘁𝘀")
    m_admin.add("🎭 𝗗𝗘𝗠𝗢 𝗢𝗧𝗣")
    m_admin.add("➕ 𝗔𝗱𝗱 𝗣𝗮𝗻𝗲𝗹", "🗑️ 𝗥𝗲𝗺𝗼𝘃𝗲 𝗣𝗮𝗻𝗲𝗹")
    m_admin.add("➕ 𝗔𝗱𝗱 𝗦𝗲𝗿𝘃𝗶𝗰𝗲", "🗑️ 𝗥𝗲𝗺𝗼𝘃𝗲 𝗦𝗲𝗿𝘃𝗶𝗰𝗲")
    m_admin.add("📊 𝗣𝗮𝗻𝗲𝗹𝘀", "🔍 𝗧𝗲𝘀𝘁 𝗣𝗮𝗻𝗲𝗹")
    m_admin.add("📤 𝗣𝘂𝗿𝗮𝗻𝗼 𝗢𝗧𝗣 𝗚𝗿𝘂𝗽𝗲 𝗦𝗲𝗻𝗱", "🛑 𝗣𝘂𝗿𝗮𝗻𝗼 𝗢𝗧𝗣 𝗕𝗼𝗻𝗱𝗵𝗼")
    if is_super_admin(uid):
        m_admin.add("👑 𝗔𝗱𝗱 𝗔𝗱𝗺𝗶𝗻", "🗑️ 𝗥𝗲𝗺𝗼𝘃𝗲 𝗔𝗱𝗺𝗶𝗻")
        m_admin.add("📞 𝗦𝘂𝗽𝗽𝗼𝗿𝘁 𝗜𝗗")
    m_admin.add("⚙️ 𝗦𝗲𝘁𝘁𝗶𝗻𝗴𝘀")
    m_admin.add("✏️ 𝗘𝗱𝗶𝘁 𝗠𝗲𝘀𝘀𝗮𝗴𝗲𝘀")
    m_admin.add("📡 𝗩𝟮 𝗠𝗲𝘀𝘀𝗮𝗴𝗲 𝗙𝗼𝗿𝗺𝗮𝘁")
    m_admin.add("⬅️🔙 𝗨𝘀𝗲𝗿 𝗠𝗲𝗻𝘂")
    bot.send_message(
        message.chat.id,
        text,
        reply_markup=m_admin,
        parse_mode="HTML",
    )


# ── Edit Message Templates ──────────────────────────────────────────────────────

def _show_edit_messages_menu(message):
    markup = types.InlineKeyboardMarkup(row_width=1)
    for key, label in _TEMPLATE_LABELS.items():
        markup.add(types.InlineKeyboardButton(label, callback_data=f"editmsg:{key}"))
    markup.add(types.InlineKeyboardButton("🔄 সব Default এ Reset করো", callback_data="editmsg_reset_all"))
    bot.send_message(
        message.chat.id,
        "✏️🔥 <b>মেসেজ ফরমেট এডিট</b> 🔥✏️\n\n"
        "কোন মেসেজ এডিট করতে চাও সিলেক্ট করো:",
        reply_markup=markup,
        parse_mode="HTML",
    )


def _ask_new_template(call, key):
    import html as _html
    label = _TEMPLATE_LABELS.get(key, key)
    vars_hint = _TEMPLATE_VARS.get(key, "")
    current = get_template(key)
    # Escape HTML so tags inside the template don't break the <code> block display
    current_escaped = _html.escape(current[:600])
    uid = call.from_user.id
    _edit_template_state[uid] = {"key": key, "msg_id": call.message.message_id}
    try:
        bot.delete_message(call.message.chat.id, call.message.message_id)
    except Exception:
        pass
    try:
        msg = bot.send_message(
            call.message.chat.id,
            f"✏️ <b>{label}</b>\n\n"
            f"📌 <b>ব্যবহারযোগ্য ভেরিয়েবল:</b>\n<code>{vars_hint}</code>\n\n"
            f"📄 <b>বর্তমান ফরমেট:</b>\n<code>{current_escaped}</code>\n\n"
            f"⬇️ <b>নতুন ফরমেট লিখো:</b>\n"
            f"<i>(HTML ট্যাগ সাপোর্টেড: &lt;b&gt;, &lt;i&gt;, &lt;code&gt;, &lt;/b&gt; etc.)</i>",
            reply_markup=_back_admin_kb(),
            parse_mode="HTML",
        )
    except Exception as e:
        print(f"[TEMPLATE-ASK] ❌ Failed to send prompt for uid={uid}: {e}")
        try:
            msg = bot.send_message(
                call.message.chat.id,
                f"✏️ {label}\n\n⬇️ নতুন ফরমেট লিখো:",
                reply_markup=_back_admin_kb(),
            )
        except Exception:
            return
    bot.register_next_step_handler(msg, _save_new_template)


def _save_new_template(message):
    uid = message.from_user.id
    try:
        if _is_back(message.text):
            _edit_template_state.pop(uid, None)
            _admin_panel_last.pop(message.chat.id, None)
            _go_admin_panel(message)
            return
        if _intercept_menu_btn(message):
            _edit_template_state.pop(uid, None)
            return
        state = _edit_template_state.pop(uid, None)
        if not state:
            _admin_panel_last.pop(message.chat.id, None)
            _go_admin_panel(message)
            return
        key = state["key"]
        new_text = message.text or ""
        if not new_text.strip():
            msg = bot.send_message(
                message.chat.id,
                "❌ খালি রাখা যাবে না। আবার লিখো:",
                reply_markup=_back_admin_kb(),
            )
            _edit_template_state[uid] = state
            bot.register_next_step_handler(msg, _save_new_template)
            return

        # ── Validate: only check for broken brace syntax, accept any {variable} ──
        class _PermissiveDict(dict):
            def __missing__(self, k):
                return f"{{{k}}}"
        _DUMMY_VARS = _PermissiveDict({
            "uname": "TestUser", "uid": "123456789",
            "svc": "INSTAGRAM", "number": "8801712345678",
            "tagged_number": "@+8801712345678", "taged_number": "@+8801712345678",
            "sms_body": "Your OTP is 123456",
            "country": "Bangladesh", "flag": "🇧🇩", "otp": "123456",
            "vname": "TestUser", "text": "Test broadcast",
        })
        try:
            new_text.format_map(_DUMMY_VARS)
        except (ValueError, IndexError) as fmt_err:
            msg = bot.send_message(
                message.chat.id,
                f"❌ <b>টেমপ্লেট ভুল আছে!</b>\n\n"
                f"🔴 <b>Error:</b> <code>{fmt_err}</code>\n\n"
                f"⚠️ <b>সমস্যা:</b> ভুল <code>{{</code> <code>}}</code> ব্যবহার।\n\n"
                f"💡 সাধারণ বন্ধনী দরকার হলে double করো: <code>{{{{</code> এবং <code>}}}}</code>\n\n"
                f"আবার লিখো:",
                reply_markup=_back_admin_kb(),
                parse_mode="HTML",
            )
            _edit_template_state[uid] = state
            bot.register_next_step_handler(msg, _save_new_template)
            return

        _templates[key] = new_text
        save_templates()
        label = _TEMPLATE_LABELS.get(key, key)
        # Reset rate limiter so admin panel shows immediately after save
        _admin_panel_last.pop(message.chat.id, None)
        _go_admin_panel(
            message,
            f"✅🔥 <b>মেসেজ আপডেট হয়েছে!</b>\n\n"
            f"✏️ <b>{label}</b>\n\n"
            f"📄 নতুন ফরমেট সেভ হয়েছে।",
        )
    except Exception as e:
        print(f"[TEMPLATE-SAVE] ❌ Error for uid={uid}: {e}")
        try:
            bot.send_message(
                message.chat.id,
                f"❌ <b>কিছু একটা সমস্যা হয়েছে!</b>\n<code>{e}</code>\n\nআবার চেষ্টা করো।",
                reply_markup=_back_admin_kb(),
                parse_mode="HTML",
            )
        except Exception:
            pass


def _cancel_kb():
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True)
    kb.add("❌ Cancel")
    return kb


def _back_admin_kb():
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True)
    kb.add("🔙 Admin Panel")
    return kb


def _is_back(txt):
    return (txt or "").strip() in ("🔙 Admin Panel", "❌ Cancel")


_ALL_MENU_BTNS = {
    "☎️ 𝗩𝟭 𝗡𝗨𝗠𝗕𝗔𝗥 ☎️", "☎️ 𝗡𝗨𝗠𝗕𝗔𝗥 ☎️", "📡 𝗩𝟮 𝗖𝗼𝗻𝘀𝗼𝗹𝗲",
    "🔄 𝗩𝟮 𝗦𝗪𝗜𝗧𝗖𝗛", "🔴 𝗟𝗜𝗩𝗘 𝗥𝗔𝗡𝗚𝗘", "⌨️ 𝗖𝗨𝗦𝗧𝗢𝗠 𝗥𝗔𝗡𝗚𝗘", "🔙 𝗩𝟭 𝗦𝗪𝗜𝗧𝗖𝗛",
    "📊 𝗦𝗧𝗢𝗖𝗞", "📞 𝗦𝗔𝗣𝗢𝗥𝗧",
    "⚙️ 𝗔𝗗𝗠𝗜𝗡 𝗣𝗔𝗡𝗘𝗟 ⚙️", "🔙 Main Menu",
    "➕ 𝗡𝘂𝗺𝗯𝗮𝗿 𝗔𝗱𝗱", "🗑️ 𝗦𝗼𝗯 𝗖𝗹𝗲𝗮𝗿",
    "🔥📢 𝗕𝗿𝗼𝗮𝗱𝗰𝗮𝘀𝘁", "⚡👥 𝗨𝘀𝗲𝗿 𝗖𝗼𝘂𝗻𝘁",
    "📋👥 𝗨𝘀𝗲𝗿 𝗟𝗶𝘀𝘁", "📈 𝗢𝗧𝗣 𝗦𝘁𝗮𝘁𝘀", "🎭 𝗗𝗘𝗠𝗢 𝗢𝗧𝗣",
    "➕ 𝗔𝗱𝗱 𝗣𝗮𝗻𝗲𝗹", "🗑️ 𝗥𝗲𝗺𝗼𝘃𝗲 𝗣𝗮𝗻𝗲𝗹",
    "➕ 𝗔𝗱𝗱 𝗦𝗲𝗿𝘃𝗶𝗰𝗲", "🗑️ 𝗥𝗲𝗺𝗼𝘃𝗲 𝗦𝗲𝗿𝘃𝗶𝗰𝗲",
    "📊 𝗣𝗮𝗻𝗲𝗹𝘀", "🔍 𝗧𝗲𝘀𝘁 𝗣𝗮𝗻𝗲𝗹", "👑 𝗔𝗱𝗱 𝗔𝗱𝗺𝗶𝗻", "🗑️ 𝗥𝗲𝗺𝗼𝘃𝗲 𝗔𝗱𝗺𝗶𝗻",
    "📞 𝗦𝘂𝗽𝗽𝗼𝗿𝘁 𝗜𝗗",
    "⚙️ 𝗦𝗲𝘁𝘁𝗶𝗻𝗴𝘀", "✏️ 𝗘𝗱𝗶𝘁 𝗠𝗲𝘀𝘀𝗮𝗴𝗲𝘀", "📡 𝗩𝟮 𝗠𝗲𝘀𝘀𝗮𝗴𝗲 𝗙𝗼𝗿𝗺𝗮𝘁", "⬅️🔙 𝗨𝘀𝗲𝗿 𝗠𝗲𝗻𝘂",
    "🔙 𝗔𝗗𝗠𝗜𝗡 𝗣𝗔𝗡𝗘𝗟", "🔙 Admin Panel", "🔙 Admin Menu",
}


def _intercept_menu_btn(message):
    """If user pressed any known menu/admin button while in a step flow,
    route it to text_handler so it is handled correctly.
    Returns True if intercepted, False otherwise."""
    txt = (message.text or "").strip()
    if txt in _ALL_MENU_BTNS:
        text_handler(message)
        return True
    return False


def process_auto_add(message):
    svc = (message.text or "").strip().lower()
    if svc == "❌ cancel":
        _go_admin_panel(message)
        return
    if svc not in stock:
        m = types.ReplyKeyboardMarkup(resize_keyboard=True)
        m.add("facebook", "instagram", "whatsapp", "telegram", "binance", "pc clone")
        m.add("❌ Cancel")
        msg = bot.send_message(
            message.chat.id,
            " <b>Vul service! Abar choose koro:</b>",
            reply_markup=m,
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, process_auto_add)
        return
    msg = bot.send_message(
        message.chat.id,
        f"🔥 <b>{svc.upper()}</b>\n\n"
        f"📝 <b>Slot name dao:</b>\n"
        f"<i>Udharan: Mali 1, Germany 2, India 3</i>",
        reply_markup=_cancel_kb(),
        parse_mode="HTML",
    )
    bot.register_next_step_handler(msg, lambda m: ask_numbers_for_slot(m, svc))


def ask_numbers_for_slot(message, svc):
    slot_name = (message.text or "").strip()
    if slot_name == "❌ Cancel":
        _go_admin_panel(message)
        return
    if not slot_name:
        msg = bot.send_message(
            message.chat.id,
            "❌ Slot name dao:",
            reply_markup=_cancel_kb(),
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, lambda m: ask_numbers_for_slot(m, svc))
        return
    msg = bot.send_message(
        message.chat.id,
        f"✅ Slot: <b>{slot_name}</b>\n\n"
        f"📱 Ekhon <b>{svc.upper()}</b> er number gulo pathao:\n"
        f"<i>(Newline ba comma diye alag koro)</i>",
        reply_markup=_cancel_kb(),
        parse_mode="HTML",
    )
    bot.register_next_step_handler(msg, lambda m: finalize_auto_add(m, svc, slot_name))


def finalize_auto_add(message, svc, slot_name=None):
    global stock
    uid = message.from_user.id
    if (message.text or "").strip() == "❌ Cancel":
        _go_admin_panel(message)
        return
    nums = [n.strip() for n in re.split(r"[,\n\r]", message.text) if n.strip()]
    if slot_name:
        if slot_name not in stock[svc]:
            stock[svc][slot_name] = []
        for num in nums:
            stock[svc][slot_name].append(num)
        added_count = len(nums)
    else:
        added_count = 0
        for num in nums:
            c_name, _ = get_country_details(num)
            if c_name == "Unknown":
                continue
            if c_name not in stock[svc]:
                stock[svc][c_name] = []
            stock[svc][c_name].append(num)
            added_count += 1
    save_stock()
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    markup.add("➕ Aro Add koro", "🔙 Admin Menu")
    bot.send_message(
        message.chat.id,
        f"✅🔥 <b>DONE!</b>\n\n"
        f"🗂 <b>Slot:</b> {slot_name or 'Auto'}\n"
        f"📱 <b>Added:</b> {added_count} টি number",
        reply_markup=markup,
        parse_mode="HTML",
    )
    bot.register_next_step_handler(
        bot.send_message(message.chat.id, "⬇️ Ki korbe?", parse_mode="HTML"),
        lambda m: _after_add_handler(m, svc),
    )


def _after_add_handler(message, last_svc):
    txt = (message.text or "").strip()
    if txt == "➕ Aro Add koro":
        msg = bot.send_message(
            message.chat.id,
            f"📝 <b>Notun slot name dao:</b>\n<i>Udharan: Mali 2, Germany 3</i>",
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, lambda m: ask_numbers_for_slot(m, last_svc))
    else:
        _go_admin_panel(message)


# ── Heartbeat / watchdog ───────────────────────────────────────────────────────



# ── Start ─────────────────────────────────────────────────────────────────────

try:
    requests.get(
        f"https://api.telegram.org/bot{API_TOKEN}/deleteWebhook?drop_pending_updates=true",
        timeout=10,
    )
    print("[START] Webhook cleared.")
except Exception as e:
    print(f"[START] Webhook clear failed: {e}")

time.sleep(3)

threading.Thread(target=panel1_monitor, daemon=True).start()
threading.Thread(target=panel2_monitor, daemon=True).start()
threading.Thread(target=panel3_monitor, daemon=True).start()
threading.Thread(target=panel4_monitor, daemon=True).start()
threading.Thread(target=panel5_monitor, daemon=True).start()
threading.Thread(target=panel6_monitor, daemon=True).start()
threading.Thread(target=demo_monitor, daemon=True).start()

for _dp in _dynamic_panels:
    _start_dynamic_panel(_dp)
    print(f"[DYN] Loaded saved panel: {_dp['id']} ({_dp['host']})")

# ── Load builtin extra panels (hardcoded) ─────────────────────────────────────
_existing_bp_ids = {p["id"] for p in _dynamic_panels}
_new_builtins_added = False
for _bp in _BUILTIN_PANELS:
    if _bp["id"] not in _existing_bp_ids:
        _dynamic_panels.append(_bp)
        _new_builtins_added = True
    _start_dynamic_panel(_bp)
    print(f"[BUILTIN] Loaded panel: {_bp['id']} ({_bp['host']} / {_bp['username']})")
# Always persist so new BUILTIN panels (like fastx1) survive across restarts
save_dynamic_panels()

# ── IVA SMS startup login check ───────────────────────────────────────────────
def _iva_startup_check():
    """After bot starts, try to login to IVA panel. If fails, notify admins."""
    time.sleep(15)  # wait for bot polling to start
    iva = _iva_find_panel()
    if not iva:
        return
    pid = iva["id"]
    # If already logged in (scraper exists), skip
    if _iva_scrapers.get(pid):
        return
    print(f"[IVA-STARTUP] Trying initial login for {pid}...")
    ok = _iva_login(iva)
    if ok:
        print(f"[IVA-STARTUP] ✅ IVA panel login OK")
        return
    # Login failed — notify all super admins
    print(f"[IVA-STARTUP] ❌ IVA panel login FAILED — notifying admins")
    for admin_uid in SUPER_ADMIN_IDS:
        try:
            bot.send_message(
                admin_uid,
                "⚠️ <b>IVA SMS Panel (bp10) login hoi nai!</b>\n\n"
                "Railway-er IP theke email/password login block hoye ache (Cloudflare).\n\n"
                "🍪 <b>Cookie diye fix koro:</b>\n"
                "1. Phone/PC-te Chrome-e <b>ivasms.com</b> login koro\n"
                "2. Browser-er Developer Tools open koro\n"
                "   • PC: F12 → Application → Cookies → ivasms.com\n"
                "   • Phone: Chrome menu → Desktop site, tahole F12\n"
                "3. <code>laravel_session</code> cookie-er value copy koro\n"
                "4. Bot-e pathao: /ivacookie\n\n"
                "<i>Command: <code>/ivacookie bp10</code></i>",
                parse_mode="HTML",
            )
        except Exception as e:
            print(f"[IVA-STARTUP] Notify error for {admin_uid}: {e}")

threading.Thread(target=_iva_startup_check, daemon=True).start()


# ── Web helper cookie watcher ──────────────────────────────────────────────
import json as _json_mod

_IVA_COOKIE_PENDING_FILE = os.path.join(os.path.dirname(__file__), "iva_cookie_pending.txt")


def _iva_web_cookie_watcher():
    """Watches for cookie file written by the web helper page and applies it."""
    while True:
        try:
            if os.path.exists(_IVA_COOKIE_PENDING_FILE):
                with open(_IVA_COOKIE_PENDING_FILE, "r") as _f:
                    _data = _json_mod.load(_f)
                os.remove(_IVA_COOKIE_PENDING_FILE)
                _panel_id = _data.get("panel", "bp10")
                _cookie_str = _data.get("cookie", "").strip()
                if _cookie_str and "=" in _cookie_str:
                    print(f"[WEB-COOKIE] Got cookie for {_panel_id} via helper page")
                    # Update dynamic_panels
                    for _p in _dynamic_panels:
                        if _p["id"] == _panel_id:
                            _p["cookie_str"] = _cookie_str
                            save_dynamic_panels()
                            break
                    # Update BUILTIN_PANELS in-memory
                    for _p in _BUILTIN_PANELS:
                        if _p["id"] == _panel_id:
                            _p["cookie_str"] = _cookie_str
                            break
                    _iva_scrapers.pop(_panel_id, None)
                    # Try to reconnect
                    _panel_obj = _iva_find_panel(_panel_id)
                    if _panel_obj:
                        ok = _iva_login(_panel_obj)
                        status = "✅ Login SUCCESSFUL! OTP ashle group-e pathabe." if ok else "❌ Cookie kaj koreni — abar try koro."
                        print(f"[WEB-COOKIE] {_panel_id}: {status}")
                        for _admin in SUPER_ADMIN_IDS:
                            try:
                                bot.send_message(_admin,
                                    f"🍪 <b>Web Helper Cookie — {_panel_id}</b>\n{status}",
                                    parse_mode="HTML")
                            except Exception:
                                pass
        except Exception as _e:
            print(f"[WEB-COOKIE] Error: {_e}")
        time.sleep(10)


threading.Thread(target=_iva_web_cookie_watcher, daemon=True).start()

print("🔥 AR OTP BOT is running with 15-PANEL AUTO OTP MONITOR... 🔥")
print("   ▸ Panel 1: Mahofuza        (91.232.105.47)")
print("   ▸ Panel 2: Sagardas50      (94.23.31.29)")
print("   ▸ Panel 3: Rabbi1_FD       (168.119.13.175)")
print("   ▸ Panel 4: Rabbi12         (144.217.71.192)")
print("   ▸ Panel 5: Rabbi12_v2      (51.75.144.178)")
print("   ▸ Panel 6: TrueSMS/Ranges  (truesms.net)")
print("   ▸ BP1: Mahofuza12          (139.99.69.196)")
print("   ▸ BP2: Rabbi12             (139.99.9.4)")
print("   ▸ BP3: Rabbi12             (54.36.173.235)")
print("   ▸ BP4: Rabbi5              (54.39.104.241)")
print("   ▸ BP5: mahofuza            (213.32.24.208)")
print("   ▸ BP6: Rabbi200            (15.235.182.3 /konekta)")
print("   ▸ BP7: Rabbi12             (nexor-iprn.com)")
print("   ▸ BP8: Rabbi12             (51.77.52.79)")
print("   ▸ BP9: Dasbabu50_FD        (51.210.208.26)")
print("   ▸ BP10: mdrashub2          (ivasms.com)")
print("   ▸ BP11: Rabbi12            (139.99.68.231)")


def _clear_webhook():
    try:
        requests.get(
            f"https://api.telegram.org/bot{API_TOKEN}/deleteWebhook?drop_pending_updates=true",
            timeout=10,
        )
    except Exception:
        pass


while True:
    try:
        _clear_webhook()
        time.sleep(3)
        bot.infinity_polling(
            timeout=60,
            long_polling_timeout=60,
            allowed_updates=["message", "callback_query"],
            none_stop=True,
            restart_on_change=False,
        )
    except requests.exceptions.ReadTimeout:
        print("[POLLING] ReadTimeout — restarting in 5s...")
        time.sleep(5)
    except requests.exceptions.ConnectionError:
        print("[POLLING] ConnectionError — restarting in 10s...")
        time.sleep(10)
    except Exception as e:
        err_str = str(e)
        if "409" in err_str or "Conflict" in err_str:
            print("[POLLING] 409 Conflict (another instance running) — waiting 30s...")
            time.sleep(30)
        else:
            print(f"[POLLING] Error: {e} — restarting in 5s...")
            time.sleep(5)
