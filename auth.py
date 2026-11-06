"""Minimal login: salted PBKDF2 password hashes stored in users.json (no plain-text passwords)."""
import hashlib
import hmac
import json
import re
import secrets

from config import USERS_FILE

USERNAME_RE = re.compile(r"^[A-Za-z0-9_-]{1,32}$")   # also keeps usernames safe to use in file names


def _hash(password, salt_hex):
    return hashlib.pbkdf2_hmac("sha256", password.encode(), bytes.fromhex(salt_hex), 200_000).hex()


def load_users():
    return json.loads(USERS_FILE.read_text()) if USERS_FILE.exists() else {}


def set_user(username, password):
    if not USERNAME_RE.match(username):
        raise ValueError("Username may only contain letters, digits, _ and - (max 32).")
    users = load_users()
    salt = secrets.token_hex(16)
    users[username] = {"salt": salt, "hash": _hash(password, salt)}
    USERS_FILE.write_text(json.dumps(users, indent=2))


def remove_user(username):
    users = load_users()
    users.pop(username, None)
    USERS_FILE.write_text(json.dumps(users, indent=2))


def verify(username, password):
    """True only for a known user with the right password."""
    u = load_users().get(username)
    salt = u["salt"] if u else "00" * 16        # hash anyway so timing doesn't reveal valid usernames
    ok = hmac.compare_digest(_hash(password, salt), u["hash"] if u else "x")
    return bool(u) and ok
