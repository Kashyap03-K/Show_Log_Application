"""
auth.py — simple username/password authentication.

Credentials are stored as salted SHA-256 hashes in users.json (created on
first run from DEFAULT_USERS below). Editing credentials:

  * easiest:  run  python auth.py add <username> <password>
              or   python auth.py remove <username>
              or   python auth.py list
  * or edit users.json directly (it stores hashes, not plain passwords).

This is lightweight protection suitable for a small trusted team on an
internal network — not a hardened public-internet auth system.
"""

import os
import json
import hashlib
import secrets
import sys

_DIR = os.path.dirname(os.path.abspath(__file__))
USERS_FILE = os.path.join(_DIR, "users.json")

# Seed accounts, created on first run. CHANGE THESE PASSWORDS.
DEFAULT_USERS = {
    "admin":  "showlog@2026",
    "user1":  "datahub@2026",
}


def _hash(password, salt):
    return hashlib.sha256((salt + password).encode("utf-8")).hexdigest()


def _load():
    if not os.path.exists(USERS_FILE):
        users = {}
        for name, pw in DEFAULT_USERS.items():
            salt = secrets.token_hex(8)
            users[name] = {"salt": salt, "hash": _hash(pw, salt)}
        with open(USERS_FILE, "w", encoding="utf-8") as f:
            json.dump(users, f, indent=2)
        return users
    with open(USERS_FILE, encoding="utf-8") as f:
        return json.load(f)


def _save(users):
    with open(USERS_FILE, "w", encoding="utf-8") as f:
        json.dump(users, f, indent=2)


def verify(username, password):
    """Return True if username/password match a stored account."""
    if not username or not password:
        return False
    users = _load()
    rec = users.get(username.strip())
    if not rec:
        return False
    return secrets.compare_digest(rec["hash"], _hash(password, rec["salt"]))


def add_user(username, password):
    users = _load()
    salt = secrets.token_hex(8)
    users[username.strip()] = {"salt": salt, "hash": _hash(password, salt)}
    _save(users)


def remove_user(username):
    users = _load()
    if username.strip() in users:
        del users[username.strip()]
        _save(users)
        return True
    return False


def list_users():
    return sorted(_load().keys())


# Command-line management ----------------------------------------------------
if __name__ == "__main__":
    args = sys.argv[1:]
    if not args or args[0] == "list":
        print("Users:", ", ".join(list_users()) or "(none)")
    elif args[0] == "add" and len(args) == 3:
        add_user(args[1], args[2])
        print(f"User '{args[1]}' added/updated.")
    elif args[0] == "remove" and len(args) == 2:
        ok = remove_user(args[1])
        print(f"User '{args[1]}' removed." if ok else "User not found.")
    else:
        print("Usage:")
        print("  python auth.py list")
        print("  python auth.py add <username> <password>")
        print("  python auth.py remove <username>")
