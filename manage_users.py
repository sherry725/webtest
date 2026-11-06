"""Manage logins.
    python manage_users.py add alice        (asks for a password)
    python manage_users.py remove alice
    python manage_users.py list
After adding a user, put alice_work.xlsx and alice_life.xlsx in the data/ folder.
"""
import getpass
import sys

import auth

cmd, *args = sys.argv[1:] or ["list"]
if cmd == "add" and args:
    pw = getpass.getpass(f"Password for {args[0]}: ")
    if pw != getpass.getpass("Repeat password: "):
        sys.exit("Passwords do not match.")
    auth.set_user(args[0], pw)
    print(f"Saved {args[0]}. Put {args[0]}_work.xlsx and {args[0]}_life.xlsx in data/.")
elif cmd == "remove" and args:
    auth.remove_user(args[0])
    print(f"Removed {args[0]}.")
elif cmd == "list":
    print("\n".join(auth.load_users()) or "(no users)")
else:
    sys.exit(__doc__)
