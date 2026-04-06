import json, hashlib

data = json.load(open('users.json', encoding='utf-8'))

def verify(pwd, stored_hash, salt):
    h = hashlib.pbkdf2_hmac('sha256', pwd.encode(), salt.encode(), 200000).hex()
    return h == stored_hash

tests = [
    ('admin',      'Admin@BEAC2026'),
    ('analyste',   'DFX@2026'),
    ('dom_export', 'Export@2026'),
]
for uname, pwd in tests:
    u = next((x for x in data['users'] if x['username']==uname), None)
    if u:
        ok = verify(pwd, u['password_hash'], u['salt'])
        mcp = u['must_change_password']
        print(f"{uname}: {'OK' if ok else 'ECHEC'} | must_change={mcp}")
    else:
        print(f"{uname}: INTROUVABLE")
