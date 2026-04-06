"""
manage_users.py — Gestion des utilisateurs de l'App DFX (BEAC)
=============================================================
Usage :
    python manage_users.py                    → menu interactif
    python manage_users.py list               → liste les utilisateurs
    python manage_users.py add                → ajouter un utilisateur (interactif)
    python manage_users.py reset <username>   → réinitialiser le mot de passe
    python manage_users.py delete <username>  → supprimer un utilisateur
"""

import json
import hashlib
import secrets
import sys
import os
import getpass

FICHIER_USERS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "users.json")

ROLES = {
    "1": ("admin",           "Administrateur"),
    "2": ("analyste_dfx",   "Analyste DFX"),
    "3": ("superviseur_dom", "Superviseur DOM Export"),
}

# ─── Crypto ───────────────────────────────────────────────────────────────────

def _hash_password(password, salt=None):
    if salt is None:
        salt = secrets.token_hex(16)
    hashed = hashlib.pbkdf2_hmac(
        "sha256", password.encode("utf-8"), salt.encode("utf-8"), 200_000
    ).hex()
    return hashed, salt

# ─── Persistance ──────────────────────────────────────────────────────────────

def _charger():
    if not os.path.exists(FICHIER_USERS):
        print(f"[ERREUR] Fichier introuvable : {FICHIER_USERS}")
        sys.exit(1)
    with open(FICHIER_USERS, "r", encoding="utf-8") as f:
        return json.load(f)

def _sauvegarder(data):
    with open(FICHIER_USERS, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"[OK] Fichier mis à jour : {FICHIER_USERS}")

def _trouver(data, username):
    for u in data["users"]:
        if u["username"] == username:
            return u
    return None

# ─── Actions ──────────────────────────────────────────────────────────────────

def lister():
    data = _charger()
    print(f"\n{'─'*58}")
    print(f"  {'Identifiant':<18} {'Nom affiché':<22} {'Rôle'}")
    print(f"{'─'*58}")
    for u in data["users"]:
        mcp = " [chgt mdp requis]" if u.get("must_change_password") else ""
        print(f"  {u['username']:<18} {u['display_name']:<22} {u['role']}{mcp}")
    print(f"{'─'*58}")
    print(f"  Total : {len(data['users'])} utilisateur(s)\n")


def ajouter():
    data = _charger()
    print("\n── Ajouter un utilisateur ──────────────────────────────")

    username = input("  Identifiant     : ").strip()
    if not username:
        print("[ERREUR] L'identifiant est obligatoire.")
        return
    if _trouver(data, username):
        print(f"[ERREUR] L'identifiant « {username} » existe déjà.")
        return

    display_name = input("  Nom affiché     : ").strip() or username

    print("  Rôles disponibles :")
    for k, (role_id, label) in ROLES.items():
        print(f"    {k}. {label}")
    choix = input("  Choisir le rôle (1/2/3) : ").strip()
    role, role_label = ROLES.get(choix, ("analyste_dfx", "Analyste DFX"))

    password = getpass.getpass("  Mot de passe initial : ")
    if len(password) < 8:
        print("[ERREUR] Le mot de passe doit comporter au moins 8 caractères.")
        return
    confirm = getpass.getpass("  Confirmer le mot de passe : ")
    if password != confirm:
        print("[ERREUR] Les mots de passe ne correspondent pas.")
        return

    hashed, salt = _hash_password(password)
    data["users"].append({
        "username":             username,
        "password_hash":        hashed,
        "salt":                 salt,
        "display_name":         display_name,
        "role":                 role,
        "must_change_password": False,
    })
    _sauvegarder(data)
    print(f"[OK] Utilisateur « {username} » ({role_label}) créé avec succès.")


def reinitialiser(username):
    data = _charger()
    user = _trouver(data, username)
    if not user:
        print(f"[ERREUR] Utilisateur « {username} » introuvable.")
        return

    password = getpass.getpass(f"  Nouveau mot de passe pour « {username} » : ")
    if len(password) < 8:
        print("[ERREUR] Le mot de passe doit comporter au moins 8 caractères.")
        return
    confirm = getpass.getpass("  Confirmer : ")
    if password != confirm:
        print("[ERREUR] Les mots de passe ne correspondent pas.")
        return

    hashed, salt = _hash_password(password)
    user["password_hash"]       = hashed
    user["salt"]                = salt
    user["must_change_password"] = False
    _sauvegarder(data)
    print(f"[OK] Mot de passe de « {username} » réinitialisé.")


def supprimer(username):
    data = _charger()
    user = _trouver(data, username)
    if not user:
        print(f"[ERREUR] Utilisateur « {username} » introuvable.")
        return
    if user["role"] == "admin" and sum(1 for u in data["users"] if u["role"] == "admin") == 1:
        print("[ERREUR] Impossible de supprimer le dernier administrateur.")
        return

    confirm = input(f"  Supprimer définitivement « {username} » ? (oui/non) : ").strip().lower()
    if confirm != "oui":
        print("  Annulé.")
        return

    data["users"] = [u for u in data["users"] if u["username"] != username]
    _sauvegarder(data)
    print(f"[OK] Utilisateur « {username} » supprimé.")


# ─── Menu interactif ──────────────────────────────────────────────────────────

def menu():
    actions = {
        "1": ("Lister les utilisateurs",         lambda: lister()),
        "2": ("Ajouter un utilisateur",           lambda: ajouter()),
        "3": ("Réinitialiser un mot de passe",   lambda: (reinitialiser(input("  Identifiant : ").strip()))),
        "4": ("Supprimer un utilisateur",         lambda: (supprimer(input("  Identifiant : ").strip()))),
        "5": ("Quitter",                          lambda: sys.exit(0)),
    }
    print("\n╔════════════════════════════════════════╗")
    print("║   Gestion des utilisateurs — App DFX  ║")
    print("╚════════════════════════════════════════╝")
    while True:
        print()
        for k, (label, _) in actions.items():
            print(f"  {k}. {label}")
        choix = input("\n  Votre choix : ").strip()
        if choix in actions:
            actions[choix][1]()
        else:
            print("  Choix invalide.")


# ─── Point d'entrée ───────────────────────────────────────────────────────────

if __name__ == "__main__":
    args = sys.argv[1:]
    if not args:
        menu()
    elif args[0] == "list":
        lister()
    elif args[0] == "add":
        ajouter()
    elif args[0] == "reset" and len(args) == 2:
        reinitialiser(args[1])
    elif args[0] == "delete" and len(args) == 2:
        supprimer(args[1])
    else:
        print(__doc__)
