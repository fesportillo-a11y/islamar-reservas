"""Genera el hash bcrypt de una contraseña para añadirla a secrets.toml.

Uso:
    python tools/hash_password.py
    python tools/hash_password.py micontraseña123

El hash resultante se pega en la sección [auth.credentials.usernames.<usuario>]
del archivo .streamlit/secrets.toml (o de la pantalla de Secrets en Streamlit Cloud).
"""
import sys
import getpass
import bcrypt


def hash_password(plain: str) -> str:
    return bcrypt.hashpw(plain.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")


def main() -> int:
    if len(sys.argv) >= 2:
        plain = sys.argv[1]
    else:
        plain = getpass.getpass("Escribe la contraseña a hashear (no se mostrará): ")
        plain2 = getpass.getpass("Repite la contraseña: ")
        if plain != plain2:
            print("ERROR: las contraseñas no coinciden.")
            return 1

    if not plain:
        print("ERROR: contraseña vacía.")
        return 1

    h = hash_password(plain)
    print()
    print("Hash bcrypt generado. Cópialo a secrets.toml:")
    print()
    print(f'password = "{h}"')
    print()
    return 0


if __name__ == "__main__":
    sys.exit(main())
