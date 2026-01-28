"""
Script para generar credenciales de usuarios con contraseñas bcrypt
Ejecutar: python generate_credentials.py
"""

import bcrypt
import yaml

def hash_password(password):
    """Genera hash bcrypt de una contraseña"""
    salt = bcrypt.gensalt()
    return bcrypt.hashpw(password.encode(), salt).decode()

def generate_credentials():
    """Genera archivo de credenciales"""
    
    print("=" * 60)
    print("🔑 Generador de Credenciales para Streamlit Auth")
    print("=" * 60)
    
    credentials = {
        "usernames": {}
    }
    
    while True:
        username = input("\nNombre de usuario (o 'listo' para terminar): ").strip()
        
        if username.lower() == "listo":
            break
        
        if not username:
            print("❌ El nombre de usuario no puede estar vacío")
            continue
        
        if username in credentials["usernames"]:
            print("❌ Este usuario ya existe")
            continue
        
        name = input("Nombre completo: ").strip()
        email = input("Email: ").strip()
        password = input("Contraseña: ").strip()
        
        if not password:
            print("❌ La contraseña no puede estar vacía")
            continue
        
        # Generar hash
        password_hash = hash_password(password)
        
        credentials["usernames"][username] = {
            "name": name,
            "email": email,
            "password": password_hash
        }
        
        print(f"✅ Usuario '{username}' creado")
    
    # Mostrar configuración para copiar a secrets.toml
    print("\n" + "=" * 60)
    print("📋 Copia esto en .streamlit/secrets.toml:")
    print("=" * 60)
    
    config = {"credentials": credentials}
    
    print("\n[credentials]")
    for username, data in credentials["usernames"].items():
        print(f'usernames.{username}.email = "{data["email"]}"')
        print(f'usernames.{username}.name = "{data["name"]}"')
        print(f'usernames.{username}.password = "{data["password"]}"')
        print()
    
    # Guardar a archivo también
    with open(".streamlit/credentials.yaml", "w") as f:
        yaml.dump(config, f)
    
    print("✓ Credenciales también guardadas en .streamlit/credentials.yaml")

if __name__ == "__main__":
    generate_credentials()
