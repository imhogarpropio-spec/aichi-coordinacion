import os
from dotenv import load_dotenv
from pathlib import Path
from datetime import timedelta

load_dotenv(dotenv_path=Path('aichi.env'))

class Config:
    SECRET_KEY = os.getenv("SECRET_KEY")
    SQLALCHEMY_DATABASE_URI = os.getenv("DATABASE_URI")
    SQLALCHEMY_BINDS = {
        "usuarios": os.getenv("USUARIOS_DATABASE_URI")
    }
    if not SECRET_KEY:
        raise RuntimeError("Falta SECRET_KEY en el .env")
    if not SQLALCHEMY_DATABASE_URI:
        raise RuntimeError("Falta DATABASE_URI en el .env")
    if not SQLALCHEMY_BINDS["usuarios"]:
        raise RuntimeError("Falta USUARIOS_DATABASE_URI en el .env")

    SQLALCHEMY_TRACK_MODIFICATIONS = False

    # Correo (opcional, si usas mail)
    MAIL_SERVER = "smtp.gmail.com"
    MAIL_PORT = 587
    MAIL_USE_TLS = True
    MAIL_USERNAME = os.getenv("MAIL_USERNAME")
    MAIL_PASSWORD = os.getenv("MAIL_PASSWORD")
    DEFAULT_ADMIN_CORREO = os.getenv("DEFAULT_ADMIN_CORREO")
    DEFAULT_ADMIN_PASSWORD = os.getenv("DEFAULT_ADMIN_PASSWORD")
    DEFAULT_ADMIN_GENERO = "H"
    DEFAULT_ADMIN_DELEGACION = "DII-127"

    PERMANENT_SESSION_LIFETIME = timedelta(minutes=30)
