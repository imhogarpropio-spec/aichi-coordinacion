# init_db.py
from app import create_app
from models import db  # ajusta el import si tu 'db' vive en otro módulo

app = create_app()

with app.app_context():
    db.create_all()
    print("✅ Tablas creadas en la base de datos de Render")
