from app import app, db
from models import Usuario  # Asegúrate de importar el modelo

with app.app_context():
    Usuario.__table__.create(bind=db.engines['usuarios'])
    print("✅ Tabla de usuarios creada correctamente.")
