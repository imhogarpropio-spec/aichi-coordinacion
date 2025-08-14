# migrar_delegacion_id.py
from app import create_app
from models import db, Usuario, Delegacion

app = create_app()
with app.app_context():
    # 1) Crear columna si no existe
    try:
        db.session.execute(db.text("ALTER TABLE usuarios ADD COLUMN delegacion_id INTEGER"))
        db.session.commit()
        print("OK: Columna delegacion_id creada")
    except Exception as e:
        # Si ya existe, seguimos
        db.session.rollback()
        print("Aviso:", e)

    # 2) Backfill: pasar del texto a la columna id (por nombre de delegación)
    total = 0
    sin_match = 0
    usuarios = Usuario.query.all()
    for u in usuarios:
        if u.delegacion and not u.delegacion_id:
            d = Delegacion.query.filter(db.func.upper(db.func.trim(Delegacion.nombre)) == u.delegacion.strip().upper()).first()
            if d:
                u.delegacion_id = d.id
                total += 1
            else:
                sin_match += 1
    db.session.commit()
    print(f"Backfill listo. Asignados: {total}. Sin coincidencia: {sin_match}.")
