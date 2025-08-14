# backfill_delegacion_id.py  (versión SQLAlchemy 2.x)
from app import create_app
from models import db, Usuario, Delegacion
from sqlalchemy import text

app = create_app()

def get_engine_for_model(model):
    # Flask‑SQLAlchemy 3.x: usa el engine del bind del modelo
    bind_key = getattr(model, "__bind_key__", None) or None
    return db.engines[bind_key]  # dict de engines; None = default

with app.app_context():
    eng_usuarios = get_engine_for_model(Usuario)

    # 1) DDL: crear columna si no existe (en la BD de usuarios)
    print(">> Asegurando columna delegacion_id en", Usuario.__tablename__)
    with eng_usuarios.begin() as conn:
        conn.execute(
            text(f'ALTER TABLE {Usuario.__tablename__} ADD COLUMN IF NOT EXISTS delegacion_id INTEGER')
        )

    # 2) Cargar delegaciones {NOMBRE_NORMALIZADO: id} (ORM usa su propio bind)
    print(">> Leyendo delegaciones desde", Delegacion.__tablename__)
    delegaciones = db.session.query(Delegacion.id, Delegacion.nombre).all()
    mapa = { ( (nombre or "").strip().upper() ): did for (did, nombre) in delegaciones }
    print(f">> Delegaciones cargadas: {len(mapa)}")

    # 3) Backfill en usuarios (ORM usa bind de Usuario automáticamente)
    print(">> Backfill delegacion_id en usuarios…")
    asignados, sin_match = 0, 0
    usuarios = db.session.query(Usuario).all()
    for u in usuarios:
        if getattr(u, "delegacion_id", None):
            continue
        nombre_txt = (getattr(u, "delegacion", "") or "").strip().upper()
        if not nombre_txt:
            continue
        did = mapa.get(nombre_txt)
        if did:
            u.delegacion_id = did
            asignados += 1
        else:
            sin_match += 1
    db.session.commit()
    print(f">> Backfill listo. Asignados: {asignados} | Sin coincidencia: {sin_match}")

    # 4) Índice para rendimiento (en la BD de usuarios)
    print(">> Creando índice (si no existe)…")
    with eng_usuarios.begin() as conn:
        conn.execute(
            text(f'CREATE INDEX IF NOT EXISTS idx_{Usuario.__tablename__}_delegacion_id ON {Usuario.__tablename__}(delegacion_id)')
        )

    print("✅ Terminado.")
