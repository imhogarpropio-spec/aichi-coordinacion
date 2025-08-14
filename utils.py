from datetime import datetime
from pytz import timezone  # <-- IMPORTANTE (opción A)
from flask_login import current_user
from models import db, HistorialCambios, Notificacion

def actor_nombre():
    return current_user.nombre if getattr(current_user, "is_authenticated", False) else "Sistema"

def registrar_notificacion(descripcion, tipo=None):
    noti = Notificacion(
        usuario=actor_nombre(),
        fecha=datetime.utcnow(),  # guarda notis en UTC (estable)
        descripcion=descripcion,
        tipo=tipo
    )
    db.session.add(noti)
    db.session.commit()

def limpiar(valor):
    if valor is None:
        return ""
    return str(valor).strip()

def registrar_historial(entidad, campo, valor_anterior, valor_nuevo, entidad_id, usuario=None, tipo=None):
    va = limpiar(valor_anterior)
    vn = limpiar(valor_nuevo)

    if va == vn:
        print(f"[OMITIDO] {campo}: sin cambios ({va} == {vn})")
        return

    try:
        nuevo_registro = HistorialCambios(
            entidad=entidad,
            entidad_id=entidad_id,
            campo=campo,
            valor_anterior=va,
            valor_nuevo=vn,
            # ✅ CDMX; si prefieres, cambia por datetime.utcnow()
            fecha=datetime.now(timezone('America/Mexico_City')),
            usuario=usuario or actor_nombre(),
            tipo=tipo
        )
        db.session.add(nuevo_registro)
        db.session.commit()

        # Notificación automática por edición
        try:
            desc = f"{usuario or actor_nombre()} modificó {entidad} #{entidad_id}: {campo} — '{va}' → '{vn}'"
            registrar_notificacion(desc, tipo or entidad)
        except Exception as e:
            print(f"⚠️ No se pudo registrar notificación: {e}")

    except Exception as e:
        print(f"❌ Error al guardar historial: {e}")
