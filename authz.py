# authz.py
from functools import wraps
from flask import abort
from flask_login import current_user
from models import db, Plantel, Personal  # 👈 añade db

# ---- ROLES & PERMISSIONS ---------------------------------------------
# Secretario: visor global (sin editar). Delegado/auxiliar/lector: limitados a su delegación.
PERMS = {
    "admin": {"*"},  # todo

    "secretario": {  # visor global sin editar
        "users.view", "deleg.view", "plantel.view", "personal.view",
        "export.run", "audit.view",
    },

    "coordinador": {
        "deleg.view", "deleg.edit",
        "plantel.view", "plantel.edit",
        "personal.view", "personal.create", "personal.edit", "personal.delete",
        "export.run",
    },

    "delegado": {
        "deleg.view",
        "plantel.view",
        "personal.view", "personal.create", "personal.edit",
        "export.run",
    },

    "auxiliar_delegado": {
        "personal.view", "personal.create", "personal.edit",
    },

    "lector_delegacion": {
        "personal.view", "deleg.view", "plantel.view", "export.run",
    },
}

SCOPED_ROLES = {"delegado", "auxiliar_delegado", "lector_delegacion"}  # se filtran por delegacion_id


# ---- HELPERS ----------------------------------------------------------
def _role() -> str:
    return getattr(current_user, "rol", None) or "lector_delegacion"


def _caps_for(role: str):
    return PERMS.get(role, set())


def can(*caps: str) -> bool:
    """Usable en plantillas: can('personal.create')"""
    if not current_user.is_authenticated:
        return False
    role = _role()
    allowed = _caps_for(role)
    return ("*" in allowed) or all(c in allowed for c in caps)


def is_global_viewer() -> bool:
    """True para roles con vista global (secretario, admin)"""
    if not current_user.is_authenticated:
        return False
    return _role() in {"secretario", "admin"}


# ---- DECORADORES ------------------------------------------------------
def role_required(*roles):
    """Bloquea si el rol no está en la lista (admin siempre pasa)."""
    def deco(fn):
        @wraps(fn)
        def wrapper(*a, **kw):
            if not current_user.is_authenticated:
                abort(401)
            if _role() == "admin" or _role() in roles:
                return fn(*a, **kw)
            abort(403)
        return wrapper
    return deco

# Alias retrocompatible para imports existentes
def roles_required(*roles):
    return role_required(*roles)

def requires(*caps):
    """Bloquea si el rol no tiene TODAS las capacidades pedidas (admin siempre pasa)."""
    def deco(fn):
        @wraps(fn)
        def wrapper(*a, **kw):
            if not current_user.is_authenticated:
                abort(401)
            role = _role()
            allowed = _caps_for(role)
            if role == "admin" or "*" in allowed or all(c in allowed for c in caps):
                return fn(*a, **kw)
            abort(403)
        return wrapper
    return deco


# ---- SCOPE POR DELEGACIÓN (versión robusta) --------------------------
def limit_query_to_user_delegacion(query, model):
    """
    Para roles acotados (delegado/auxiliar/lector) agrega el filtro por delegación.
    Secretario/Admin ven global. Si el modelo no tiene delegacion_id, intenta
    resolver vía CCT -> Plantel.delegacion_id (caso Personal).
    Uso:
        q = limit_query_to_user_delegacion(Personal.query, Personal)
    """
    if (not current_user.is_authenticated) or is_global_viewer():
        return query

    user_del = getattr(current_user, "delegacion_id", None)
    if user_del is None:
        return query  # sin delegación asociada al usuario, no filtramos

    # Caso 1: el modelo tiene delegacion_id
    if hasattr(model, "delegacion_id"):
        return query.filter(model.delegacion_id == user_del)

    # Caso 2: Personal (o modelos con campo cct) -> join a Plantel
    if model is Personal or hasattr(model, "cct"):
        return query.join(Plantel, Plantel.cct == getattr(model, "cct")).filter(Plantel.delegacion_id == user_del)

    # Si no sabemos filtrar, devolvemos la query tal cual
    return query


def require_same_delegacion(record_getter):
    """
    Bloquea cambios cuando el registro NO pertenece a la delegación del usuario
    (para roles acotados). Secretario/Admin pasan.
    Uso:
        @requires("personal.edit")
        @require_same_delegacion(lambda pid: Personal.query.get_or_404(pid))
        def editar(pid): ...
    """
    def deco(fn):
        @wraps(fn)
        def wrapper(*a, **kw):
            if (not current_user.is_authenticated) or is_global_viewer():
                return fn(*a, **kw)

            obj = record_getter(*a, **kw)

            # Prioridad 1: campo directo
            obj_del = getattr(obj, "delegacion_id", None)

            # Prioridad 2: relación plantel
            if obj_del is None:
                plantel = getattr(obj, "plantel", None)
                if plantel is not None:
                    obj_del = getattr(plantel, "delegacion_id", None)

            # Prioridad 3: resolver por CCT (e.g., Personal.cct)
            if obj_del is None:
                cct = getattr(obj, "cct", None)
                if cct:
                    obj_del = db.session.query(Plantel.delegacion_id).filter_by(cct=cct).scalar()

            if obj_del is None or obj_del != getattr(current_user, "delegacion_id", None):
                abort(403)

            return fn(*a, **kw)
        return wrapper
    return deco

def has_role(*roles):
    """True si el usuario actual está autenticado y su rol está en roles."""
    return (
        getattr(current_user, "is_authenticated", False)
        and getattr(current_user, "rol", None) in roles
    )