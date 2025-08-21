from functools import wraps
from flask import abort
from flask_login import current_user, login_required

def has_role(*roles: str) -> bool:
    """
    True si el usuario autenticado tiene alguno de los roles dados.
    Normaliza a minúsculas para evitar errores de casing.
    """
    if not getattr(current_user, "is_authenticated", False):
        return False
    user_role = getattr(current_user, "rol", None)
    if user_role is None:
        return False
    lr = str(user_role).lower()
    valid = {str(r).lower() for r in roles}
    return lr in valid

def roles_required(*roles: str):
    """
    Decorator: requiere login y que el rol del usuario esté en 'roles'.
    Si no está autenticado, redirige a login (por login_required).
    Si no tiene rol permitido, aborta 403.
    """
    def deco(f):
        @wraps(f)
        @login_required
        def wrapper(*args, **kwargs):
            if not has_role(*roles):
                abort(403)
            return f(*args, **kwargs)
        return wrapper
    return deco
