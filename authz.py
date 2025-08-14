from functools import wraps
from flask_login import current_user, login_required
from flask import abort

def roles_required(*roles):
    def deco(f):
        @wraps(f)
        @login_required
        def wrapper(*args, **kwargs):
            if getattr(current_user, "rol", None) not in roles:
                abort(403)
            return f(*args, **kwargs)
        return wrapper
    return deco
