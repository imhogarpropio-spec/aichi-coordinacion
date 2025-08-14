from flask import Blueprint, render_template
from flask_login import login_required, current_user
from models import Notificacion

dashboard_bp = Blueprint('dashboard_bp', __name__)

@dashboard_bp.route('/dashboard')
@login_required
def dashboard():
    total_notificaciones = 0
    if current_user.rol == 'admin':
        total_notificaciones = Notificacion.query.filter_by(leida=False).count()

    return render_template(
        'dashboard.html',
        nombre=current_user.nombre,
        total_notificaciones=total_notificaciones
    )