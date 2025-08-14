from flask import Blueprint, render_template, redirect, url_for, flash
from flask_login import login_required, current_user
from models import Acceso

accesos_bp = Blueprint('accesos_bp', __name__)

@accesos_bp.route('/registro_accesos')
@login_required
def registro_accesos():
    if current_user.rol != 'admin':
        flash("Acceso no autorizado", "danger")
        return redirect(url_for('dashboard_bp.dashboard'))

    accesos = Acceso.query.order_by(Acceso.fecha_entrada.desc()).all()
    activos = Acceso.query.filter_by(fecha_salida=None).count()

    return render_template("registro_accesos.html", accesos=accesos, activos=activos)