from flask import Blueprint, render_template, request, redirect, url_for
from models import db, Notificacion
from flask_login import login_required, current_user

notificacion_bp = Blueprint('notificacion_bp', __name__)

@notificacion_bp.route('/notificaciones')
@login_required
def ver_notificaciones():
    if current_user.rol != 'admin':
        return "No autorizado", 403

    # Marcar todas como leídas al entrar
    notis_no_leidas = Notificacion.query.filter_by(leida=False).all()
    for noti in notis_no_leidas:
        noti.leida = True
    db.session.commit()

    notificaciones = Notificacion.query.order_by(Notificacion.fecha.desc()).all()
    return render_template('notificaciones.html', notificaciones=notificaciones)

@notificacion_bp.route('/notificaciones/eliminar/<int:id>', methods=['POST'])
@login_required
def eliminar_notificacion(id):
    noti = Notificacion.query.get_or_404(id)
    db.session.delete(noti)
    db.session.commit()
    return redirect(url_for('notificacion_bp.ver_notificaciones'))

@notificacion_bp.route('/notificaciones/eliminar_todas', methods=['POST'])
@login_required
def eliminar_todas_notificaciones():
    db.session.query(Notificacion).delete()
    db.session.commit()
    return redirect(url_for('notificacion_bp.ver_notificaciones'))
