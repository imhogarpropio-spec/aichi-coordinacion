from flask import Blueprint, render_template, request, redirect, url_for, flash, session
from flask_login import login_user, logout_user, login_required, current_user
from werkzeug.security import check_password_hash
from extensiones import db, login_manager
from models import Usuario, Acceso
from datetime import datetime, timedelta
import pytz

auth_bp = Blueprint('auth_bp', __name__)

# Login
@auth_bp.route('/', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        correo = request.form['correo']
        contraseña = request.form['contraseña']
        usuario = Usuario.query.filter_by(correo=correo).first()
        ahora = datetime.now(pytz.timezone('America/Mexico_City'))

        if not usuario:
            flash('Correo o contraseña incorrectos', 'danger')
            return render_template('login.html')

        if usuario.bloqueado_hasta and ahora < usuario.bloqueado_hasta:
            minutos_restantes = int((usuario.bloqueado_hasta - ahora).total_seconds() // 60) + 1
            flash(f'Acceso bloqueado temporalmente. Intenta en {minutos_restantes} minuto(s).', 'danger')
            return render_template('login.html')

        if check_password_hash(usuario.contraseña, contraseña):
            usuario.intentos_fallidos = 0
            usuario.bloqueado_hasta = None
            db.session.commit()

            login_user(usuario)
            session.permanent = True

            nuevo_acceso = Acceso(
                usuario_id=usuario.id,
                correo=usuario.correo,
                nombre=usuario.nombre,
                rol=usuario.rol,
                fecha_entrada=ahora
            )
            db.session.add(nuevo_acceso)
            db.session.commit()
            session['acceso_id'] = nuevo_acceso.id

            return redirect(url_for('dashboard_bp.dashboard'))

        else:
            usuario.intentos_fallidos += 1
            if usuario.intentos_fallidos >= 5:
                usuario.bloqueado_hasta = ahora + timedelta(minutes=15)
                flash('❌ Demasiados intentos. Acceso bloqueado por 15 minutos.', 'danger')
            else:
                restantes = 5 - usuario.intentos_fallidos
                flash(f'Contraseña incorrecta. Intentos restantes: {restantes}', 'warning')

            db.session.commit()
            return render_template('login.html')

    return render_template('login.html')

# Logout
@auth_bp.route('/logout')
@login_required
def logout():
    if 'acceso_id' in session:
        acceso = Acceso.query.get(session['acceso_id'])
        if acceso and acceso.fecha_salida is None:
            acceso.fecha_salida = datetime.now(pytz.timezone('America/Mexico_City'))
            db.session.commit()
        session.pop('acceso_id', None)

    logout_user()

    if request.args.get('inactivo') == '1':
        flash('⚠️ La sesión se cerró automáticamente por inactividad.', 'warning')

    return redirect(url_for('auth_bp.login'))

# Carga de usuario (para Flask-Login)
@login_manager.user_loader
def load_user(user_id):
    return Usuario.query.get(int(user_id))