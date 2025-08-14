from flask import Blueprint, render_template, request, redirect, url_for, flash
from flask_login import login_required, current_user
from models import db, Usuario, Delegacion
from werkzeug.security import generate_password_hash
from utils import registrar_notificacion
from authz import roles_required

usuarios_bp = Blueprint('usuarios_bp', __name__)

@usuarios_bp.route('/usuarios')
@roles_required('admin')
def listar_usuarios():
    usuarios = Usuario.query.all()
    return render_template('usuarios/listar.html', usuarios=usuarios)

ROLES_VALIDOS = {'admin', 'coordinador', 'secretario', 'delegado'}
ROLES_MAP = {
    'admin': 'admin',
    'coordinador': 'coordinador',
    'secretario': 'secretario',
    'delegado': 'delegado',
}

@usuarios_bp.route('/crear', methods=['GET','POST'])
@roles_required('admin')  # solo admin puede crear usuarios
def crear_usuario():
    # 👉 Para el GET (y para re-render si hay errores), cargamos delegaciones
    delegaciones = Delegacion.query.order_by(Delegacion.nivel, Delegacion.nombre).all()

    if request.method == 'POST':
        nombre      = (request.form.get('nombre') or '').strip()
        genero      = (request.form.get('genero') or '').strip()            # H/M como ya lo manejas
        correo      = (request.form.get('correo') or '').strip().lower()
        contrasena  = (request.form.get('contrasena') or '').strip()
        rol_in      = (request.form.get('rol') or '').strip().lower()
        rol         = ROLES_MAP.get(rol_in, 'delegado')   # normaliza
        deleg_id_in = (request.form.get('delegacion_id') or '').strip()

        # Validaciones básicas
        if not all([nombre, genero, correo, contrasena, deleg_id_in]):
            flash("Completa nombre, género, correo, contraseña y delegación.", "warning")
            return render_template('usuarios/crear.html', delegaciones=delegaciones)

        if rol not in ROLES_VALIDOS:
            flash(f"Rol inválido: {rol_in}", "danger")
            return render_template('usuarios/crear.html', delegaciones=delegaciones)

        if not deleg_id_in.isdigit():
            flash("Delegación inválida.", "danger")
            return render_template('usuarios/crear.html', delegaciones=delegaciones)

        deleg = Delegacion.query.get(int(deleg_id_in))
        if not deleg:
            flash("Delegación no encontrada.", "danger")
            return render_template('usuarios/crear.html', delegaciones=delegaciones)

        if Usuario.query.filter_by(correo=correo).first():
            flash("Ese correo ya está registrado.", "danger")
            return render_template('usuarios/crear.html', delegaciones=delegaciones)

        # Crear usuario
        nuevo = Usuario(
            nombre=nombre,
            genero=genero,  # 👈 tal cual lo recibes (H/M)
            correo=correo,
            delegacion_id=deleg.id,     # 👈 guarda FK
            delegacion=deleg.nombre,    # (opcional) por compatibilidad con vistas viejas
            contraseña=generate_password_hash(contrasena),
            rol=rol
        )
        db.session.add(nuevo)
        db.session.commit()

        flash("Usuario creado correctamente.", "success")
        return redirect(url_for('usuarios_bp.listar_usuarios'))

    # GET inicial
    return render_template('usuarios/crear.html', delegaciones=delegaciones)

@usuarios_bp.route('/usuarios/resetear/<int:id>', methods=['POST'])
@roles_required('admin')
def resetear_contrasena(id):
    usuario = Usuario.query.get_or_404(id)
    nueva = request.form.get('nueva')
    if nueva:
        usuario.contrasena = generate_password_hash(nueva)
        db.session.commit()
        registrar_notificacion(
            f"{current_user.nombre} reseteó la contraseña de '{usuario.nombre}'",
            tipo="usuario"
        )
        flash('Contraseña actualizada', 'success')
    return redirect(url_for('usuarios_bp.listar_usuarios'))

@usuarios_bp.route('/usuarios/eliminar_todos', methods=['POST'])
@roles_required('admin')
def eliminar_todos_los_usuarios():
    Usuario.query.delete()
    db.session.commit()
    registrar_notificacion(
        f"{current_user.nombre} eliminó TODOS los usuarios",
        tipo="usuario"
    )
    flash('Todos los usuarios han sido eliminados.', 'warning')
    return redirect(url_for('usuarios_bp.listar_usuarios'))

@usuarios_bp.route('/eliminar/<int:id>', methods=['POST'])
@roles_required('admin')
def eliminar_usuario(id):
    u = Usuario.query.get_or_404(id)

    # 🔒 No borrar al último administrador
    if u.rol == 'admin':
        admins = Usuario.query.filter_by(rol='admin').count()
        if admins <= 1:
            flash("No puedes eliminar al único usuario con rol Administrador.", "warning")
            return redirect(url_for('usuarios_bp.listar_usuarios'))

    # (opcional) No permitir que te borres a ti mismo si eres el único admin
    if u.id == current_user.id and u.rol == 'admin':
        admins = Usuario.query.filter_by(rol='admin').count()
        if admins <= 1:
            flash("No puedes eliminar tu cuenta: eres el único Administrador.", "warning")
            return redirect(url_for('usuarios_bp.listar_usuarios'))

    db.session.delete(u)
    db.session.commit()

    registrar_notificacion(
        f"{current_user.nombre} eliminó el usuario '{u.nombre}' ({u.correo})",
        tipo="usuario"
    )
    flash("Usuario eliminado.", "success")
    return redirect(url_for('usuarios_bp.listar_usuarios'))
