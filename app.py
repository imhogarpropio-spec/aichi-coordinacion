from flask import Flask
from config import Config
from extensiones import db, mail, login_manager
from werkzeug.security import generate_password_hash


def ensure_admin(app):
    """Crea o promueve un admin por defecto si no existe ninguno."""
    with app.app_context():
        from models import Usuario  # usa el bind 'usuarios' del modelo

        # ¿Ya hay administradores?
        if Usuario.query.filter_by(rol='admin').count() > 0:
            return

        # Valores desde Config (o fallback)
        correo     = app.config.get('DEFAULT_ADMIN_CORREO',     'admin@sistema.local')
        contrasena = app.config.get('DEFAULT_ADMIN_PASSWORD',   'Cambiar123!')
        genero     = app.config.get('DEFAULT_ADMIN_GENERO',     'H')
        deleg      = app.config.get('DEFAULT_ADMIN_DELEGACION', 'SISTEMA')

        existente = Usuario.query.filter_by(correo=correo).first()
        if existente:
            # Promover a admin si ya existe
            existente.rol = 'admin'
            existente.genero = existente.genero or genero
            existente.delegacion = existente.delegacion or deleg
            if not existente.contraseña:
                existente.contraseña = generate_password_hash(contrasena)
            db.session.commit()
            return

        # Crear admin por defecto
        nuevo = Usuario(
            nombre='Admin',
            genero=genero,
            correo=correo,
            delegacion=deleg,
            contraseña=generate_password_hash(contrasena),
            rol='admin'
        )
        db.session.add(nuevo)
        db.session.commit()

def create_app():
    app = Flask(__name__)
    app.config.from_object(Config)

    # Inicializar extensiones
    db.init_app(app)
    mail.init_app(app)
    login_manager.init_app(app)
    login_manager.login_view = 'auth_bp.login'

    # Registrar blueprints
    from routes.auth_routes import auth_bp
    from routes.dashboard_routes import dashboard_bp
    from routes.personal_routes import personal_bp
    from routes.delegaciones_routes import delegaciones_bp
    from routes.accesos_routes import accesos_bp
    from routes.usuarios_routes import usuarios_bp
    from routes.notificacion_routes import notificacion_bp


    app.register_blueprint(auth_bp)
    app.register_blueprint(dashboard_bp)
    app.register_blueprint(personal_bp)
    app.register_blueprint(delegaciones_bp)
    app.register_blueprint(accesos_bp)
    app.register_blueprint(usuarios_bp)
    app.register_blueprint(notificacion_bp)

    try:
        ensure_admin(app)
    except Exception as e:
        app.logger.warning(f"No se pudo asegurar administrador por defecto: {e}")

    return app
# ⬇️ Solo se ejecuta si corres el archivo directamente
if __name__ == '__main__':
    import os
    app = create_app()
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=True, host='0.0.0.0', port=port)

