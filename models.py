from extensiones import db
from flask_login import UserMixin
from datetime import datetime
from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from datetime import datetime
from pytz import timezone
from werkzeug.security import check_password_hash, generate_password_hash
from sqlalchemy.orm import foreign





class Usuario(UserMixin, db.Model):
    __bind_key__ = 'usuarios'
    __tablename__ = 'usuarios'


    id = db.Column(db.Integer, primary_key=True)
    nombre = db.Column(db.String(100), nullable=False)
    genero = db.Column(db.String(1), nullable=False)  # 'H' o 'M'
    correo = db.Column(db.String(100), unique=True, nullable=False)
    delegacion = db.Column(db.String(50), nullable=False)
    contraseña = db.Column(db.Text, nullable=False)
    rol = db.Column(db.String(20), nullable=False, default='delegado')
    zona = db.Column(db.String(20))
    intentos_fallidos = db.Column(db.Integer, default=0)
    bloqueado_hasta = db.Column(db.DateTime, nullable=True)
    delegacion_id = db.Column(db.Integer, nullable=True)  # <-- NUEVA COLUMNA

    @property
    def password(self):
        return self.contraseña

    @password.setter
    def password(self, value):
        self.contraseña = value

    # Helpers opcionales (si tu login usa estos helpers, mejor aún)
    def set_password(self, raw_password: str):
        self.contraseña = generate_password_hash(raw_password)

    def check_password(self, raw_password: str) -> bool:
        try:
            return check_password_hash(self.contraseña, raw_password)
        except Exception:
            return False

class Plantel(db.Model):

    id = db.Column(db.Integer, primary_key=True)
    cct = db.Column(db.String(15), unique=True, nullable=False)
    nombre = db.Column(db.String(100), nullable=False)
    turno = db.Column(db.String(20), nullable=False)
    nivel = db.Column(db.String(50), nullable=False)
    modalidad = db.Column(db.String(50), nullable=False)
    zona_escolar = db.Column(db.String(20), nullable=False)
    sector = db.Column(db.String(20), nullable=False)
    calle = db.Column(db.String(100))
    num_exterior = db.Column(db.String(10))
    num_interior = db.Column(db.String(10))
    cruce_1 = db.Column(db.String(100))
    cruce_2 = db.Column(db.String(100))
    localidad = db.Column(db.String(100))
    colonia = db.Column(db.String(100))
    municipio = db.Column(db.String(100))
    cp = db.Column(db.String(10))
    coordenadas_gps = db.Column(db.String(100))
    estado = db.Column(db.String(50), default='HIDALGO')

    delegacion_id = db.Column(db.Integer, db.ForeignKey('delegacion.id'), nullable=False)
    delegacion = db.relationship('Delegacion', back_populates='planteles')

    personal = db.relationship(
        'Personal',
        backref='plantel',                     # crea el atributo p.plantel
        cascade='all, delete-orphan',
        lazy='dynamic',
        primaryjoin=lambda: foreign(Personal.cct) == Plantel.cct
    )


class Personal(db.Model):
    __tablename__ = 'personal'
    __table_args__ = (db.UniqueConstraint('curp', 'clave_presupuestal', name='uq_curp_clave'),)

    id = db.Column(db.Integer, primary_key=True)

    # --- Campos base ---
    apellido_paterno = db.Column(db.String(100), nullable=False)
    apellido_materno = db.Column(db.String(100), nullable=False)
    nombre = db.Column(db.String(100), nullable=False)
    genero = db.Column(db.String(1), nullable=False)
    rfc = db.Column(db.String(13), nullable=False)
    curp = db.Column(db.String(18), nullable=False)
    clave_presupuestal = db.Column(db.String(100))
    funcion = db.Column(db.String(100))
    funcion_coordinacion = db.Column(db.String(150), nullable=True)
    grado_estudios = db.Column(db.String(100))
    titulado = db.Column(db.String(20))
    fecha_ingreso = db.Column(db.Date)
    fecha_baja_jubilacion = db.Column(db.Date)
    estatus_membresia = db.Column(db.String(50))
    nombramiento = db.Column(db.String(100))
    domicilio = db.Column(db.String(200))
    numero = db.Column(db.String(10))
    localidad = db.Column(db.String(100))
    colonia = db.Column(db.String(100))
    municipio = db.Column(db.String(100))
    cp = db.Column(db.String(10))
    tel1 = db.Column(db.String(20))
    tel2 = db.Column(db.String(20))
    correo_electronico = db.Column(db.String(100))

    # --- Campos v2 ---
    num = db.Column(db.Integer)
    dp_num_int = db.Column(db.Text)
    dp_cruce1 = db.Column(db.Text)
    dp_cruce2 = db.Column(db.Text)

    escuela_nombre = db.Column(db.Text)
    turno = db.Column(db.Text)
    nivel = db.Column(db.Text)
    subs_modalidad = db.Column(db.Text)
    zona_escolar = db.Column(db.Text)
    sector = db.Column(db.Text)

    dom_esc_calle = db.Column(db.Text)
    dom_esc_num_ext = db.Column(db.Text)
    dom_esc_num_int = db.Column(db.Text)
    dom_esc_cruce1 = db.Column(db.Text)
    dom_esc_cruce2 = db.Column(db.Text)
    dom_esc_localidad = db.Column(db.Text)
    dom_esc_colonia = db.Column(db.Text)
    dom_esc_mun_nom = db.Column(db.Text)
    dom_esc_cp = db.Column(db.Text)
    dom_esc_coordenadas_gps = db.Column(db.Text)

    estado = db.Column(db.Text)
    seccion_snte = db.Column(db.Text)
    del_o_ct = db.Column(db.Text)
    org = db.Column(db.Text)
    coord_reg = db.Column(db.Text)
    fun_sin = db.Column(db.Text)

    # FK
    cct = db.Column(db.String(15), db.ForeignKey('plantel.cct'))

    # Relaciones
    observaciones = db.relationship('ObservacionPersonal', backref='persona', lazy='dynamic')
    # OJO: NO declares aquí `plantel = db.relationship(...)` si en Plantel ya tienes:
    # personal = db.relationship('Personal', backref='plantel', ...)


class Acceso(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    usuario_id = db.Column(db.Integer)
    correo = db.Column(db.String(100))
    nombre = db.Column(db.String(100))
    rol = db.Column(db.String(20))
    fecha_entrada = db.Column(db.DateTime, default=datetime.utcnow)
    fecha_salida = db.Column(db.DateTime, nullable=True)

class Delegacion(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nombre = db.Column(db.String(100), unique=True, nullable=False)
    nivel = db.Column(db.String(50), nullable=False)
    delegado = db.Column(db.String(250))  # <-- campo nuevo
    planteles = db.relationship('Plantel', back_populates='delegacion', cascade='all, delete-orphan')

class HistorialCambios(db.Model):
    __tablename__ = 'historial_cambios'

    id = db.Column(db.Integer, primary_key=True)
    entidad = db.Column(db.String(50), nullable=False)
    entidad_id = db.Column(db.Integer, nullable=False)
    campo = db.Column(db.String(50), nullable=False)
    valor_anterior = db.Column(db.String(255))
    valor_nuevo = db.Column(db.String(255))
    fecha = db.Column(db.DateTime, default=lambda: datetime.now(timezone('America/Mexico_City')))
    usuario = db.Column(db.String(100))
    tipo = db.Column(db.String(50))  # <- esta línea es la que falta


class ObservacionPersonal(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    personal_id = db.Column(db.Integer, db.ForeignKey('personal.id'), nullable=False)
    usuario_id = db.Column(db.Integer, nullable=False)  # sin foreign key porque está en otra BD
    texto = db.Column(db.Text, nullable=False)
    fecha = db.Column(db.DateTime, default=lambda: datetime.now(timezone('America/Mexico_City')))


class Notificacion(db.Model):
    __tablename__ = 'notificaciones'

    id = db.Column(db.Integer, primary_key=True)
    usuario = db.Column(db.String(100), nullable=False)
    fecha = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    descripcion = db.Column(db.Text, nullable=False)
    tipo = db.Column(db.String(50), nullable=True)  # ejemplo: 'cct', 'delegacion', 'personal'
    leida = db.Column(db.Boolean, default=False)

