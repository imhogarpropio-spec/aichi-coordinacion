from flask import Blueprint, render_template, request, redirect, url_for, flash, session, send_from_directory, jsonify
from flask_login import login_required, current_user
from datetime import datetime, timedelta
from models import db, Plantel, Personal, Usuario, HistorialCambios, Delegacion, ObservacionPersonal as Observacion
from utils import registrar_historial, registrar_notificacion
from sqlalchemy import distinct, func
import pandas as pd
import re
from sqlalchemy.exc import IntegrityError
from authz import roles_required


personal_bp = Blueprint('personal_bp', __name__)

@personal_bp.route("/busqueda_personal", methods=["GET", "POST"])
@login_required
def busqueda_personal():
    resultados = []
    if request.method == 'POST':
        filtros = {
            'apellido_paterno': request.form.get('apellido_paterno', '').strip(),
            'apellido_materno': request.form.get('apellido_materno', '').strip(),
            'nombre': request.form.get('nombre', '').strip(),
            'curp': request.form.get('curp', '').strip(),
            'rfc': request.form.get('rfc', '').strip()
        }

        consulta = Personal.query
        for campo, valor in filtros.items():
            if valor:
                consulta = consulta.filter(getattr(Personal, campo).ilike(f"%{valor}%"))

        resultados = consulta.all()

    return render_template("busqueda_personal.html", resultados=resultados)

@personal_bp.route("/detalle_personal/<int:id>")
@login_required
def vista_detalle_personal(id):
    persona = Personal.query.get_or_404(id)
    usuarios = Usuario.query.all()
    usuarios_por_id = {u.id: u for u in usuarios}
    niveles_disponibles = [n[0] for n in db.session.query(distinct(Plantel.nivel)).order_by(Plantel.nivel).all()]
    return render_template(
        "detalle_personal.html",
        persona=persona,
        niveles_disponibles=niveles_disponibles,
        usuarios_por_id=usuarios_por_id,
        timedelta=timedelta
    )

@personal_bp.route("/eliminar_personal/<int:id>", methods=["POST"])
@roles_required('admin')
def eliminar_personal(id):
    persona = Personal.query.get_or_404(id)

    # Guarda todo lo necesario ANTES de borrar (evita expire_on_commit)
    cct = persona.cct
    delegacion = persona.plantel.delegacion.nombre if persona.plantel and persona.plantel.delegacion else ''
    nombre_completo = f"{(persona.apellido_paterno or '').strip()} {(persona.apellido_materno or '').strip()} {(persona.nombre or '').strip()}".strip()
    curp = (persona.curp or "SIN CURP").strip()

    try:
        # 1) Borra dependientes primero (observaciones)
        Observacion.query.filter_by(personal_id=id).delete(synchronize_session=False)

        # 2) Borra el personal
        db.session.delete(persona)
        db.session.commit()

    except IntegrityError as e:
        db.session.rollback()
        flash("No se pudo eliminar: hay registros relacionados que lo impiden.", "danger")
        return redirect(url_for('personal_bp.vista_detalle_personal', id=id))
    except Exception as e:
        db.session.rollback()
        flash(f"Error al eliminar: {e}", "danger")
        return redirect(url_for('personal_bp.vista_detalle_personal', id=id))

    # Notificación (usa los datos preservados)
    registrar_notificacion(
        f"{current_user.nombre} eliminó a {nombre_completo} ({curp}) de {cct}",
        tipo="personal"
    )

    flash("El personal ha sido eliminado correctamente.", "success")
    return redirect(url_for('personal_bp.vista_personal', delegacion=delegacion, cct=cct))


@personal_bp.route("/editar_personal/<int:id>", methods=["POST"])
@login_required
def editar_personal(id):
    persona = Personal.query.get_or_404(id)
    nuevo_curp = request.form.get("curp", "").strip().upper()
    nuevo_rfc = request.form.get("rfc", "").strip().upper()
    otro_con_curp = Personal.query.filter(Personal.curp == nuevo_curp, Personal.id != id).first()
    if otro_con_curp:
        flash(f"❌ CURP duplicado: ya pertenece a {otro_con_curp.nombre}", "danger")
        return redirect(url_for("personal_bp.vista_detalle_personal", id=id))
    otro_con_rfc = Personal.query.filter(Personal.rfc == nuevo_rfc, Personal.id != id).first()
    if otro_con_rfc:
        flash(f"❌ RFC duplicado: ya pertenece a {otro_con_rfc.nombre}", "danger")
        return redirect(url_for("personal_bp.vista_detalle_personal", id=id))

    campos = ["apellido_paterno", "apellido_materno", "nombre", "genero", "rfc", "curp",
              "clave_presupuestal", "funcion", "grado_estudios", "titulado", "fecha_ingreso",
              "fecha_baja_jubilacion", "estatus_membresia", "nombramiento", "domicilio", "numero",
              "localidad", "colonia", "municipio", "cp", "tel1", "tel2", "correo_electronico"]

    for campo in campos:
        valor_anterior = getattr(persona, campo)
        valor_nuevo = request.form.get(campo)
        if campo in ["curp"]: valor_nuevo = nuevo_curp
        if campo in ["rfc"]: valor_nuevo = nuevo_rfc
        if "fecha" in campo and valor_nuevo == "": valor_nuevo = None
        if "fecha" in campo and valor_nuevo:
            try:
                valor_nuevo = datetime.strptime(valor_nuevo, "%Y-%m-%d").date()
            except ValueError:
                flash("Formato de fecha inválido.", "danger")
                return redirect(url_for("personal_bp.vista_detalle_personal", id=persona.id))
        if str(valor_anterior) != str(valor_nuevo):
            registrar_historial("personal", campo, valor_anterior, valor_nuevo, persona.id, current_user.nombre, "edición")
            setattr(persona, campo, valor_nuevo)

    db.session.commit()
    registrar_notificacion(
        f"{current_user.nombre} actualizó datos de {persona.nombre} ({getattr(persona, 'curp', 'SIN CURP')})",
        tipo="personal"
    )
    flash("Cambios guardados correctamente.", "success")
    return redirect(url_for("personal_bp.vista_detalle_personal", id=persona.id))

@personal_bp.route('/subir_excel_personal/<cct>', methods=['POST'])
@roles_required('admin')
def subir_excel_personal(cct):
    plantel = Plantel.query.filter_by(cct=cct).first_or_404()

    def convertir_fecha(valor):
        try:
            return pd.to_datetime(valor).date() if pd.notna(valor) else None
        except Exception:
            return None

    def limpiar_entero(valor):
        try:
            if pd.isna(valor) or str(valor).strip() == '':
                return None
            return int(valor)
        except:
            return None

    def limpiar_str(valor, max_len=None):
        if pd.isna(valor):
            return ''
        valor = str(valor).strip()
        return valor[:max_len] if max_len else valor

    if 'archivo_excel' not in request.files:
        flash('No se envió ningún archivo.', 'danger')
        return redirect(url_for('personal_bp.vista_personal', delegacion=plantel.delegacion, cct=cct))

    archivo = request.files['archivo_excel']

    if archivo.filename == '':
        flash('Nombre de archivo vacío.', 'danger')
        return redirect(url_for('personal_bp.vista_personal', delegacion=plantel.delegacion, cct=cct))

    if archivo and archivo.filename.endswith('.xlsx'):
        try:
            df = pd.read_excel(archivo)
            registros_agregados = 0
            registros_ignorados = 0

            for _, row in df.iterrows():
                if not str(row.get('nombre', '')).strip():
                    registros_ignorados += 1
                    continue

                nuevo = Personal(
                    cct=cct,
                    apellido_paterno=limpiar_str(row.get('apellido_paterno')),
                    apellido_materno=limpiar_str(row.get('apellido_materno')),
                    nombre=limpiar_str(row.get('nombre')),
                    genero=limpiar_str(row.get('genero')),
                    rfc=limpiar_str(row.get('rfc')),
                    curp=limpiar_str(row.get('curp')),
                    clave_presupuestal=limpiar_str(row.get('clave_presupuestal')),
                    funcion=limpiar_str(row.get('funcion')),
                    grado_estudios=limpiar_str(row.get('grado_maximo_estudios')),
                    titulado=limpiar_str(row.get('titulado')),
                    fecha_ingreso=convertir_fecha(row.get('fecha_ingreso')),
                    fecha_baja_jubilacion=convertir_fecha(row.get('fecha_baja_jubilacion')),
                    estatus_membresia=limpiar_str(row.get('estatus_membresia')),
                    nombramiento=limpiar_str(row.get('nombramiento')),
                    domicilio=limpiar_str(row.get('domicilio')),
                    numero=limpiar_entero(row.get('numero')),
                    localidad=limpiar_str(row.get('localidad')),
                    colonia=limpiar_str(row.get('colonia')),
                    municipio=limpiar_str(row.get('municipio')),
                    cp=limpiar_entero(row.get('cp')),
                    tel1=limpiar_entero(row.get('tel1')),
                    tel2=limpiar_entero(row.get('tel2')),
                    correo_electronico=limpiar_str(row.get('correo_electronico'))
                )

                db.session.add(nuevo)
                registros_agregados += 1

            db.session.commit()
            registrar_notificacion(
                f"{current_user.nombre} importó {registros_agregados} personas (ignoradas: {registros_ignorados}) desde Excel en {cct}",
                tipo="personal"
            )
            flash(f'Se cargaron {registros_agregados} personas. {registros_ignorados} filas ignoradas.', 'success')

        except Exception as e:
            db.session.rollback()
            flash(f'❌ Error al procesar el archivo: {str(e)}', 'danger')
    else:
        flash('Formato de archivo no permitido. Usa archivos .xlsx', 'danger')

    return redirect(url_for('personal_bp.vista_personal', delegacion=plantel.delegacion, cct=cct))

@personal_bp.route("/historial/<entidad>/<int:id>")
@login_required
def ver_historial(entidad, id):
    historial = HistorialCambios.query.filter_by(entidad=entidad, entidad_id=id).order_by(HistorialCambios.fecha.desc()).all()
    return render_template("ver_historial.html", historial=historial, tipo=entidad)

@personal_bp.route('/agregar_observacion/<int:personal_id>', methods=['POST'])
@login_required
def agregar_observacion(personal_id):
    texto = (request.form.get('texto') or '').strip()
    if not texto:
        flash("El texto no puede estar vacío.", "warning")
        return redirect(request.referrer or url_for('personal_bp.vista_detalle_personal', id=personal_id))

    nueva_obs = Observacion(personal_id=personal_id, usuario_id=current_user.id, texto=texto)
    db.session.add(nueva_obs)
    db.session.commit()

    persona = Personal.query.get(personal_id)
    if persona:
        nombre = f"{(persona.apellido_paterno or '').strip()} {(persona.apellido_materno or '').strip()} {(persona.nombre or '').strip()}".strip()
        curp = (persona.curp or 'SIN CURP').strip()
        msg = f"{current_user.nombre} agregó una observación a {nombre} ({curp})"
    else:
        msg = f"{current_user.nombre} agregó una observación a personal #{personal_id}"

    registrar_notificacion(msg, tipo="personal")

    flash("✅ Observación registrada correctamente.", "success")
    return redirect(request.referrer or url_for('personal_bp.vista_detalle_personal', id=personal_id))


from sqlalchemy import text

@personal_bp.route('/editar_observacion/<int:obs_id>', methods=['POST'])
@roles_required('admin')
def editar_observacion(obs_id):
    # 1) Obtener el personal_id SIN cargar el objeto a la sesión
    pid = db.session.query(Observacion.personal_id).filter_by(id=obs_id).scalar()
    if pid is None:
        flash("Observación no encontrada.", "warning")
        return redirect(request.referrer or url_for('personal_bp.vista_detalle_personal', id=0))

    # 2) Validar texto
    texto = (request.form.get('nuevo_texto') or request.form.get('texto') or '').strip()
    if not texto:
        flash("El texto no puede estar vacío.", "warning")
        return redirect(request.referrer or url_for('personal_bp.vista_detalle_personal', id=pid))

    try:
        # 3) 🔒 UPDATE dirigido con SQL crudo: SOLO 'texto'
        db.session.execute(
            text("UPDATE observacion_personal SET texto = :t WHERE id = :i"),
            {"t": texto, "i": obs_id}
        )
        db.session.commit()
    except Exception as e:
        db.session.rollback()
        flash(f"Error al actualizar observación: {e}", "danger")
        return redirect(request.referrer or url_for('personal_bp.vista_detalle_personal', id=pid))

    # 4) Notificación con nombre + CURP
    persona = Personal.query.get(pid)
    if persona:
        nombre = f"{(persona.apellido_paterno or '').strip()} {(persona.apellido_materno or '').strip()} {(persona.nombre or '').strip()}".strip()
        curp = (persona.curp or 'SIN CURP').strip()
        msg = f"{current_user.nombre} editó una observación de {nombre} ({curp})"
    else:
        msg = f"{current_user.nombre} editó una observación de personal #{pid}"

    registrar_notificacion(msg, tipo="personal")
    flash("✅ Observación actualizada correctamente.", "success")
    return redirect(request.referrer or url_for('personal_bp.vista_detalle_personal', id=pid))


@personal_bp.route('/eliminar_observacion/<int:obs_id>', methods=['POST'])
@roles_required('admin')
def eliminar_observacion(obs_id):
    obs = Observacion.query.get_or_404(obs_id)
    pid = obs.personal_id

    # Guarda datos para la noti antes de borrar
    persona = Personal.query.get(pid)
    if persona:
        nombre = f"{(persona.apellido_paterno or '').strip()} {(persona.apellido_materno or '').strip()} {(persona.nombre or '').strip()}".strip()
        curp = (persona.curp or 'SIN CURP').strip()
        msg = f"{current_user.nombre} eliminó una observación de {nombre} ({curp})"
    else:
        msg = f"{current_user.nombre} eliminó una observación de personal #{pid}"

    db.session.delete(obs)
    db.session.commit()

    registrar_notificacion(msg, tipo="personal")

    flash("Observación eliminada correctamente.", "success")
    return redirect(request.referrer or url_for('personal_bp.vista_detalle_personal', id=pid))



@personal_bp.route('/cambiar_adscripcion/<int:id>', methods=['POST'])
@roles_required('admin')
def cambiar_adscripcion(id):
    persona = Personal.query.get_or_404(id)
    nuevo_cct = (request.form.get("nuevo_cct") or "").strip()
    motivo = (request.form.get("motivo") or "").strip()

    # Validaciones básicas
    if not nuevo_cct:
        flash("⚠️ Debes seleccionar un CCT válido.", "warning")
        return redirect(url_for("personal_bp.vista_detalle_personal", id=id))

    if nuevo_cct == (persona.cct or "").strip():
        flash("ℹ️ El personal ya está adscrito a ese CCT.", "info")
        return redirect(url_for("personal_bp.vista_detalle_personal", id=id))

    if not motivo:
        flash("⚠️ Debes escribir un motivo del cambio.", "warning")
        return redirect(url_for("personal_bp.vista_detalle_personal", id=id))

    # Datos previos para auditoría
    cct_anterior = persona.cct
    estado_anterior = persona.estatus_membresia or ""

    # Resolver plantel destino (si aplica mostrar nombre en observación)
    nuevo_plantel = Plantel.query.filter_by(cct=nuevo_cct).first()
    nombre_plantel = nuevo_plantel.nombre if nuevo_plantel else ""

    # HISTORIAL: cambio de CCT (antes → después)
    registrar_historial(
        entidad="personal",
        campo="cct",
        valor_anterior=cct_anterior,
        valor_nuevo=nuevo_cct,
        entidad_id=persona.id,
        usuario=current_user.nombre,
        tipo="cambio de adscripción"
    )

    # OBSERVACIÓN: motivo del cambio
    observacion = Observacion(
        personal_id=persona.id,
        usuario_id=current_user.id,
        texto=f"Motivo de cambio de adscripción: {motivo} (ahora adscrito al CCT {nuevo_cct} – {nombre_plantel})",
        fecha=datetime.now()
    )
    db.session.add(observacion)

    # APLICAR cambio y LIMPIAR estado "baja en proceso"
    persona.cct = nuevo_cct
    persona.estatus_membresia = "ACTIVO"   # <- vuelve al color normal en la UI

    # HISTORIAL: dejar rastro de que el estatus se limpió/normalizó tras el movimiento
    registrar_historial(
        entidad="personal",
        campo="estatus_membresia",
        valor_anterior=estado_anterior,
        valor_nuevo="ACTIVO (tras cambio de adscripción)",
        entidad_id=persona.id,
        usuario=current_user.nombre,
        tipo="ADSCRIPCION_CAMBIO"
    )

    # GUARDAR todo junto
    db.session.commit()

    # NOTIFICACIÓN: incluye CCT anterior → nuevo y aclaración de estatus
    registrar_notificacion(
        f"{current_user.nombre} cambió la adscripción de {persona.apellido_paterno} {persona.apellido_materno} {persona.nombre} "
        f"({getattr(persona, 'curp', 'SIN CURP')}) de {cct_anterior} a {nuevo_cct}. "
        f"Estatus regresó a ACTIVO.",
        tipo="personal"
    )

    flash("✅ Adscripción actualizada, estado normalizado y observación registrada.", "success")
    return redirect(url_for("personal_bp.vista_detalle_personal", id=persona.id))


@personal_bp.route('/personal/<delegacion>/<cct>')
@login_required
def vista_personal(delegacion, cct):
    plantel = Plantel.query.filter_by(cct=cct).first_or_404()

    delegacion_obj = Delegacion.query.filter(func.lower(Delegacion.nombre) == delegacion.lower()).first()

    if not delegacion_obj:
        delegacion_obj = plantel.delegacion

    personal = Personal.query.filter_by(cct=cct).order_by(
        Personal.apellido_paterno,
        Personal.apellido_materno,
        Personal.nombre
    ).all()

    return render_template("consulta_personal.html", plantel=plantel, personal=personal, delegacion=delegacion_obj)

@personal_bp.route('/agregar_personal/<cct>', methods=['POST'])
@login_required
def agregar_personal_manual(cct):
    plantel = Plantel.query.filter_by(cct=cct).first_or_404()

    def validar_curp(curp):
        return re.fullmatch(r'^[A-Z]{4}\d{6}[HM][A-Z]{5}[0-9A-Z]\d$', curp)

    def validar_rfc(rfc):
        return re.fullmatch(r'^[A-ZÑ&]{3,4}\d{6}[A-Z0-9]{3}$', rfc)

    curp = request.form.get('curp', '').strip().upper()
    rfc = request.form.get('rfc', '').strip().upper()

    if not validar_curp(curp):
        flash("⚠️ CURP inválido. Verifica el formato (18 caracteres, en mayúsculas).", "danger")
        return redirect(url_for('personal_bp.vista_personal', delegacion=plantel.delegacion.nombre, cct=cct))

    if not validar_rfc(rfc):
        flash("⚠️ RFC inválido. Verifica el formato (13 caracteres, en mayúsculas).", "danger")
        return redirect(url_for('personal_bp.vista_personal', delegacion=plantel.delegacion.nombre, cct=cct))

    existe_curp = Personal.query.filter_by(curp=curp).first()
    existe_rfc = Personal.query.filter_by(rfc=rfc).first()

    if existe_curp:
        flash(f"❌ El CURP ya está registrado para: {existe_curp.nombre} {existe_curp.apellido_paterno} {existe_curp.apellido_materno} en el CCT {existe_curp.cct}.", "danger")
        return redirect(url_for('personal_bp.vista_personal', delegacion=plantel.delegacion.nombre, cct=cct))

    if existe_rfc:
        flash(f"❌ El RFC ya está registrado para: {existe_rfc.nombre} {existe_rfc.apellido_paterno} {existe_rfc.apellido_materno} en el CCT {existe_rfc.cct}.", "danger")
        return redirect(url_for('personal_bp.vista_personal', delegacion=plantel.delegacion.nombre, cct=cct))

    def convertir_fecha(campo):
        valor = request.form.get(campo)
        if valor:
            try:
                return datetime.strptime(valor, "%Y-%m-%d").date()
            except ValueError:
                flash(f"⚠️ Fecha inválida en campo: {campo}. Usa el selector de fecha.", "warning")
        return None

    nuevo = Personal(
        cct=cct,
        apellido_paterno=request.form.get('apellido_paterno', '').strip(),
        apellido_materno=request.form.get('apellido_materno', '').strip(),
        nombre=request.form.get('nombre', '').strip(),
        genero=request.form.get('genero', '').strip(),
        rfc=rfc,
        curp=curp,
        clave_presupuestal=request.form.get('clave_presupuestal', '').strip(),
        funcion=request.form.get('funcion', '').strip(),
        grado_estudios=request.form.get('grado_estudios', '').strip(),
        titulado=request.form.get('titulado', '').strip(),
        fecha_ingreso=convertir_fecha('fecha_ingreso'),
        fecha_baja_jubilacion=convertir_fecha('fecha_baja_jubilacion'),
        estatus_membresia=request.form.get('estatus_membresia', '').strip(),
        nombramiento=request.form.get('nombramiento', '').strip(),
        domicilio=request.form.get('domicilio', '').strip(),
        numero=request.form.get('numero', '').strip(),
        localidad=request.form.get('localidad', '').strip(),
        colonia=request.form.get('colonia', '').strip(),
        municipio=request.form.get('municipio', '').strip(),
        cp=request.form.get('cp', '').strip(),
        tel1=request.form.get('tel1', '').strip(),
        tel2=request.form.get('tel2', '').strip(),
        correo_electronico=request.form.get('correo_electronico', '').strip()
    )

    try:
        db.session.add(nuevo)
        db.session.commit()
        registrar_notificacion(
            f"{current_user.nombre} dio de alta a {nuevo.nombre} ({getattr(nuevo, 'curp', 'SIN CURP')}) en {cct}",
            tipo="personal"
        )
        flash("✅ Personal agregado correctamente.", "success")
    except Exception as e:
        db.session.rollback()
        flash("❌ Error al guardar el personal: " + str(e), "danger")

    return redirect(url_for('personal_bp.vista_personal', delegacion=plantel.delegacion.nombre, cct=cct))

@personal_bp.route('/ccts_por_nivel')
@login_required
def obtener_ccts_por_nivel():
    nivel = request.args.get('nivel')
    
    if not nivel:
        return jsonify([])

    # Búsqueda insensible a mayúsculas
    planteles = Plantel.query.filter(
        func.upper(Plantel.nivel) == nivel.upper()
    ).order_by(Plantel.nombre).all()

    resultado = [
        {"cct": p.cct, "nombre": p.nombre}
        for p in planteles
    ]

    return jsonify(resultado)


@personal_bp.route('/solicitar_baja/<int:id>', methods=['POST'])
@roles_required('admin', 'coordinador', 'delegado', 'secretario')
def solicitar_baja(id):
    persona = Personal.query.get_or_404(id)
    motivo = (request.form.get('motivo_baja') or '').strip()

    if not motivo:
        flash('Debes indicar el motivo de la baja.', 'warning')
        return redirect(url_for('personal_bp.vista_detalle_personal', id=id))

    # 1) Historial
    registrar_historial(
        entidad='personal',
        campo='solicitud_baja',
        valor_anterior='',
        valor_nuevo=f"Motivo: {motivo}",
        entidad_id=persona.id,
        usuario=getattr(current_user, 'nombre', None),
        tipo='BAJA'
    )

    # 2) Notificación
    nombre_completo = f"{persona.apellido_paterno} {persona.apellido_materno} {persona.nombre}".strip()
    cct_texto = getattr(persona, 'cct', None) or getattr(persona, 'plantel', None).cct

    descripcion_noti = (
        f"Solicitud de baja de personal: {nombre_completo} "
        f"(CCT {cct_texto}). Motivo: {motivo}. "
        f"Solicitado por: {getattr(current_user,'nombre','Usuario')}"
    )

    registrar_notificacion(descripcion_noti, tipo='personal')




    # 3) Marcar estado visible en UI
    persona.estatus_membresia = 'BAJA EN PROCESO'
    db.session.commit()

    flash('Solicitud de baja enviada y registrada. La ficha quedó en “Baja en proceso”.', 'success')

    # Redirige al listado del plantel (ajusta si usas otra ruta)
    try:
        return redirect(url_for('personal_bp.vista_personal',
                                delegacion=persona.plantel.delegacion.nombre,
                                cct=cct_texto))
    except Exception:
        return redirect(url_for('personal_bp.vista_detalle_personal', id=persona.id))
    

@personal_bp.route('/rechazar_baja/<int:id>', methods=['POST'])
@roles_required('admin', 'coordinador')
def rechazar_baja(id):
    persona = Personal.query.get_or_404(id)

    # Permite rechazar solo si realmente está en proceso
    if (persona.estatus_membresia or '').upper() != 'BAJA EN PROCESO':
        flash('Solo puedes rechazar bajas que estén en proceso.', 'warning')
        return redirect(url_for('personal_bp.vista_detalle_personal', id=id))

    motivo = (request.form.get('motivo_rechazo') or '').strip()
    if not motivo:
        flash('Debes indicar el motivo del rechazo.', 'warning')
        return redirect(url_for('personal_bp.vista_detalle_personal', id=id))

    estado_anterior = persona.estatus_membresia or ''

    # 1) Historial
    registrar_historial(
        entidad='personal',
        campo='rechazo_baja',
        valor_anterior=estado_anterior,
        valor_nuevo=f"Rechazada. Motivo: {motivo}",
        entidad_id=persona.id,
        usuario=getattr(current_user, 'nombre', None),
        tipo='BAJA_RECHAZADA'
    )

    # 2) Notificación
    nombre_completo = f"{persona.apellido_paterno} {persona.apellido_materno} {persona.nombre}".strip()
    cct_texto = getattr(persona, 'cct', None) or getattr(persona, 'plantel', None).cct
    descripcion_noti = (
        f"Rechazo de solicitud de baja para {nombre_completo} (CCT {cct_texto}). "
        f"Motivo: {motivo}. Por: {getattr(current_user,'nombre','Usuario')}."
    )
    registrar_notificacion(descripcion_noti, tipo='personal')

    # 3) Revertir estatus -> la tarjeta vuelve a color normal
    persona.estatus_membresia = 'ACTIVO'
    db.session.commit()

    flash('Solicitud de baja rechazada. Estatus regresó a ACTIVO.', 'success')

    try:
        return redirect(url_for('personal_bp.vista_personal',
                                delegacion=persona.plantel.delegacion.nombre,
                                cct=cct_texto))
    except Exception:
        return redirect(url_for('personal_bp.vista_detalle_personal', id=persona.id))
