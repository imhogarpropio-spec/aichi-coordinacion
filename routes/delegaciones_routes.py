from flask import Blueprint, render_template, request, redirect, url_for, flash, send_from_directory
from flask_login import login_required, current_user
from utils import registrar_notificacion
from models import db, Delegacion, Plantel
import pandas as pd
from sqlalchemy.exc import IntegrityError
from authz import roles_required



delegaciones_bp = Blueprint('delegaciones_bp', __name__)

@delegaciones_bp.route('/delegaciones', methods=['GET'])
@login_required
def vista_delegaciones():
    # Delegado: solo su propia delegación
    if current_user.rol == 'delegado':
        delegaciones = Delegacion.query.filter(
            Delegacion.id == current_user.delegacion_id
        ).all()
    else:
        delegaciones = Delegacion.query.order_by(Delegacion.nivel, Delegacion.nombre).all()

    # Agrupar por nivel (como ya lo tenías)
    niveles = {}
    for d in delegaciones:
        niveles.setdefault(d.nivel, []).append(d)

    niveles_disponibles = [
        'PREESCOLAR GENERAL', 'PRIMARIA GENERAL', 'SECUNDARIA GENERAL',
        'SECUNDARIAS TÉCNICAS', 'TELESECUNDARIAS', 'NIVELES ESPECIALES',
        'MEDIA SUPERIOR', 'JUBILADOS'
    ]
    return render_template('consulta_delegaciones.html',
                           niveles=niveles,
                           niveles_disponibles=niveles_disponibles)

@delegaciones_bp.route('/eliminar_delegacion/<int:id>', methods=['POST'])
@roles_required('admin', 'coordinador')
def eliminar_delegacion(id):
    deleg = Delegacion.query.get_or_404(id)
    nombre = deleg.nombre
    nivel = deleg.nivel

    try:
        db.session.delete(deleg)
        db.session.commit()

        # ✅ Notificación correcta (delegación, no cct)
        registrar_notificacion(
            f"{current_user.nombre} eliminó la delegación '{nombre}' (nivel {nivel})",
            tipo="delegacion"
        )

        flash(f"Delegación '{nombre}' eliminada.", "success")

    except IntegrityError:
        db.session.rollback()
        flash(f"No se puede eliminar '{nombre}' porque tiene planteles asociados.", "warning")
        # Si quieres permitir borrado en cascada, hay que configurar el modelo con cascade / ondelete.

    except Exception as e:
        db.session.rollback()
        flash(f"Error al eliminar la delegación: {e}", "danger")

    return redirect(url_for('delegaciones_bp.vista_delegaciones'))


@delegaciones_bp.route('/editar_delegacion/<int:id>', methods=['POST'])
@roles_required('admin', 'coordinador')
def editar_delegacion(id):
    deleg = Delegacion.query.get_or_404(id)
    deleg.nombre = request.form['nuevo_nombre']
    deleg.nivel = request.form['nuevo_nivel']
    deleg.delegado = request.form['nuevo_delegado']
    db.session.commit()
    registrar_notificacion(
        f"{current_user.nombre} actualizó la delegación '{deleg.nombre}' (nivel {deleg.nivel})",
        tipo="delegacion"
    )
    flash("Delegación actualizada correctamente.", "success")
    return redirect(url_for('delegaciones_bp.vista_delegaciones'))
    flash('Delegación actualizada.', 'success')
    return redirect(url_for('delegaciones_bp.vista_delegaciones'))

@delegaciones_bp.route('/delegacion/<int:delegacion_id>/ccts', methods=['GET', 'POST'])
@login_required
def vista_ccts_por_delegacion(delegacion_id):
    delegacion = Delegacion.query.get_or_404(delegacion_id)

    # Delegado: solo su propia delegación
    if current_user.rol == 'delegado' and current_user.delegacion_id != delegacion.id:
        from flask import abort
        abort(403)

    if request.method == 'POST':
        # Solo admin/coordinador pueden crear CCTs
        if current_user.rol not in ('admin', 'coordinador'):
            from flask import abort
            abort(403)
        nuevo = Plantel(
            cct=request.form['cct'].strip().upper(),
            nombre=request.form.get('nombre').strip(),
            turno=request.form['turno'],
            nivel=request.form.get('nivel'),
            modalidad=request.form['modalidad'],
            zona_escolar=request.form['zona_escolar'],
            sector=request.form['sector'],
            calle=request.form.get('calle'),
            num_exterior=request.form.get('num_exterior'),
            num_interior=request.form.get('num_interior'),
            cruce_1=request.form.get('cruce_1'),
            cruce_2=request.form.get('cruce_2'),
            localidad=request.form.get('localidad'),
            colonia=request.form.get('colonia'),
            municipio=request.form.get('municipio'),
            cp=request.form.get('cp'),
            coordenadas_gps=request.form.get('coordenadas_gps'),
            estado='HIDALGO',
            delegacion_id=delegacion.id
        )
        db.session.add(nuevo)
        db.session.commit()
        flash('CCT registrado correctamente.', 'success')
        return redirect(url_for('delegaciones_bp.vista_ccts_por_delegacion', delegacion_id=delegacion.id))

    ccts = delegacion.planteles
    return render_template('consulta_ccts.html', delegacion=delegacion, ccts=ccts)

@delegaciones_bp.route('/eliminar_cct/<int:id>', methods=['POST'])
@roles_required('admin', 'coordinador')
def eliminar_cct(id):
    cct = Plantel.query.get_or_404(id)
    delegacion_id = cct.delegacion_id
    db.session.delete(cct)
    db.session.commit()
    registrar_notificacion(
        f"{current_user.nombre} eliminó el CCT {cct.cct} ({cct.nombre})",
        tipo="cct"
    )
    flash(f'CCT {cct.cct} eliminado.', 'warning')
    return redirect(url_for('delegaciones_bp.vista_ccts_por_delegacion', delegacion_id=delegacion_id))

@delegaciones_bp.route('/editar_cct/<int:id>', methods=['POST'])
@roles_required('admin', 'coordinador')
def editar_cct(id):
    cct = Plantel.query.get_or_404(id)
    cct.cct = request.form['nuevo_cct'].strip().upper()
    cct.nombre = request.form['nuevo_nombre'].strip()
    cct.turno = request.form['nuevo_turno'].strip()
    cct.nivel = request.form['nuevo_nivel'].strip()
    cct.modalidad = request.form['nuevo_modalidad'].strip()
    cct.zona_escolar = request.form['nuevo_zona_escolar'].strip()
    cct.sector = request.form['nuevo_sector'].strip()
    cct.calle = request.form.get('nuevo_calle', '').strip()
    cct.num_exterior = request.form.get('nuevo_num_exterior', '').strip()
    cct.num_interior = request.form.get('nuevo_num_interior', '').strip()
    cct.cruce_1 = request.form.get('nuevo_cruce_1', '').strip()
    cct.cruce_2 = request.form.get('nuevo_cruce_2', '').strip()
    cct.localidad = request.form.get('nuevo_localidad', '').strip()
    cct.colonia = request.form.get('nuevo_colonia', '').strip()
    cct.municipio = request.form.get('nuevo_municipio', '').strip()
    cct.cp = request.form.get('nuevo_cp', '').strip()
    cct.coordenadas_gps = request.form.get('nuevo_coordenadas_gps', '').strip()
    db.session.commit()
    flash('Plantel actualizado correctamente.', 'success')
    return redirect(url_for('delegaciones_bp.vista_ccts_por_delegacion', delegacion_id=cct.delegacion_id))

@delegaciones_bp.route('/descargar_plantilla_delegaciones')
@roles_required('admin')
def descargar_plantilla_delegaciones():
    return send_from_directory('static/plantillas', 'plantilla_delegaciones.xlsx', as_attachment=True)

@delegaciones_bp.route('/subir_excel/<int:delegacion_id>', methods=['POST'])
@roles_required('admin')
def subir_excel_delegaciones(delegacion_id):
    if 'archivo_excel' not in request.files:
        flash('No se envió ningún archivo.', 'danger')
        return redirect(url_for('delegaciones_bp.vista_delegaciones'))

    archivo = request.files['archivo_excel']

    if archivo.filename == '':
        flash('Nombre de archivo vacío.', 'danger')
        return redirect(url_for('delegaciones_bp.vista_delegaciones'))

    if archivo and archivo.filename.endswith('.xlsx'):
        try:
            df = pd.read_excel(archivo)
            registros_agregados = 0
            registros_ignorados = 0

            for index, row in df.iterrows():
                nombre = row.get('nombre')
                nivel = row.get('nivel')

                if isinstance(nombre, str) and isinstance(nivel, str) and nombre.strip() and nivel.strip():
                    nueva = Delegacion(nombre=nombre.strip(), nivel=nivel.strip().upper())
                    db.session.add(nueva)
                    registros_agregados += 1
                else:
                    registros_ignorados += 1

            db.session.commit()
            flash(f'Se cargaron {registros_agregados} delegaciones. {registros_ignorados} filas ignoradas.', 'success')
        except Exception as e:
            flash(f'Ocurrió un error al procesar el archivo: {str(e)}', 'danger')
    else:
        flash('Formato de archivo no permitido. Usa .xlsx', 'danger')

    return redirect(url_for('delegaciones_bp.vista_delegaciones'))

@delegaciones_bp.route('/delegacion/<int:delegacion_id>/subir_excel', methods=['POST'])
@roles_required('admin')
def subir_excel_ccts(delegacion_id):
    if 'archivo_excel' not in request.files:
        flash('No se envió ningún archivo.', 'danger')
        return redirect(url_for('delegaciones_bp.vista_ccts_por_delegacion', delegacion_id=delegacion_id))

    archivo = request.files['archivo_excel']

    if archivo.filename == '':
        flash('Nombre de archivo vacío.', 'danger')
        return redirect(url_for('delegaciones_bp.vista_ccts_por_delegacion', delegacion_id=delegacion_id))

    if archivo and archivo.filename.endswith('.xlsx'):
        try:
            df = pd.read_excel(archivo)
            registros_agregados = 0
            registros_ignorados = 0

            def limpiar(valor):
                return str(valor).strip() if pd.notna(valor) else ''

            for _, row in df.iterrows():
                cct = row.get('cct')
                nombre = row.get('nombre')
                turno = row.get('turno')
                nivel = row.get('nivel')
                modalidad = row.get('modalidad')

                if all(isinstance(v, str) and v.strip() for v in [cct, nombre, turno, nivel, modalidad]):
                    nuevo = Plantel(
                        cct=limpiar(cct).upper(),
                        nombre=limpiar(nombre),
                        turno=limpiar(turno),
                        nivel=limpiar(nivel),
                        modalidad=limpiar(modalidad),
                        zona_escolar=limpiar(row.get('zona_escolar')),
                        sector=limpiar(row.get('sector')),
                        calle=limpiar(row.get('calle')),
                        num_exterior=limpiar(row.get('num_exterior')),
                        num_interior=limpiar(row.get('num_interior')),
                        cruce_1=limpiar(row.get('cruce_1')),
                        cruce_2=limpiar(row.get('cruce_2')),
                        localidad=limpiar(row.get('localidad')),
                        colonia=limpiar(row.get('colonia')),
                        municipio=limpiar(row.get('municipio')),
                        cp=limpiar(row.get('cp')),
                        coordenadas_gps=limpiar(row.get('coordenadas_gps')),
                        estado='HIDALGO',
                        delegacion_id=delegacion_id
                    )
                    db.session.add(nuevo)
                    registros_agregados += 1
                else:
                    registros_ignorados += 1

            db.session.commit()
            flash(f'Se cargaron {registros_agregados} CCTs. {registros_ignorados} filas ignoradas.', 'success')
        except Exception as e:
            flash(f'Ocurrió un error al procesar el archivo: {str(e)}', 'danger')
    else:
        flash('Formato de archivo no permitido. Usa .xlsx', 'danger')

    return redirect(url_for('delegaciones_bp.vista_ccts_por_delegacion', delegacion_id=delegacion_id))

