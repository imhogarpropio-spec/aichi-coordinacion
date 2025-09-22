# routes/planteles_api.py
from flask import Blueprint, jsonify, abort
from flask_login import login_required
from models import Plantel, Delegacion

planteles_api = Blueprint("planteles_api", __name__, url_prefix="/api/planteles")

@planteles_api.get("/<string:cct>")
@login_required
def get_plantel_por_cct(cct):
    p = Plantel.query.filter_by(cct=cct).first()
    if not p:
        abort(404, description="CCT no encontrado")

    delega = p.delegacion.nombre if p.delegacion else None

    return jsonify({
        # claves básicas
        "cct": p.cct,
        "plantel_nombre": p.nombre,
        "turno": p.turno,
        "nivel": p.nivel,
        "subs_modalidad": getattr(p, "modalidad", None),  # tu modelo usa "modalidad"
        "zona_escolar": p.zona_escolar,
        "sector": p.sector,

        # domicilio del plantel
        "dom_esc_calle": p.calle,
        "dom_esc_num_ext": p.num_exterior,
        "dom_esc_num_int": p.num_interior,
        "dom_esc_cruce1": p.cruce_1,
        "dom_esc_cruce2": p.cruce_2,
        "dom_esc_localidad": p.localidad,
        "dom_esc_colonia": p.colonia,
        "dom_esc_mun_nom": p.municipio,
        "dom_esc_cp": p.cp,
        "dom_esc_coordenadas_gps": p.coordenadas_gps,

        # metadatos varios
        "estado": p.estado,
        "delegacion_id": p.delegacion_id,
        "delegacion_nombre": delega,
    })
