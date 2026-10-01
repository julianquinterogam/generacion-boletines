"""Generación de todos los boletines del grado en un ZIP (un Excel por estudiante).

Estructura del ZIP:
    Boletines_<GRADO>_<PERIODO>_<AÑO>/
        Boletin_<GRADO>_<n>_<NOMBRE>.xlsx
"""
import io
import re
import zipfile

from boletin import generar_boletin_xlsx


def _limpiar(texto):
    return re.sub(r"_+", "_", re.sub(r"[^\w]+", "_", texto, flags=re.UNICODE)).strip("_")


def nombre_base(meta, estudiante, posicion):
    numero = estudiante.get("n")
    numero = int(numero) if isinstance(numero, (int, float)) else posicion
    return f"Boletin_{meta['grado']}_{numero:02d}_{_limpiar(estudiante['nombre'])}"


def generar_lote(meta, estudiantes, logros, progreso=None):
    """Devuelve (bytes_zip, advertencias). `estudiantes` ya debe traer promedio y puesto."""
    advertencias = []
    raiz = f"Boletines_{meta['grado']}_{meta['periodo']}_{meta['anio']}"
    total = len(estudiantes)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for i, e in enumerate(estudiantes, start=1):
            if progreso:
                progreso((i - 1) / total, f"Generando boletín {i} de {total}: {e['nombre']}")
            datos, adv = generar_boletin_xlsx(meta, e, logros)
            advertencias += [f"{e['nombre']}: {a}" for a in adv]
            z.writestr(f"{raiz}/{nombre_base(meta, e, i)}.xlsx", datos)
    if progreso:
        progreso(1.0, "Listo")
    return buf.getvalue(), advertencias
