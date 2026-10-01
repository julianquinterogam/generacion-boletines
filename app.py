import tempfile
from decimal import Decimal

import pandas as pd
import streamlit as st

from boletin import generar_boletin_xlsx
from lectura import MATERIAS, calcular_puestos, desempeno, fmt, leer_logros, leer_notas

st.set_page_config(page_title="Boletines – Revisión de datos", layout="wide")
st.title("Boletines – Paso 1: revisión de datos")
st.caption("Se leen los archivos sin modificarlos. Aún no se generan boletines.")

c1, c2 = st.columns(2)
f_xlsx = c1.file_uploader("Planilla de notas del grado (.xlsx)", type=["xlsx"])
f_docx = c2.file_uploader("Logros del grado (.docx)", type=["docx"])


def _guardar(archivo, sufijo):
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=sufijo)
    tmp.write(archivo.getvalue())
    tmp.close()
    return tmp.name


if f_xlsx and f_docx:
    meta, estudiantes, adv_x = leer_notas(_guardar(f_xlsx, ".xlsx"))
    logros, adv_w = leer_logros(_guardar(f_docx, ".docx"))

    st.subheader(f"Grado {meta['grado']} · Periodo {meta['periodo']} · {meta['anio']}")

    advertencias = adv_x + adv_w
    if advertencias:
        st.warning("Advertencias de lectura:\n\n" + "\n".join(f"- {a}" for a in advertencias))
    else:
        st.success("Lectura sin advertencias: todas las materias cruzan y no hay notas vacías ni fuera de rango.")

    # Puesto: sobre el promedio redondeado; empates comparten puesto, numeración consecutiva.
    calcular_puestos(estudiantes)

    st.markdown("#### Resumen del grado")
    st.dataframe(pd.DataFrame([{
        "N°": e["n"],
        "Estudiante": e["nombre"],
        "Promedio exacto": float(e["promedio_exacto"].quantize(Decimal("0.001"))) if e["promedio_exacto"] is not None else None,
        "Promedio (boletín)": fmt(e["promedio"]) if e["promedio"] is not None else "—",
        "Puesto": e["puesto"],
    } for e in estudiantes]), hide_index=True, width="stretch")

    st.markdown("#### Detalle por estudiante")
    for e in estudiantes:
        with st.expander(f"{e['n']}. {e['nombre']}  —  promedio {fmt(e['promedio'])}  ·  puesto {e['puesto']}"):
            grupo_previo = None
            for m in MATERIAS:
                if m["grupo"] and m["grupo"] != grupo_previo:
                    st.markdown(f"**{m['grupo']}**")
                grupo_previo = m["grupo"]
                nota = e["notas"][m["col_excel"]]
                if nota is None:
                    st.markdown(f"**{m['titulo']}** — _sin nota_")
                    continue
                st.markdown(f"**{m['titulo']}** — {desempeno(nota)} · {fmt(nota)}")
                for logro in logros.get(m["clave_word"], []):
                    st.markdown(f"- {logro}")
    st.markdown("#### Boletín de un estudiante (prueba)")
    elegido = st.selectbox("Estudiante", [e["nombre"] for e in estudiantes])
    seleccionado = next(e for e in estudiantes if e["nombre"] == elegido)
    datos_xlsx, adv_boletin = generar_boletin_xlsx(meta, seleccionado, logros)
    for aviso in adv_boletin:
        st.warning(aviso)
    st.download_button(
        "Descargar boletín (Excel)",
        data=datos_xlsx,
        file_name=f"Boletin_{meta['grado']}_{elegido.replace(' ', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Sube la planilla de notas y el Word de logros para ver la revisión.")
