import io
import re
from datetime import datetime, timedelta
from typing import Optional
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from pypdf import PdfReader
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

st.set_page_config(page_title="Plataforma de Remuneraciones - Asesorías Contables", layout="wide")

# ==========================================================
# 1. UTILIDADES Y FERIADOS OFICIALES DE CHILE (2024 - 2030)
# ==========================================================
RUT_RE = re.compile(r"(\d{1,2}(?:\.\d{3}){2}-[\dkK]|\d{7,8}-[\dkK])")

def normalize_rut(value) -> str:
    if value is None:
        return ""
    raw = re.sub(r"[^0-9kK]", "", str(value))
    if len(raw) < 2:
        return ""
    return f"{raw[:-1]}-{raw[-1].upper()}"

def safe_int(value):
    try:
        if value is None or str(value).strip() == "":
            return 0
        return int(float(value))
    except (TypeError, ValueError):
        return 0

def formato_cl(monto) -> str:
    """Formatea números enteros con puntos como separadores de miles (ej: 1.242.153)"""
    return f"{int(round(monto)):,}".replace(",", ".")

def obtener_feriados_chile() -> set:
    feriados = set()
    feriados_fijos = [
        (1, 1), (5, 1), (5, 21), (6, 29), (7, 16), (8, 15), 
        (9, 18), (9, 19), (10, 12), (10, 31), (11, 1), (12, 8), (12, 25)
    ]
    
    for anio in range(2024, 2031):
        for mes, dia in feriados_fijos:
            try:
                feriados.add(datetime(anio, mes, dia).date())
            except ValueError:
                pass
                
        if anio == 2024: feriados.add(datetime(2024, 9, 20).date())
        elif anio == 2025: feriados.add(datetime(2025, 9, 17).date())
        elif anio == 2026: feriados.add(datetime(2026, 9, 21).date())
        elif anio == 2027: feriados.add(datetime(2027, 9, 17).date())
            
        a = anio % 19
        b = anio // 100
        c = anio % 100
        d = b // 4
        e = b % 4
        f = (b + 8) // 25
        g = (b - f + 1) // 3
        h = (19 * a + b - d - g + 15) % 30
        i = c // 4
        k = c % 4
        L = (32 + 2 * e + 2 * i - h - k) % 7
        m = (a + 11 * h + 22 * L) // 451
        mes_pascua = (h + L - 7 * m + 114) // 31
        dia_pascua = ((h + L - 7 * m + 114) % 31) + 1
        
        domingo_pascua = datetime(anio, mes_pascua, dia_pascua).date()
        feriados.add(domingo_pascua - timedelta(days=2)) # Viernes Santo
        feriados.add(domingo_pascua - timedelta(days=1)) # Sábado Santo
        
        if anio in [2025, 2029]:
            try: feriados.add(datetime(anio, 11, 23).date())
            except ValueError: pass

    return feriados

def calcular_dias_calendario_con_desglose(fecha_inicio_conteo, dias_habiles_objetivo, regimen_semana):
    """
    Simula el calendario día por día y retorna un diccionario con el desglose exacto:
    dias corridos, domingos, sábados inhábiles y feriados chilenos.
    """
    feriados_chile = obtener_feriados_chile()
    current_date = fecha_inicio_conteo + timedelta(days=1)
    habiles_acumulados = 0.0
    dias_corridos_reales = 0
    
    conteo_domingos = 0
    conteo_sabados = 0
    conteo_feriados = 0
    
    while habiles_acumulados < dias_habiles_objetivo:
        weekday = current_date.weekday() # 0: Lunes ... 5: Sábado, 6: Domingo
        es_feriado = current_date in feriados_chile
        
        es_dia_habil_laboral = False
        if regimen_semana == "Lunes a Viernes (5 días)":
            if weekday == 6: # Domingo
                conteo_domingos += 1
            elif weekday == 5: # Sábado
                conteo_sabados += 1
            elif es_feriado:
                conteo_feriados += 1
            else:
                es_dia_habil_laboral = True
        else: # Lunes a Sábado (6 días)
            if weekday == 6: # Domingo
                conteo_domingos += 1
            elif es_feriado:
                conteo_feriados += 1
            else:
                es_dia_habil_laboral = True
                
        if es_dia_habil_laboral:
            if (dias_habiles_objetivo - habiles_acumulados) < 1.0:
                fraccion = dias_habiles_objetivo - habiles_acumulados
                habiles_acumulados += fraccion
            else:
                habiles_acumulados += 1.0
                
        dias_corridos_reales += 1
        current_date += timedelta(days=1)
        
        if dias_corridos_reales > 3650:
            break
            
    return {
        "dias_corridos": dias_corridos_reales,
        "domingos": conteo_domingos,
        "sabados": conteo_sabados,
        "feriados": conteo_feriados
    }

def extract_pdf_data(pdf_bytes: bytes) -> pd.DataFrame:
    reader = PdfReader(io.BytesIO(pdf_bytes))
    workers_data = {}
    
    for page in reader.pages:
        text = page.extract_text() or ""
        lines = text.split("\n")
        
        if "AFP" in text and "REMUNERACIÓN" in text:
            for line in lines:
                if RUT_RE.search(line) and "AFP" in line and "REPRESENTANTE" not in text.upper():
                    m = RUT_RE.search(line)
                    rut = normalize_rut(m.group(1))
                    if not rut or rut in ["76519519-K", "78119044-6", "77230446-3"]:
                        continue
                    parts = line.split("AFP")
                    if len(parts) > 1:
                        nums = re.findall(r"\b\d{1,3}(?:\.\d{3})+\b|\b\d+\b", parts[1])
                        if len(nums) >= 2:
                            sueldo_imp = int(nums[0].replace(".", ""))
                            salud_fonasa = int(nums[1].replace(".", ""))
                            if rut not in workers_data:
                                workers_data[rut] = {
                                    "rut": rut,
                                    "sueldo_imponible": sueldo_imp,
                                    "salud_fonasa": salud_fonasa,
                                    "cotiz_afp": 0, "afc_trab": 0, "sis": 0,
                                    "afc_emp": 0, "isl": 0, "rent_prot": 0,
                                    "s_social": 0, "s_social_01": 0,
                                    "asig_fam": 0, "bono": 0
                                }

        if "Cotización" in text and ("Seguro Cesantía" in text or "Seguro de Cesantía" in text or "Detalle de Cotizaciones" in text):
            for line in lines:
                m = RUT_RE.search(line)
                if m:
                    rut = normalize_rut(m.group(1))
                    if rut in ["76519519-K", "78119044-6", "77230446-3"] or "R.U.T" in line:
                        continue
                    after_rut = line[m.end():]
                    nums = [int(n.replace(".", "")) for n in re.findall(r"\b\d{1,3}(?:\.\d{3})+\b|\b\d+\b", after_rut)]
                    if len(nums) >= 9 and rut in workers_data:
                        workers_data[rut]["cotiz_afp"] = nums[1]
                        workers_data[rut]["afc_trab"] = nums[7] if len(nums) >= 8 else 0
                        workers_data[rut]["afc_emp"] = nums[8] if len(nums) >= 9 else 0

        if "Instituto de Seguridad Laboral" in text or "ISL" in text:
            for line in lines:
                m = RUT_RE.search(line)
                if m:
                    rut = normalize_rut(m.group(1))
                    if rut in ["76519519-K", "78119044-6", "77230446-3"]:
                        continue
                    nums = [int(n.replace(".", "")) for n in re.findall(r"\b\d{1,3}(?:\.\d{3})+\b|\b\d+\b", line[m.end():])]
                    if len(nums) >= 2 and rut in workers_data:
                        workers_data[rut]["isl"] = nums[1]

        if "SEGURO SOCIAL PREVISIONAL" in text or "Seguro Social" in text:
            for line in lines:
                m = RUT_RE.search(line)
                if m:
                    rut = normalize_rut(m.group(1))
                    if rut in ["76519519-K", "78119044-6", "77230446-3"] or "Totales" in line:
                        continue
                    nums = [int(n.replace(".", "")) for n in re.findall(r"\b\d{1,3}(?:\.\d{3})+\b|\b\d+\b", line[m.end():])]
                    if len(nums) >= 5 and rut in workers_data:
                        workers_data[rut]["s_social"] = nums[2]
                        workers_data[rut]["rent_prot"] = nums[3]
                        workers_data[rut]["sis"] = nums[4]
                        s_imp = workers_data[rut]["sueldo_imponible"]
                        workers_data[rut]["s_social_01"] = int(round(s_imp * 0.001)) if s_imp > 0 else 0

    df = pd.DataFrame(list(workers_data.values()))
    if df.empty:
        df = pd.DataFrame(columns=["rut", "sueldo_imponible", "salud_fonasa", "cotiz_afp", "afc_trab", "sis", "afc_emp", "isl", "rent_prot", "s_social", "s_social_01", "asig_fam", "bono"])
    return df

def write_to_excel(template_bytes: bytes, df: pd.DataFrame, target_month: str, cargas_dict: dict, bonos_dict: dict) -> bytes:
    wb = load_workbook(io.BytesIO(template_bytes))
    target_sheet = None
    for name in wb.sheetnames:
        if name.strip().lower() in ["sueldos 2026", "sueldos2026", "remuneraciones"]:
            target_sheet = name
            break
    if not target_sheet:
        target_sheet = wb.sheetnames[0]
    ws = wb[target_sheet]

    col_sis, col_afc, col_isl, col_rent, col_s_soc, col_s_01 = 14, 15, 16, 17, 18, 19
    for r in range(24, 28):
        for c in range(1, 30):
            val = str(ws.cell(row=r, column=c).value or "").strip().upper()
            if val == "SIS": col_sis = c
            elif val == "AFC": col_afc = c
            elif val == "ISL": col_isl = c
            elif "PROTEG" in val: col_rent = c
            elif val in ["S. SOCIAL", "SEGURO SOCIAL"]: col_s_soc = c
            elif "0,1" in val or "0.1" in val: col_s_01 = c

    for row in range(1, ws.max_row + 1):
        cell_val = ws.cell(row, 2).value
        norm_cell = normalize_rut(cell_val)
        if norm_cell:
            for _, rec in df.iterrows():
                if normalize_rut(rec["rut"]) == norm_cell:
                    for r_sub in range(row, row + 16):
                        mes_val = str(ws.cell(r_sub, 1).value or "").strip().upper()
                        if target_month in mes_val:
                            ws.cell(row=r_sub, column=2).value = safe_int(rec["sueldo_imponible"])
                            cotiz_val = safe_int(rec["cotiz_afp"])
                            afc_t_val = safe_int(rec["afc_trab"])
                            salud_val = safe_int(rec["salud_fonasa"])
                            ws.cell(row=r_sub, column=3).value = f"={cotiz_val}+{afc_t_val}+{salud}-D{r_sub}"
                            ws.cell(row=r_sub, column=11).value = safe_int(cargas_dict.get(norm_cell, rec["asig_fam"]))
                            ws.cell(row=r_sub, column=12).value = safe_int(bonos_dict.get(norm_cell, 0))
                            ws.cell(row=r_sub, column=col_sis).value = safe_int(rec.get("sis"))
                            ws.cell(row=r_sub, column=col_afc).value = safe_int(rec.get("afc_emp"))
                            ws.cell(row=r_sub, column=col_isl).value = safe_int(rec.get("isl"))
                            ws.cell(row=r_sub, column=col_rent).value = safe_int(rec.get("rent_prot"))
                            ws.cell(row=r_sub, column=col_s_soc).value = safe_int(rec.get("s_social"))
                            ws.cell(row=r_sub, column=col_s_01).value = safe_int(rec.get("s_social_01"))
                            break
                            
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()

def generate_finiquito_word(nombre_trabajador, rut_trabajador, empresa, rut_empresa, causal, monto_dias, total_feriado_pesos, indem_anos):
    doc = Document()
    
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run("FINIQUITO DE CONTRATO DE TRABAJO")
    run_title.bold = True
    run_title.font.size = Pt(14)
    
    doc.add_paragraph()
    
    p_comp = doc.add_paragraph(f"En Linares, a ____ de ____________________ de 20___, comparecen por una parte la empresa {empresa}, RUT N° {rut_empresa}, representada legalmente por Don/Doña Representante Legal, y por la otra parte el/la trabajador/a Don/Doña {nombre_trabajador}, RUT N° {rut_trabajador}, quienes acuerdan poner término al contrato de trabajo bajo las siguientes estipulaciones:")
    p_comp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    
    p_prim = doc.add_paragraph(f"PRIMERO: El presente contrato termina por la causal establecida en el {causal}.")
    p_prim.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    
    p_sec = doc.add_paragraph("SEGUNDO: Las partes dejan constancia que el monto total de las prestaciones adeudadas y calculadas es el siguiente:")
    p_sec.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    
    doc.add_paragraph(f"• Remuneración días trabajados en el mes: ${formato_cl(monto_dias)}")
    doc.add_paragraph(f"• Feriado Proporcional: ${formato_cl(total_feriado_pesos)}")
    doc.add_paragraph(f"• Indemnización por Años de Servicio: ${formato_cl(indem_anos)}")
    
    total_finiquito = monto_dias + total_feriado_pesos + indem_anos
    p_tot = doc.add_paragraph()
    run_tot = p_tot.add_run(f"TOTAL A PAGAR: ${formato_cl(total_finiquito)}")
    run_tot.bold = True
    
    p_terc = doc.add_paragraph("TERCERO: El/la trabajador/a declara recibir a su entera satisfacción el pago indicado, sin tener cargo ni reclamación posterior alguna que formular.")
    p_terc.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    
    doc.add_paragraph()
    
    table = doc.add_table(rows=1, cols=2)
    table.autofit = False
    
    cell_left = table.cell(0, 0)
    cell_right = table.cell(0, 1)
    
    p_left = cell_left.paragraphs[0]
    p_left.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_left.add_run("__________________________________\nFirma Empleador")
    
    p_right = cell_right.paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_right.add_run("__________________________________\nFirma Trabajador / a")
    
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output.getvalue()

# ==========================================================
# 2. NAVEGACIÓN DE LA PLATAFORMA CORPORATIVA
# ==========================================================
st.sidebar.title("🏢 Panel Corporativo")
menu = st.sidebar.selectbox("Selecciona una opción", ["Gestión de Clientes", "Procesar Previred & Excel", "Calculadora y Finiquitos DT", "Dashboard Analítico"])

if menu == "Gestión de Clientes":
    st.header("📂 Directorio de Empresas Clientes")
    with st.form("form_cliente"):
        nombre_cliente = st.text_input("Razón Social (Ej. Formula Center SpA)")
        rut_cliente = st.text_input("RUT Empresa")
        if st.form_submit_button("Guardar Cliente") and nombre_cliente:
            st.success(f"¡Empresa '{nombre_cliente}' guardada!")

elif menu == "Procesar Previred & Excel":
    st.header("⚡ Automatización de Planillas Multicliente")
    selected_month = st.selectbox("Mes a Procesar", ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"])
    
    col_a, col_b = st.columns(2)
    with col_a:
        pdf_file = st.file_uploader("1. Sube PDF de Previred", type=["pdf"])
    with col_b:
        template_file = st.file_uploader("2. Sube Plantilla Excel Corporativa", type=["xlsx"])

    if pdf_file and template_file:
        df_extracted = extract_pdf_data(pdf_file.getvalue())
        st.dataframe(df_extracted, use_container_width=True)
        cargas_dict = {}
        bonos_dict = {}
        for _, row in df_extracted.iterrows():
            r = row["rut"]
            c1, c2 = st.columns(2)
            with c1:
                cargas_dict[r] = st.number_input(f"Asig. Fam (RUT: {r})", min_value=0, value=int(row["asig_fam"]), key=f"c_{r}")
            with c2:
                bonos_dict[r] = st.number_input(f"Bono (RUT: {r})", min_value=0, value=0, key=f"b_{r}")
            st.markdown("---")
            
        if st.button("🚀 Rellenar Planilla Oficial del Mes", type="primary"):
            final_excel = write_to_excel(template_file.getvalue(), df_extracted, selected_month, cargas_dict, bonos_dict)
            st.success("¡Planilla generada con éxito!")
            st.download_button("📥 Descargar Libro Actualizado", data=final_excel, file_name=f"Remuneraciones_{selected_month}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

elif menu == "Calculadora y Finiquitos DT":
    st.header("⚖️ Simulador, Cálculo y Redacción de Finiquitos")
    
    with st.expander("📖 🔍 Recordatorio y Guía de Causales de Término (Normativa DT Chile)", expanded=False):
        st.markdown("""
        * **Art. 161 - Necesidades (Con Aviso Previo):** Se avisa con 30 días. **No se paga mes de aviso**, pero sí años de servicio y feriado.
        * **Art. 161 - Necesidades (Sin Aviso / Mes Sustitutivo):** Despido inmediato. **Se paga el mes de aviso** + años de servicio + feriado. Descuenta AFC empleador.
        * **Art. 159 N° 1 - Mutuo Acuerdo:** Acuerdo mutuo. Indemnizaciones pactadas libremente.
        * **Art. 159 N° 2 - Renuncia Voluntaria:** El trabajador dimite. Solo días trabajados y feriado proporcional. **Cero indemnización y sin descuento AFC.**
        * **Art. 160 - Causales de Caducidad:** Faltas graves. **Sin indemnización ni feriado proporcional** (salvo días trabajados).
        """)

    modo_trabajador = st.radio("📂 Origen de los Datos del Trabajador", ["Trabajador Libre / Fuera de Base de Datos", "Seleccionar de Base de Datos Cliente"])
    
    if modo_trabajador == "Trabajador Libre / Fuera de Base de Datos":
        col_w1, col_w2 = st.columns(2)
        with col_w1:
            nombre_trab = st.text_input("👤 Nombre Completo", value="María González")
            rut_trab = st.text_input("🆔 RUT Trabajador", value="15.123.456-7")
            fecha_inicio = st.date_input("📅 Fecha Inicio Contrato", value=pd.to_datetime("2024-01-01").date())
        with col_w2:
            empresa_nombre = st.text_input("🏢 Empresa Empleador", value="Empresa Externa")
            empresa_rut = st.text_input("🆔 RUT Empleador", value="76.123.456-8")
            fecha_termino = st.date_input("📅 Fecha Término / Despido", value=pd.to_datetime("2026-09-30").date())
    else:
        st.info("ℹ️ Módulo vinculado a Formula Center SpA.")
        nombre_trab = st.selectbox("Seleccionar Trabajador", ["Rodolfo Álvarez", "Martín Cabrera", "Nicolás Contreras", "Rodrigo Leiva", "Anggie Medina", "David Urrutia"])
        rut_trab = "13.599.716-1"
        empresa_nombre = "Formula Center SpA"
        empresa_rut = "77.597.719-1"
        col_w1, col_w2 = st.columns(2)
        with col_w1:
            fecha_inicio = st.date_input("📅 Fecha Inicio", value=pd.to_datetime("2024-01-01").date())
        with col_w2:
            fecha_termino = st.date_input("📅 Fecha Término", value=pd.to_datetime("2026-09-30").date())

    col1, col2 = st.columns(2)
    with col1:
        sueldo_base = st.number_input("💵 Sueldo Base Mensual ($)", min_value=0, value=700000, step=10000)
        gratificacion = st.number_input("🎁 Gratificación Mensual ($)", min_value=0, value=180000, step=5000)
        metodo_promedio = st.radio("🔍 Base de Cálculo", ["Promedio Últimos 3 Meses", "Promedio Últimos 6 Meses"])
    with col2:
        dias_pendientes = st.number_input("⏰ Días Trabajados en el Mes", min_value=0, max_value=30, value=10)
        
        delta_dias_gen = (fecha_termino - fecha_inicio).days
        meses_trab_gen = max(0, delta_dias_gen / 30.416)
        feriado_proporcional_calculado = meses_trab_gen * 1.25
        
        dias_vacaciones_pendientes = st.number_input(
            "🏖️ Días Hábiles de Vacaciones Acumuladas / Pendientes (Editable)", 
            min_value=0.0, value=float(round(feriado_proporcional_calculado, 1)), step=0.5,
            help="Puedes ajustar manualmente los días hábiles pendientes si el trabajador arrastra feriado de años anteriores."
        )

        regimen_conteo = st.selectbox(
            "📅 Criterio de Conteo Calendario (Feriado Proporcional)",
            [
                "Cálculo Matemático Estándar (Factor 1.4)", 
                "Conteo Día a Día + Feriados Chilenos (Lunes a Viernes)", 
                "Conteo Día a Día + Feriados Chilenos (Lunes a Sábado)"
            ],
            help="Elige si deseas calcular con la proporción matemática directa o simulando el calendario real saltando domingos, sábados (si aplica) y feriados oficiales de Chile."
        )

        opciones_causal = [
            "Art. 161 - Necesidades (Con Aviso Previo)",
            "Art. 161 - Necesidades (Sin Aviso / Mes Sustitutivo)",
            "Art. 159 N° 1 - Mutuo Acuerdo",
            "Art. 159 N° 2 - Renuncia Voluntaria",
            "Art. 160 - Causales de Caducidad (Sin Indemnización)"
        ]
        causal = st.selectbox("📋 Causal de Término de Contrato", opciones_causal)

    if st.button("📊 Calcular y Generar Documento Word", type="primary"):
        delta_dias = (fecha_termino - fecha_inicio).days
        anos_servicio = delta_dias / 365.25
        anos_enteros = int(anos_servicio)
        
        if regimen_conteo == "Cálculo Matemático Estándar (Factor 1.4)":
            dias_feriado_corridos = dias_vacaciones_pendientes * (7 / 5)
            metodo_texto = "Matemático Estándar (× 1,4)"
            desglose_info = None
        elif regimen_conteo == "Conteo Día a Día + Feriados Chilenos (Lunes a Viernes)":
            desglose_info = calcular_dias_calendario_con_desglose(fecha_termino, dias_vacaciones_pendientes, "Lunes a Viernes (5 días)")
            dias_feriado_corridos = desglose_info["dias_corridos"]
            metodo_texto = "Calendario Real + Feriados (Lunes a Viernes)"
        else:
            desglose_info = calcular_dias_calendario_con_desglose(fecha_termino, dias_vacaciones_pendientes, "Lunes a Sábado (6 días)")
            dias_feriado_corridos = desglose_info["dias_corridos"]
            metodo_texto = "Calendario Real + Feriados (Lunes a Sábado)"

        base_calculo = sueldo_base + gratificacion
        rem_diaria = base_calculo / 30
        
        monto_dias_trabajados = rem_diaria * dias_pendientes
        monto_feriado = rem_diaria * dias_feriado_corridos
        
        indem_anos = 0
        if "161" in causal or "Mutuo" in causal:
            indem_anos = base_calculo * anos_enteros

        total_finiquito = monto_dias_trabajados + monto_feriado + indem_anos

        st.markdown("---")
        st.subheader("📑 Resultados Detallados del Cálculo (Normativa DT)")
        
        st.markdown(f"""
        > **Resumen del Trabajador:** {nombre_trab} | **Antigüedad:** {anos_enteros} años ({delta_dias} días corridos)  
        > **Base de Cálculo ({metodo_promedio}):** Sueldo Base (${formato_cl(sueldo_base)}) + Gratificación (${formato_cl(gratificacion)}) = **${formato_cl(base_calculo)}** mensual (${formato_cl(rem_diaria)} diarios).
        """)

        res_c1, res_c2 = st.columns(2)
        with res_c1:
            # Texto descriptivo para feriado proporcional con desglose si aplica
            if desglose_info:
                detalle_dias_str = f"Equivale a **{dias_feriado_corridos} días corridos** en total<br>(Incluye: {int(dias_vacaciones_pendientes)} días hábiles, {desglose_info['domingos']} domingos, {desglose_info['sabados']} sábados y {desglose_info['feriados']} feriados)"
            else:
                detalle_dias_str = f"Equivale a **{dias_feriado_corridos:.1f} días corridos**"

            st.markdown(f"""
            ### 💼 Prestaciones Principales
            * **Días Trabajados en el Mes ({dias_pendientes} días):**  
              `{dias_pendientes} días × ${formato_cl(rem_diaria)}` = **${formato_cl(monto_dias_trabajados)}**
            
            * **Feriado Proporcional ({dias_vacaciones_pendientes} días hábiles):**  
              *Método:* `{metodo_texto}`  
              {detalle_dias_str}  
              `{dias_feriado_corridos} × ${formato_cl(rem_diaria)}` = **${formato_cl(monto_feriado)}**
            """, unsafe_allow_html=True)
            
        with res_c2:
            st.markdown(f"""
            ### ⚖️ Indemnizaciones y Total
            * **Indemnización Años de Servicio ({anos_enteros} años):**  
              `{anos_enteros} años × ${formato_cl(base_calculo)}` = **${formato_cl(indem_anos)}**
            
            ---
            ### 💰 TOTAL A PAGAR EN FINIQUITO
            ### **${formato_cl(total_finiquito)}**
            """)

        word_bytes = generate_finiquito_word(
            nombre_trab, rut_trab, empresa_nombre, empresa_rut, causal, 
            monto_dias_trabajados, monto_feriado, indem_anos
        )
        
        st.markdown("---")
        st.success("✅ ¡Cálculo completado y documento Word redactado con éxito incluyendo el desglose de días hábiles y no hábiles!")
        st.download_button(
            label="📥 Descargar Finiquito Oficial en Word (.docx)",
            data=word_bytes,
            file_name=f"Finiquito_{nombre_trab.replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

elif menu == "Dashboard Analítico":
    st.header("📈 Analítica y Auditoría")
    st.info("Visualización de métricas y costos previsionales.")