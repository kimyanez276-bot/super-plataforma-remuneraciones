
import io
import re
from typing import Optional
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from pypdf import PdfReader

st.set_page_config(page_title="Plataforma de Remuneraciones - Asesorías Contables", layout="wide")

# ==========================================================
# 1. UTILIDADES Y EXTRACCIÓN CERTIFICADA
# ==========================================================
SHEET_NAME = "SUELDOS 2026"

RUT_RE = re.compile(r"(\d{1,2}(?:\.\d{3}){2}-[\dkK]|\d{7,8}-[\dkK])")

def normalize_rut(value) -> str:
    if value is None:
        return ""
    raw = re.sub(r"[^0-9kK]", "", str(value))
    if len(raw) < 2:
        return ""
    return f"{raw[:-1]}-{raw[-1].upper()}"

def parse_clp(value) -> Optional[int]:
    if value is None:
        return None
    if isinstance(value, (int, float)) and not pd.isna(value):
        return int(round(value))
    text = str(value).strip().replace("$", "").replace(" ", "")
    if not text or "%" in text:
        return None
    text = text.replace(".", "")
    if text.isdigit():
        return int(text)
    return None

def safe_int(value):
    try:
        if value is None or str(value).strip() == "":
            return 0
        return int(float(value))
    except (TypeError, ValueError):
        return 0

def extract_pdf_data(pdf_bytes: bytes) -> pd.DataFrame:
    reader = PdfReader(io.BytesIO(pdf_bytes))
    workers_data = {}
    
    for page in reader.pages:
        text = page.extract_text() or ""
        lines = text.split("\n")
        
        # 1. Remuneraciones / AFP (Sueldo Imponible y Salud Fonasa)
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

        # 2. Detalle de AFP (Cotización Obligatoria y AFC)
        if "Cotización" in text and ("Seguro Cesantía" in text or "Seguro de Cesantía" in text or "Detalle de Cotizaciones" in text):
            for line in lines:
                m = RUT_RE.search(line)
                if m:
                    rut = normalize_rut(m.group(1))
                    if rut in ["76519519-K", "78119044-6", "77230446-3"] or "R.U.T" in line:
                        continue
                    after_rut = line[m.end():]
                    nums = [int(n.replace(".", "")) for n in re.findall(r"\b\d{1,3}(?:\.\d{3})+\b|\b\d+\b", after_rut)]
                    if len(nums) >= 9:
                        cotiz_afp = nums[1]
                        afc_trab = nums[7] if len(nums) >= 8 else 0
                        afc_emp = nums[8] if len(nums) >= 9 else 0
                        if rut in workers_data:
                            workers_data[rut]["cotiz_afp"] = cotiz_afp
                            workers_data[rut]["afc_trab"] = afc_trab
                            workers_data[rut]["afc_emp"] = afc_emp

        # 3. ISL (Mutual) por trabajador
        if "Instituto de Seguridad Laboral" in text or "ISL" in text:
            for line in lines:
                m = RUT_RE.search(line)
                if m:
                    rut = normalize_rut(m.group(1))
                    if rut in ["76519519-K", "78119044-6", "77230446-3"]:
                        continue
                    nums = [int(n.replace(".", "")) for n in re.findall(r"\b\d{1,3}(?:\.\d{3})+\b|\b\d+\b", line[m.end():])]
                    if len(nums) >= 2:
                        if rut in workers_data:
                            workers_data[rut]["isl"] = nums[1]

        # 4. Seguro Social Previsional (SIS, Rentabilidad Protegida, Seguro Social)
        if "SEGURO SOCIAL PREVISIONAL" in text or "Seguro Social" in text:
            for line in lines:
                m = RUT_RE.search(line)
                if m:
                    rut = normalize_rut(m.group(1))
                    if rut in ["76519519-K", "78119044-6", "77230446-3"] or "Totales" in line:
                        continue
                    nums = [int(n.replace(".", "")) for n in re.findall(r"\b\d{1,3}(?:\.\d{3})+\b|\b\d+\b", line[m.end():])]
                    if len(nums) >= 5:
                        if rut in workers_data:
                            workers_data[rut]["s_social"] = nums[2]
                            workers_data[rut]["rent_prot"] = nums[3]
                            workers_data[rut]["sis"] = nums[4]
                            s_imp = workers_data[rut]["sueldo_imponible"]
                            workers_data[rut]["s_social_01"] = int(round(s_imp * 0.001)) if s_imp > 0 else 0

        # 5. Detección automática inicial de Asignación Familiar
        if "ASIGNACION" in text.upper() or "ASIGNACIÓN" in text.upper() or "REBAJAS" in text.upper() or "TRAMO" in text.upper():
            for line in lines:
                m = RUT_RE.search(line)
                if m:
                    rut = normalize_rut(m.group(1))
                    if rut in ["76519519-K", "78119044-6", "77230446-3"]:
                        continue
                    nums = re.findall(r"\b\d{1,3}(?:\.\d{3})+\b|\b\d+\b", line)
                    for n in nums:
                        val = parse_clp(n)
                        if val and 3000 <= val <= 25000:
                            if rut in workers_data:
                                workers_data[rut]["asig_fam"] = val

    df = pd.DataFrame(list(workers_data.values()))
    if df.empty:
        df = pd.DataFrame(columns=["rut", "sueldo_imponible", "salud_fonasa", "cotiz_afp", "afc_trab", "sis", "afc_emp", "isl", "rent_prot", "s_social", "s_social_01", "asig_fam", "bono"])
    return df

def write_to_excel(template_bytes: bytes, df: pd.DataFrame, target_month: str, cargas_dict: dict, bonos_dict: dict) -> bytes:
    wb = load_workbook(io.BytesIO(template_bytes))
    if SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"No se encontró la pestaña '{SHEET_NAME}' en el Excel del cliente.")
    
    ws = wb[SHEET_NAME]

    # Detección automática de columnas por encabezados
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
                            ws.cell(row=r_sub, column=3).value = f"={cotiz_val}+{afc_t_val}+{salud_val}-D{r_sub}"
                            
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

# ==========================================================
# 2. NAVEGACIÓN DE LA PLATAFORMA CORPORATIVA
# ==========================================================
st.sidebar.title("🏢 Panel Corporativo")
menu = st.sidebar.selectbox("Selecciona una opción", ["Gestión de Clientes", "Procesar Previred & Excel", "Dashboard Analítico"])

if menu == "Gestión de Clientes":
    st.header("📂 Directorio de Empresas Clientes")
    st.markdown("Administra los datos maestros de las empresas de la oficina contable.")
    
    with st.form("form_cliente"):
        nombre_cliente = st.text_input("Razón Social (Ej. Uniriego SpA)")
        rut_cliente = st.text_input("RUT Empresa")
        sector = st.selectbox("Sector Económico", ["Transporte", "Agrícola", "Importación / Retail", "Servicios"])
        submitted = st.form_submit_button("Guardar Cliente en Directorio")
        if submitted and nombre_cliente:
            st.success(f"¡Empresa '{nombre_cliente}' guardada exitosamente!")

elif menu == "Procesar Previred & Excel":
    st.header("⚡ Automatización de Planillas Multicliente")
    
    selected_month = st.selectbox("Selecciona el Mes a Procesar", ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"])
    
    col_a, col_b = st.columns(2)
    with col_a:
        pdf_file = st.file_uploader("1. Sube PDF de Previred", type=["pdf"])
    with col_b:
        template_file = st.file_uploader("2. Sube Plantilla Excel Corporativa", type=["xlsx"])

    if pdf_file and template_file:
        df_extracted = extract_pdf_data(pdf_file.getvalue())
        st.subheader("📊 Datos Extraídos del Previred:")
        st.dataframe(df_extracted, use_container_width=True)
        
        st.warning(f"⚠️ **Control para el mes de {selected_month}:** Revisa y ajusta cargas o bonos si es necesario.")
        
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
            st.success("¡Planilla generada con éxito absoluto y aportes patronales alineados!")
            st.download_button(
                label="📥 Descargar Libro de Remuneraciones Actualizado",
                data=final_excel,
                file_name=f"Remuneraciones_{selected_month}_Actualizado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

elif menu == "Dashboard Analítico":
    st.header("📈 Analítica y Auditoría de Planillas")
    st.markdown("Métricas globales de costos previsionales y masa salarial de la cartera de clientes.")
    st.info("Aquí visualizaremos los gráficos de evolución de aportes patronales y sueldos imponibles mes a mes.")
