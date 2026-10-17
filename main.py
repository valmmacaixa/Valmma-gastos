import os
import io
import time
import json
import tempfile
import unicodedata
import asyncio
from datetime import datetime
from typing import List, Optional

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from PIL import Image, ImageEnhance, ImageOps
import openpyxl
import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from google import genai

app = FastAPI(title="Valmma Gastos API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

API_KEY = os.getenv("GEMINI_API_KEY", "")
SESION_IMAGENES = {}

def limpiar_y_formatear_localidad(loc_str):
    if not loc_str:
        return ""
    s = str(loc_str).strip()
    invalidos = ["desconocido", "desconocida", "no especificada", "no especificado", "no disponible", "n/a", "-", "no indicada"]
    if s.lower() in invalidos:
        return ""

    nfkd_form = unicodedata.normalize('NFKD', s)
    s_sin_tildes = "".join([c for c in nfkd_form if not unicodedata.combining(c)])

    preposiciones = {"de", "del", "el", "la", "los", "las", "en", "a", "por", "para", "con"}
    palabras = s_sin_tildes.split()
    palabras_formateadas = []
    for idx, p in enumerate(palabras):
        p_limpia = p.strip()
        if not p_limpia:
            continue
        if idx > 0 and p_limpia.lower() in preposiciones:
            palabras_formateadas.append(p_limpia.lower())
        else:
            palabras_formateadas.append(p_limpia.capitalize())
    return " ".join(palabras_formateadas)

def optimizar_para_ocr(bytes_imagen):
    try:
        with Image.open(io.BytesIO(bytes_imagen)) as img:
            img = ImageOps.exif_transpose(img)
            img = img.convert('L')
            max_dim = 1000
            if max(img.size) > max_dim:
                img.thumbnail((max_dim, max_dim), Image.Resampling.LANCZOS)
            enhancer = ImageEnhance.Contrast(img)
            img = enhancer.enhance(1.3)
            out_io = io.BytesIO()
            img.save(out_io, format='JPEG', quality=75, optimize=True)
            out_io.seek(0)
            return out_io
    except Exception:
        return io.BytesIO(bytes_imagen)

def convertir_a_escaneo_bn(bytes_imagen):
    try:
        with Image.open(io.BytesIO(bytes_imagen)) as img:
            img = ImageOps.exif_transpose(img)
            gray = img.convert('L')
            enhancer = ImageEnhance.Contrast(gray)
            img_bn = enhancer.enhance(1.5)
            
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".jpg")
            temp_path = temp_file.name
            temp_file.close()
            img_bn.save(temp_path, format='JPEG', quality=80, optimize=True)
            return temp_path
    except Exception:
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".jpg")
        temp_path = temp_file.name
        temp_file.write(bytes_imagen)
        temp_file.close()
        return temp_path

def parsear_fecha(str_fecha):
    if not str_fecha:
        return None
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y"):
        try:
            return datetime.strptime(str_fecha.strip(), fmt)
        except Exception:
            pass
    return None

@app.post("/api/procesar")
async def procesar_documentos(files: List[UploadFile] = File(...)):
    global SESION_IMAGENES
    
    if not API_KEY:
        raise HTTPException(status_code=500, detail="GEMINI_API_KEY no configurada.")
        
    client = genai.Client(api_key=API_KEY)
    
    prompt_base = """
    Analiza la siguiente imagen de un ticket, factura o comprobante de gasto.
    
    REGLAS ESTRICTAS DE EXTRACCIÓN:
    1. TIPO DE DOCUMENTO ("tipo_doc"):
       - "FACTURA": ÚNICAMENTE si es una factura formal con NIF/CIF del emisor, numeración de factura y datos fiscales completos. Las facturas simplificadas o tickets de caja NO son facturas, deben ir como "TICKET".
       - "TICKET": Tiques de caja, restaurantes, parkings, peajes, taxis o facturas simplificadas.
    2. IGNORAR COMPROBANTES DE TARJETA (DATÁFONO): Si la imagen es exclusivamente el resguardo impreso de pago con tarjeta (el papel secundario con códigos de autorización, VISA/Mastercard, etc.) y no el ticket comercial de la compra, IGNÓRALO devolviendo importe 0 para evitar duplicar el gasto.
    3. LOCALIDAD: Si la localidad o ciudad no aparece de forma explícita en el texto, deja el campo estrictamente como "" (cadena vacía). Jamás pongas "No indicada" ni inventes datos.
    
    Extrae los datos en formato JSON estricto con las claves:
    - fecha: (DD/MM/YYYY)
    - localidad: (Nombre del municipio o "" si no consta)
    - tipo_doc: ("FACTURA" o "TICKET")
    - forma_pago: ("TARJETA" o "CONTADO")
    - importe: (número decimal con el total)
    - descripcion: (nombre del establecimiento o concepto breve)
    """

    async def procesar_archivo(idx: int, file: UploadFile):
        try:
            contenido = await file.read()
            img_id = f"img_{time.time()}_{idx}"
            SESION_IMAGENES[img_id] = contenido

            buf_opt = optimizar_para_ocr(contenido)
            img_pil = Image.open(buf_opt)

            response = client.models.generate_content(
                model="gemini-3.8-flash",
                contents=[prompt_base, img_pil]
            )
            raw = response.text.replace("```json", "").replace("```", "").strip()
            data = json.loads(raw)
            
            if not data or float(data.get("importe", 0)) == 0:
                return None

            data["img_id"] = img_id
            data["filename"] = file.filename
            data["localidad"] = limpiar_y_formatear_localidad(data.get("localidad", ""))
            return data
        except Exception as e:
            print(f"Error OCR {file.filename}: {e}")
            return None

    # Procesar todos los archivos en paralelo de forma eficiente
    tareas = [procesar_archivo(idx, f) for idx, f in enumerate(files)]
    resultados = await asyncio.gather(*tareas)
    
    items_extraidos = [r for r in resultados if r is not None]

    con_f = [x for x in items_extraidos if parsear_fecha(x.get("fecha"))]
    sin_f = [x for x in items_extraidos if not parsear_fecha(x.get("fecha"))]
    con_f_ord = sorted(con_f, key=lambda k: parsear_fecha(k.get("fecha")))

    resultado = con_f_ord + sin_f

    return {"status": "ok", "items": resultado}

@app.post("/api/generar-excel")
async def generar_excel(payload: dict):
    nombre_usuario = payload.get("nombre", "Javier Cosme").strip() or "Javier Cosme"
    otros_trabajadores = payload.get("otros", "").strip()
    items = payload.get("items", [])

    ruta_plantilla = "Gastos_Viaje_VALMMA_V5.xlsx"
    wb = openpyxl.load_workbook(ruta_plantilla)
    ws = wb["GASTOS VIAJE"]

    ws.cell(row=3, column=4, value=time.strftime("%d/%m/%Y"))
    ws.cell(row=7, column=3, value=nombre_usuario)
    ws.cell(row=7, column=7, value=otros_trabajadores)

    localidades = []
    for it in items:
        loc = limpiar_y_formatear_localidad(it.get("localidad", ""))
        if loc and loc not in localidades:
            localidades.append(loc)

    ws.cell(row=8, column=3, value=" - ".join(localidades))

    fechas = [it.get("fecha") for it in items if parsear_fecha(it.get("fecha"))]
    ws.cell(row=8, column=8, value=fechas[0] if fechas else "")
    ws.cell(row=8, column=11, value=fechas[-1] if fechas else "")

    for idx, item in enumerate(items):
        r = 20 + idx
        ws.cell(row=r, column=1, value=item.get("fecha", ""))
        ws.cell(row=r, column=2, value=limpiar_y_formatear_localidad(item.get("localidad", "")))
        ws.cell(row=r, column=3, value=item.get("concepto", "Estancia y manutención + Aparc./autop. + Desplazamiento"))
        ws.cell(row=r, column=4, value=item.get("descripcion", ""))
        ws.cell(row=r, column=5, value=otros_trabajadores)

        pago = str(item.get("forma_pago", "")).upper()
        tipo = str(item.get("tipo_doc", "")).upper()
        imp = float(item.get("importe", 0.0))

        if "TARJETA" in pago:
            ws.cell(row=r, column=8 if "FACTURA" in tipo else 7, value=imp)
        else:
            ws.cell(row=r, column=11 if "FACTURA" in tipo else 10, value=imp)

    temp_out = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    wb.save(temp_out.name)
    
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    nombre_archivo = f"Gastos_Valmma_{nombre_usuario.replace(' ', '_')}_{timestamp}.xlsx"

    return FileResponse(
        temp_out.name, 
        filename=nombre_archivo,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.post("/api/generar-doc")
async def generar_documento_word(payload: dict):
    nombre_usuario = payload.get("nombre", "Javier Cosme").strip() or "Javier Cosme"
    es_ticket = payload.get("es_ticket", False)
    items = payload.get("items", [])

    items_filtrados = []
    for it in items:
        es_factura = "FACTURA" in str(it.get("tipo_doc", "")).upper()
        if es_ticket:
            if not es_factura:
                items_filtrados.append(it)
        else:
            if es_factura:
                items_filtrados.append(it)

    if not items_filtrados:
        raise HTTPException(status_code=400, detail="No hay documentos de este tipo.")

    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    archivos_temp = []
    for item in items_filtrados:
        img_id = item.get("img_id")
        if img_id in SESION_IMAGENES:
            bytes_img = SESION_IMAGENES[img_id]
            tmp_bn = convertir_a_escaneo_bn(bytes_img)
            archivos_temp.append(tmp_bn)

            try:
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p_img.paragraph_format.space_after = Pt(8)
                p_img.paragraph_format.space_before = Pt(0)
                p_img.add_run().add_picture(tmp_bn, width=Inches(4.5))
            except Exception:
                pass

    temp_doc = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    doc.save(temp_doc.name)

    for tmp in archivos_temp:
        try:
            os.remove(tmp)
        except Exception:
            pass

    prefijo = "Anexo_Tickets" if es_ticket else "Anexo_Facturas"
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    nombre_doc = f"{prefijo}_{nombre_usuario.replace(' ', '_')}_{timestamp}.docx"

    return FileResponse(
        temp_doc.name,
        filename=nombre_doc,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )