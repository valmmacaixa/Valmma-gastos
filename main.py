import os
import io
import time
import json
import tempfile
import unicodedata
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

SESION_IMAGENES = {}

def obtener_cliente_genai():
    api_key = os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY", "")
    if not api_key:
        raise HTTPException(status_code=500, detail="Falta la API Key de Gemini en las variables de entorno de Render.")
    return genai.Client(api_key=api_key)

def limpiar_y_formatear_localidad(loc_str):
    if not loc_str:
        return ""
    s = str(loc_str).strip()
    invalidos = ["desconocido", "desconocida", "no especificada", "no especificado", "no disponible", "n/a", "-"]
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
            enhancer = Image