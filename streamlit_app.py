from __future__ import annotations

import io
import re
import shutil
import subprocess
import tempfile
import unicodedata
from collections import Counter
from datetime import date
from functools import lru_cache
from pathlib import Path
from typing import Any

import fitz
import pandas as pd
import streamlit as st
from PIL import Image, ImageEnhance, ImageFilter, ImageOps
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

APP_NAME = "Rendición Inteligente"
APP_VERSION = "0.3.0"

COLUMNAS = [
    "archivo",
    "pagina",
    "tipo_documento",
    "rut_emisor",
    "razon_social",
    "fecha",
    "folio",
    "monto_neto",
    "monto_exento",
    "monto_iva",
    "otros_impuestos",
    "monto_total",
    "glosa",
    "estado",
    "observacion",
]


def configurar_pagina() -> None:
    st.set_page_config(
        page_title=APP_NAME,
        page_icon="🧾",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    st.markdown(
        """
        <style>
        .block-container {
            padding-top: 1.8rem;
            padding-bottom: 3rem;
            max-width: 1450px;
        }
        .hero {
            padding: 1.45rem 1.6rem;
            border: 1px solid rgba(128,128,128,.22);
            border-radius: 18px;
            background: linear-gradient(135deg, rgba(26,115,232,.10), rgba(0,184,148,.08));
            margin-bottom: 1.2rem;
        }
        .hero h1 {
            margin: 0 0 .35rem 0;
            font-size: 2rem;
        }
        .hero p {
            margin: 0;
            opacity: .82;
            font-size: 1.02rem;
        }
        div[data-testid="stMetric"] {
            border: 1px solid rgba(128,128,128,.20);
            padding: 1rem;
            border-radius: 14px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def inicializar_estado() -> None:
    if "resultados" not in st.session_state:
        st.session_state.resultados = pd.DataFrame(columns=COLUMNAS)
    if "procesamiento_realizado" not in st.session_state:
        st.session_state.procesamiento_realizado = False


def _texto_busqueda(texto: str) -> str:
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return texto.upper().replace("–", "-").replace("—", "-")


def _lineas_utiles(texto: str) -> list[str]:
    return [re.sub(r"\s+", " ", linea).strip() for linea in texto.splitlines() if linea.strip()]


def _digito_verificador(cuerpo: str) -> str:
    suma = 0
    factor = 2
    for digito in reversed(cuerpo):
        suma += int(digito) * factor
        factor = factor + 1 if factor < 7 else 2
    resto = 11 - suma % 11
    if resto == 11:
        return "0"
    if resto == 10:
        return "K"
    return str(resto)


def _formatear_rut(cuerpo: str, dv: str) -> str:
    grupos: list[str] = []
    restante = cuerpo
    while restante:
        grupos.insert(0, restante[-3:])
        restante = restante[:-3]
    return f"{'.'.join(grupos)}-{dv.upper()}"


def _extraer_rut(textos: list[str]) -> tuple[str, bool]:
    patron = re.compile(r"(?<!\d)([0-9][0-9.\s]{5,15})\s*[-~:]\s*([0-9K])(?!\w)", re.I)
    candidatos: list[tuple[int, str, bool]] = []

    for texto in textos:
        buscable = _texto_busqueda(texto)
        posicion_factura = buscable.find("FACTURA")
        posicion_boleta = buscable.find("BOLETA")
        
        for coincidencia in patron.finditer(buscable):
            cuerpo_leido = re.sub(r"\D", "", coincidencia.group(1))
            dv_leido = coincidencia.group(2).upper()
            
            zona_previa = buscable[max(0, coincidencia.start() - 25):coincidencia.start()]
            cerca_de_rut = any(term in zona_previa for term in ["RUT", "R.U.T", "EMISOR", "R.U.N"])

            longitudes = [len(cuerpo_leido)] if len(cuerpo_leido) <= 8 else [8, 7]
            for largo in longitudes:
                if largo < 6 or len(cuerpo_leido) < largo:
                    continue
                cuerpo = cuerpo_leido[-largo:]
                dv_calculado = _digito_verificador(cuerpo)
                valido = (dv_calculado == dv_leido)
                
                antes_de_doc = (
                    (posicion_factura >= 0 and coincidencia.start() < posicion_factura) or
                    (posicion_boleta >= 0 and coincidencia.start() < posicion_boleta)
                )
                
                puntaje = 25
                puntaje += 160 if valido else 0
                puntaje += 60 if antes_de_doc else 0
                puntaje += 35 if cerca_de_rut else 0
                puntaje += 15 if largo == 8 else 0

                candidatos.append(
                    (puntaje, _formatear_rut(cuerpo, dv_leido if valido else dv_calculado), not valido)
                )

    if not candidatos:
        return "", False

    mejor_por_rut: dict[str, tuple[int, bool, int]] = {}
    for puntaje, rut, corregido in candidatos:
        previo = mejor_por_rut.get(rut, (0, corregido, 0))
        mejor_por_rut[rut] = (max(previo[0], puntaje), previo[1] and corregido, previo[2] + 1)

    rut, (puntaje, corregido, repeticiones) = max(
        mejor_por_rut.items(), key=lambda item: (item[1][0], item[1][2], len(item[0]))
    )
    return rut, corregido


def _tipo_documento(textos: list[str]) -> str:
    buscable = "\n".join(_texto_busqueda(t) for t in textos)
    if "FACTURA" in buscable or "ACTURA ELECTRON" in buscable:
        return "Factura"
    if "BOLETA" in buscable or "COMPROBANTE DE PAGO" in buscable or "TICKET" in buscable:
        return "Boleta"
    if any(t in buscable for t in ["VOUCHER", "TRANSBANK", "REDELCOM", "SUMUP", "KLAP", "COMPRA CON TARJETA"]):
        return "Voucher"
    return "Boleta" if "TOTAL" in buscable else "Sin identificar"


def _extraer_folio(textos: list[str]) -> str:
    candidatos: list[tuple[str, int]] = []
    patrones_encabezado = [
        re.compile(r"(?:FACTURA|BOLETA|COMPROBANTE|TICKET|GUIA|FOLIO|DOC)\s*(?:ELECTRONIC[AO]|DE\s+VENTA)?\s*(?:N[°O\.:\s]+|NUMERO[:\s]+)?([0-9]{1,12})", re.I),
        re.compile(r"(?:N[°O\.]|NUM(?:ERO)?)\s*[:\s]*([0-9]{2,10})", re.I),
        re.compile(r"(?:OP|OPERACION|TRANSACCION|ORDEN)\s*[:\s#]*([0-9]{2,10})", re.I)
    ]

    for texto in textos:
        lineas = _lineas_utiles(_texto_busqueda(texto))
        for indice, linea in enumerate(lineas):
            for patron in patrones_encabezado:
                match = patron.search(linea)
                if match:
                    num = match.group(1).strip()
                    if num not in {"19", "2023", "2024", "2025", "2026", "2027"}:
                        candidatos.append((num, 40 + len(num)))
            
            if any(term in linea for term in ["ACTURA", "BOLETA", "COMPROBANTE"]):
                ventana = " ".join(lineas[indice:indice + 3])
                for numero in re.findall(r"(?<!\d)(\d{2,10})(?!\d)", ventana):
                    if numero not in {"19", "2023", "2024", "2025", "2026", "2027"}:
                        candidatos.append((numero, 15 + len(numero)))

    if not candidatos:
        return ""
    conteo = Counter(numero for numero, _ in candidatos)
    return max(candidatos, key=lambda item: (conteo[item[0]], item[1]))[0]


def _fecha_valida_o_corregida(dia: int, mes: int, anio: int) -> tuple[date | None, bool]:
    anio_maximo = date.today().year + 1
    candidatos_anio: list[tuple[int, int]] = []
    if 2000 <= anio <= anio_maximo:
        candidatos_anio.append((0, anio))
    else:
        texto_anio = f"{anio:04d}"
        for posible in range(2020, anio_maximo + 1):
            distancia = sum(a != b for a, b in zip(texto_anio, str(posible)))
            if distancia <= 1:
                candidatos_anio.append((distancia, posible))

    for distancia, posible in sorted(candidatos_anio, key=lambda item: (item[0], -item[1])):
        try:
            return date(posible, mes, dia), distancia > 0
        except ValueError:
            continue
    return None, False


def _extraer_fecha(textos: list[str]) -> tuple[str, bool]:
    candidatos: list[tuple[int, date, bool]] = []
    patron = re.compile(r"(?<!\d)(\d{1,2})\s*[-/.]\s*(\d{1,2})\s*[-/.]\s*(\d{2,4})(?!\d)")
    
    for texto in textos:
        buscable = _texto_busqueda(texto)
        for coincidencia in patron.finditer(buscable):
            d, m, a_raw = int(coincidencia.group(1)), int(coincidencia.group(2)), int(coincidencia.group(3))
            a = a_raw + 2000 if a_raw < 100 else a_raw
            
            fecha, corregida = _fecha_valida_o_corregida(d, m, a)
            if fecha is None:
                continue
            contexto = buscable[max(0, coincidencia.start() - 150):coincidencia.start()]
            puntaje = 10
            if re.search(r"FECHA|EMISION|HORA", contexto):
                puntaje += 60
            if "VENCIMIENTO" in contexto[-40:] or "VCTO" in contexto[-40:]:
                puntaje -= 30
            if corregida:
                puntaje -= 10
            candidatos.append((puntaje, fecha, corregida))

    if not candidatos:
        return "", False
    puntaje, fecha, corregida = max(candidatos, key=lambda item: (item[0], -item[1].toordinal()))
    return fecha.isoformat(), corregida


def _limpiar_razon_social(texto: str) -> str:
    texto = re.sub(r"^[\s:;|]+|[\s:;|]+$", "", texto)
    texto = re.sub(r"\s+", " ", texto)
    texto = re.sub(r"\bS\s*[.,]?\s*A\s*[.,]?\b", "S.A.", texto, flags=re.I)
    texto = re.sub(r"\bS\s*P\s*A\b", "SPA", texto, flags=re.I)
    texto = re.sub(r"\bL\s*T\s*D\s*A\b", "LTDA.", texto, flags=re.I)
    return texto.upper().strip(" .,-")


def _extraer_razon_social(textos: list[str]) -> str:
    candidatos: list[tuple[int, str]] = []
    patron_sociedad = re.compile(r"\b(?:S\.?A\.?|SPA|LTDA\.?|E\.?I\.?R\.?L\.?|SOCIEDAD|COMERCIAL|DISTRIBUIDORA|RETAIL)\b", re.I)

    for orden_texto, texto in enumerate(textos):
        lineas = _lineas_utiles(texto)
        normalizadas = [_texto_busqueda(linea) for linea in lineas]
        
        for indice, linea in enumerate(lineas[:10]):
            norm = normalizadas[indice]
            if any(ignorar in norm for ignorar in ["FACTURA", "BOLETA", "R.U.T", "RUT", "FECHA", "GIRO", "DIRECCION"]):
                continue
            
            if patron_sociedad.search(norm):
                razon = _limpiar_razon_social(linea)
                candidatos.append((80 - indice * 5 - orden_texto * 15, razon))
                break
        
        if not candidatos and len(lineas) >= 1:
            for linea in lineas[:3]:
                norm = _texto_busqueda(linea)
                if not any(ignorar in norm for ignorar in ["RUT", "FACTURA", "BOLETA", "VOUCHER", "BIENVENIDO"]):
                    if len(linea.strip()) > 3:
                        candidatos.append((30 - orden_texto * 10, _limpiar_razon_social(linea)))
                        break

    if not candidatos:
        return ""
    return max(candidatos, key=lambda item: item[0])[1]


def _parsear_monto(token: str) -> int | None:
    token = token.upper().replace("O", "0").replace("I", "1").replace("L", "1")
    token = re.sub(r"[$\s]", "", token)
    if not token:
        return None
    if "," in token and "." in token:
        if token.rfind(",") > token.rfind("."):
            token = token.replace(".", "").split(",", 1)[0]
        else:
            token = token.replace(",", "")
    elif "." in token:
        partes = token.split(".")
        if all(len(p) == 3 for p in partes[1:]):
            token = "".join(partes)
        elif len(partes[-1]) <= 2:
            token = partes[0]
    elif "," in token:
        partes = token.split(",")
        if all(len(p) == 3 for p in partes[1:]):
            token = "".join(partes)
        elif len(partes[-1]) <= 2:
            token = partes[0]
    
    token = re.sub(r"\D", "", token)
    if not token:
        return None
    valor = int(token)
    return valor if 0 <= valor <= 100_000_000_000 else None


def _montos_en_texto(texto: str) -> list[int]:
    patron = re.compile(r"(?:\$\s*)?[0-9OIL]{1,3}(?:[.,][0-9OIL]{3})+(?!\d)|(?:\$\s*)?[0-9OIL]{2,9}(?!\d)", re.I)
    valores: list[int] = []
    for token in patron.findall(texto):
        valor = _parsear_monto(token)
        if valor is not None and valor not in {19, 2024, 2025, 2026}:
            valores.append(valor)
    return valores


def _extraer_monto_etiquetado(textos: list[str], patrones: list[str]) -> int | None:
    candidatos: list[tuple[int, int]] = []
    compilados = [re.compile(patron, re.I) for patron in patrones]

    for orden, texto in enumerate(textos):
        lineas = _lineas_utiles(_texto_busqueda(texto))
        for indice, linea in enumerate(lineas):
            for patron in compilados:
                coincidencia = patron.search(linea)
                if not coincidencia:
                    continue
                
                # Buscar monto en la misma línea
                valores = _montos_en_texto(linea[coincidencia.end():])
                if valores:
                    candidatos.append((valores[-1], 30 - orden * 5))
                    break
                
                # Buscar monto en las 2 líneas siguientes
                for offset in range(1, 3):
                    if indice + offset < len(lineas):
                        valores_sig = _montos_en_texto(lineas[indice + offset])
                        if valores_sig:
                            candidatos.append((valores_sig[-1], 20 - offset * 4 - orden * 5))
                            break

    if not candidatos:
        return None
    conteo = Counter(val for val, _ in candidatos)
    return max(candidatos, key=lambda item: (item[1] + conteo[item[0]] * 5, item[0]))[0]


def _extraer_montos(textos: list[str]) -> dict[str, int | None]:
    return {
        "monto_neto": _extraer_monto_etiquetado(
            textos, [r"\bMONTO\s+NETO\b", r"\bTOTAL\s+NETO\b", r"\bNETO\b"]
        ),
        "monto_exento": _extraer_monto_etiquetado(
            textos, [r"\bMONTO\s+EXENTO\b", r"\bTOTAL\s+EXENTO\b", r"\bEXENTO\b"]
        ),
        "monto_iva": _extraer_monto_etiquetado(
            textos, [r"\bI\s*\.?\s*V\s*\.?\s*A\b", r"\bIVA\s*19\s*%\b"]
        ),
        "otros_impuestos": _extraer_monto_etiquetado(
            textos, [r"\bIMPUESTOS?\b", r"\bOTROS?\s+IMP\b", r"\bPROPINA\b"]
        ),
        "monto_total": _extraer_monto_etiquetado(
            textos, [
                r"\bTOTAL\s+A\s+PAGAR\b",
                r"\bMONTO\s+TOTAL\b",
                r"\bVALOR\s+TOTAL\b",
                r"\bTOTAL\s+VENTA\b",
                r"\bTOTAL\s+COMPRA\b",
                r"\bTOTAL\b",
            ]
        ),
    }


def _reconciliar_montos(montos: dict[str, int | None], tipo_doc: str) -> tuple[dict[str, int], list[str]]:
    notas: list[str] = []
    neto = montos.get("monto_neto")
    exento = montos.get("monto_exento") or 0
    iva = montos.get("monto_iva")
    otros = montos.get("otros_impuestos") or 0
    total = montos.get("monto_total")

    # Si es Boleta o Voucher y no hay desglose pero sí Total
    if (neto is None or neto == 0) and (iva is None or iva == 0) and total and total > 0:
        neto = round(total / 1.19)
        iva = total - neto
        notas.append("Neto e IVA desglosados automáticamente desde el total.")

    # Reconciliación con Neto
    if neto is not None and neto > 0:
        iva_esperado = round(neto * 0.19)
        tolerancia = max(2, round(iva_esperado * 0.01))
        
        if total is not None:
            diferencia = total - neto - exento - otros
            if diferencia >= 0 and abs(diferencia - iva_esperado) <= tolerancia:
                iva = diferencia
            elif iva is None or abs(iva - iva_esperado) > tolerancia:
                iva = iva_esperado
        elif iva is None:
            iva = iva_esperado

    if neto is not None and iva is not None:
        total_calculado = neto + exento + iva + otros
        if total is None or abs(total - total_calculado) > 2:
            total = total_calculado

    return {
        "monto_neto": int(neto or 0),
        "monto_exento": int(exento),
        "monto_iva": int(iva or 0),
        "otros_impuestos": int(otros),
        "monto_total": int(total or 0),
    }, notas


@lru_cache(maxsize=1)
def _idioma_tesseract() -> str:
    ejecutable = shutil.which("tesseract")
    if not ejecutable:
        raise RuntimeError("Tesseract OCR no está disponible en el entorno.")
    resultado = subprocess.run([ejecutable, "--list-langs"], capture_output=True, text=True, check=False)
    idiomas = set(resultado.stdout.splitlines())
    return "spa+eng" if "spa" in idiomas and "eng" in idiomas else "spa" if "spa" in idiomas else "eng"


def _ejecutar_ocr(imagen: Image.Image, psm: int) -> str:
    salida = io.BytesIO()
    imagen.save(salida, format="PNG")
    resultado = subprocess.run(
        [shutil.which("tesseract") or "tesseract", "stdin", "stdout", "-l", _idioma_tesseract(), "--psm", str(psm)],
        input=salida.getvalue(),
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        timeout=60,
        check=False,
    )
    if resultado.returncode != 0:
        return ""
    return resultado.stdout.decode("utf-8", errors="replace")


def _preparar_imagen(imagen: Image.Image) -> Image.Image:
    # Corrección de rotación EXIF y conversión a escala de grises
    imagen = ImageOps.exif_transpose(imagen).convert("L")
    
    # Escalado óptimo para Tesseract (DPI equivalente ~300)
    ancho_objetivo = 1600
    if imagen.width < ancho_objetivo:
        escala = ancho_objetivo / imagen.width
        imagen = imagen.resize((round(imagen.width * escala), round(imagen.height * escala)), Image.Resampling.LANCZOS)
    
    # Mejora de contraste para tickets térmicos desvanecidos
    imagen = ImageOps.autocontrast(imagen, cutoff=2)
    imagen = ImageEnhance.Contrast(imagen).enhance(1.4)
    return imagen.filter(ImageFilter.UnsharpMask(radius=1.5, percent=150, threshold=2))


def _obtener_textos_ocr(imagen: Image.Image) -> list[str]:
    preparada = _preparar_imagen(imagen)
    # Ejecuta únicamente 2 modos estratégicos en vez de 4 para optimizar tiempo
    return [
        _ejecutar_ocr(preparada, psm=6),  # Bloque uniforme de texto
        _ejecutar_ocr(preparada, psm=4),  # Columna única variable
    ]


def _extraer_registro(textos: list[str], archivo: str, pagina: int, uso_ocr: bool) -> dict[str, Any]:
    notas: list[str] = []
    tipo_doc = _tipo_documento(textos)
    rut, rut_corregido = _extraer_rut(textos)
    if rut_corregido:
        notas.append("RUT verificado y ajustado por Módulo 11.")
    
    fecha, fecha_corregida = _extraer_fecha(textos)
    if fecha_corregida:
        notas.append("Año ajustado por lectura OCR.")
    
    montos, notas_montos = _reconciliar_montos(_extraer_montos(textos), tipo_doc)
    notas.extend(notas_montos)
    
    folio = _extraer_folio(textos)
    razon_social = _extraer_razon_social(textos)

    faltantes: list[str] = []
    if not rut:
        faltantes.append("RUT")
    if not fecha:
        faltantes.append("Fecha")
    if not montos["monto_total"]:
        faltantes.append("Total")
    
    if faltantes:
        notas.append(f"Verificar: {', '.join(faltantes)}.")

    requiere_revision = bool(faltantes or rut_corregido or fecha_corregida)
    if uso_ocr and not requiere_revision and not notas:
        notas.append("Lectura OCR correcta.")

    resultado = {
        "archivo": archivo,
        "pagina": pagina,
        "tipo_documento": tipo_doc,
        "rut_emisor": rut,
        "razon_social": razon_social,
        "fecha": fecha,
        "folio": folio,
        **montos,
        "glosa": f"{tipo_doc} - {razon_social}" if razon_social else tipo_doc,
        "estado": "Requiere revisión" if requiere_revision else "Procesado",
        "observacion": " ".join(dict.fromkeys(notas)),
    }
    return {col: resultado.get(col, "") for col in COLUMNAS}


def procesar_con_motor_real(ruta_archivo: str) -> list[dict[str, Any]]:
    ruta = Path(ruta_archivo)
    extension = ruta.suffix.lower()
    registros: list[dict[str, Any]] = []

    if extension == ".pdf":
        with fitz.open(ruta) as documento:
            for indice, pagina_pdf in enumerate(documento):
                texto_nativo = pagina_pdf.get_text("text")
                uso_ocr = len(_texto_busqueda(texto_nativo).strip()) < 60
                textos = [texto_nativo] if texto_nativo.strip() else []
                
                if uso_ocr:
                    pixmap = pagina_pdf.get_pixmap(matrix=fitz.Matrix(2.5, 2.5), alpha=False)
                    imagen = Image.open(io.BytesIO(pixmap.tobytes("png")))
                    textos.extend(_obtener_textos_ocr(imagen))
                
                registro = _extraer_registro(textos, ruta.name, indice + 1, uso_ocr)
                registros.append(registro)
        return registros

    if extension in {".jpg", ".jpeg", ".png"}:
        with Image.open(ruta) as imagen:
            textos = _obtener_textos_ocr(imagen.copy())
        return [_extraer_registro(textos, ruta.name, 1, True)]

    raise ValueError(f"Formato no soportado: {extension}")


def procesar_archivos(archivos) -> pd.DataFrame:
    registros: list[dict[str, Any]] = []
    barra = st.progress(0, text="Iniciando extracción…")
    total_archivos = len(archivos)

    with tempfile.TemporaryDirectory(prefix="rendiciones_") as carpeta_temporal:
        carpeta = Path(carpeta_temporal)

        for indice, archivo in enumerate(archivos):
            barra.progress(
                (indice + 1) / total_archivos,
                text=f"Analizando {indice + 1} de {total_archivos}: {archivo.name}",
            )

            ruta = carpeta / Path(archivo.name).name
            ruta.write_bytes(archivo.getvalue())

            try:
                resultados = procesar_con_motor_real(str(ruta))
                registros.extend(resultados)
            except Exception as error:
                registros.append(
                    {
                        "archivo": archivo.name,
                        "pagina": 1,
                        "tipo_documento": "Sin identificar",
                        "rut_emisor": "",
                        "razon_social": "",
                        "fecha": "",
                        "folio": "",
                        "monto_neto": 0,
                        "monto_exento": 0,
                        "monto_iva": 0,
                        "otros_impuestos": 0,
                        "monto_total": 0,
                        "glosa": "",
                        "estado": "Error",
                        "observacion": f"Error: {str(error)}",
                    }
                )

    return pd.DataFrame(registros, columns=COLUMNAS)


def crear_excel(df: pd.DataFrame) -> bytes:
    salida = io.BytesIO()
    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Rendiciones")
        hoja = writer.sheets["Rendiciones"]
        hoja.freeze_panes = "A2"
        hoja.auto_filter.ref = hoja.dimensions

        relleno = PatternFill("solid", fgColor="1F4E79")
        fuente = Font(color="FFFFFF", bold=True)

        for celda in hoja[1]:
            celda.fill = relleno
            celda.font = fuente
            celda.alignment = Alignment(horizontal="center", vertical="center")

        for indice_col, nombre_col in enumerate(df.columns, start=1):
            valores = [str(nombre_col)] + [str(v) for v in df[nombre_col].fillna("").tolist()]
            hoja.column_dimensions[get_column_letter(indice_col)].width = min(max(len(v) for v in valores) + 3, 40)

        for col in ["monto_neto", "monto_exento", "monto_iva", "otros_impuestos", "monto_total"]:
            if col in df.columns:
                num_col = df.columns.get_loc(col) + 1
                for fila in range(2, len(df) + 2):
                    hoja.cell(fila, num_col).number_format = '$#,##0'

    salida.seek(0)
    return salida.getvalue()


def mostrar_vista_previa(archivos) -> None:
    with st.expander("Archivos cargados", expanded=False):
        for archivo in archivos:
            extension = Path(archivo.name).suffix.lower()
            c1, c2 = st.columns([1, 3])
            with c1:
                st.write(f"**{archivo.name}**")
                st.caption(f"{len(archivo.getvalue()) / 1024:.1f} KB")
            with c2:
                if extension in {".jpg", ".jpeg", ".png"}:
                    st.image(archivo.getvalue(), width=220)
                else:
                    st.caption("Documento PDF listo para procesar.")


def mostrar_metricas(df: pd.DataFrame) -> None:
    total = len(df)
    procesados = int((df["estado"] == "Procesado").sum()) if total else 0
    revision = int((df["estado"] == "Requiere revisión").sum()) if total else 0
    errores = int((df["estado"] == "Error").sum()) if total else 0
    monto_total = pd.to_numeric(df["monto_total"], errors="coerce").fillna(0).sum()

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total Documentos", total)
    c2.metric("Completados", procesados)
    c3.metric("Por Revisar", revision)
    c4.metric("Errores", errores)
    c5.metric("Monto Consolidado", f"${monto_total:,.0f}".replace(",", "."))


def mostrar_resultados() -> None:
    df = st.session_state.resultados
    if df.empty:
        st.info("Carga documentos en la pestaña 'Cargar y procesar' para comenzar.")
        return

    mostrar_metricas(df)
    st.markdown("### Planilla Consolidada")
    st.caption("Edita directamente cualquier celda antes de exportar.")

    config_cols = {
        "archivo": st.column_config.TextColumn("Archivo", disabled=True),
        "pagina": st.column_config.NumberColumn("Pág.", format="%d", disabled=True),
        "tipo_documento": st.column_config.SelectboxColumn(
            "Tipo", options=["Boleta", "Factura", "Voucher", "Sin identificar"]
        ),
        "rut_emisor": st.column_config.TextColumn("RUT Emisor"),
        "razon_social": st.column_config.TextColumn("Razón Social"),
        "fecha": st.column_config.TextColumn("Fecha (AAAA-MM-DD)"),
        "folio": st.column_config.TextColumn("Folio / N° OP"),
        "monto_neto": st.column_config.NumberColumn("Neto", format="$ %d"),
        "monto_exento": st.column_config.NumberColumn("Exento", format="$ %d"),
        "monto_iva": st.column_config.NumberColumn("IVA", format="$ %d"),
        "otros_impuestos": st.column_config.NumberColumn("Otros Imp.", format="$ %d"),
        "monto_total": st.column_config.NumberColumn("Total", format="$ %d"),
        "glosa": st.column_config.TextColumn("Glosa"),
        "estado": st.column_config.SelectboxColumn("Estado", options=["Procesado", "Requiere revisión", "Error"]),
        "observacion": st.column_config.TextColumn("Observaciones"),
    }

    df_editado = st.data_editor(
        df,
        use_container_width=True,
        hide_index=True,
        column_config=config_cols,
        key="editor_resultados",
    )
    st.session_state.resultados = df_editado

    excel_data = crear_excel(df_editado)
    csv_data = df_editado.to_csv(index=False).encode("utf-8-sig")

    c_excel, c_csv, c_clean = st.columns([1, 1, 2])
    with c_excel:
        st.download_button(
            "📥 Descargar Excel (.xlsx)",
            data=excel_data,
            file_name="rendicion_gastos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with c_csv:
        st.download_button(
            "📥 Descargar CSV",
            data=csv_data,
            file_name="rendicion_gastos.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with c_clean:
        if st.button("🗑️ Limpiar Todo", type="secondary"):
            st.session_state.resultados = pd.DataFrame(columns=COLUMNAS)
            st.session_state.procesamiento_realizado = False
            st.rerun()


def main() -> None:
    configurar_pagina()
    inicializar_estado()

    st.markdown(
        """
        <div class="hero">
            <h1>🧾 Rendición Inteligente v0.3.0</h1>
            <p>Automatización de extracción para Facturas, Boletas y Vouchers chilenos a Excel.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    tab_procesar, tab_resultados = st.tabs(["1. Cargar y procesar", "2. Resultados"])

    with tab_procesar:
        st.subheader("Cargar Documentos Tributarios")
        archivos = st.file_uploader(
            "Arrastra archivos PDF, JPG o PNG",
            type=["pdf", "jpg", "jpeg", "png"],
            accept_multiple_files=True,
        )

        if archivos:
            st.success(f"{len(archivos)} archivo(s) listo(s) para procesar.")
            mostrar_vista_previa(archivos)

            if st.button("🚀 Procesar Documentos", type="primary", use_container_width=True):
                with st.spinner("Procesando y cuadrando montos..."):
                    st.session_state.resultados = procesar_archivos(archivos)
                    st.session_state.procesamiento_realizado = True
                st.success("¡Completado! Revisa la pestaña 'Resultados'.")

    with tab_resultados:
        mostrar_resultados()


if __name__ == "__main__":
    main()
