"""
Microservizio per la compilazione del template pptx (Valutazione Preliminare).
Gestisce:
  1. Sostituzione placeholder di testo (WWW.DOMINIOCLIENTE.IT, {{TOKEN}})
  2. Colore dinamico dei 4 cerchi punteggio PageSpeed (slide 3), individuati
     tramite il campo "descr" (testo alternativo impostato in Google Slides)

Endpoint: POST /compila
Form-data:
  - file: il file .pptx (binario, scaricato da Drive già convertito da Slides)
  - dati: stringa JSON con la struttura prodotta dal nodo n8n "Componi payload finale"
Risposta: il file .pptx compilato (binario)
"""

import io
import json
import re
import shutil
import tempfile

from fastapi import FastAPI, File, Form, UploadFile, HTTPException
from fastapi.responses import StreamingResponse
from pptx import Presentation
from pptx.dml.color import RGBColor

app = FastAPI(title="MP Quadro - Compilazione template Valutazione Preliminare")

NS_P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"

COLORE_ROSSO = RGBColor(0xEA, 0x43, 0x35)
COLORE_GIALLO = RGBColor(0xF4, 0xB4, 0x00)
COLORE_VERDE = RGBColor(0x34, 0xA8, 0x53)
COLORE_GRIGIO = RGBColor(0x9E, 0x9E, 0x9E)  # dato mancante/errore


def colore_per_punteggio(punteggio):
    if punteggio is None:
        return COLORE_GRIGIO
    if punteggio >= 90:
        return COLORE_VERDE
    if punteggio >= 50:
        return COLORE_GIALLO
    return COLORE_ROSSO


def colore_per_secondi(secondi):
    if secondi is None:
        return COLORE_GRIGIO
    if secondi <= 1.5:
        return COLORE_VERDE
    if secondi <= 3:
        return COLORE_GIALLO
    return COLORE_ROSSO


def colora_testo_livello(text_frame):
    """Colora il run il cui testo finale è esattamente OTTIMO/BUONO/PESSIMO/N/D."""
    mappa = {"OTTIMO": COLORE_VERDE, "BUONO": COLORE_GIALLO, "PESSIMO": COLORE_ROSSO, "N/D": COLORE_GRIGIO}
    n = 0
    for paragraph in text_frame.paragraphs:
        for run in paragraph.runs:
            if run.text.strip() in mappa:
                run.font.color.rgb = mappa[run.text.strip()]
                n += 1
    return n


def livello_per_tempo(secondi):
    if secondi is None:
        return "N/D"
    if secondi <= 1.5:
        return "OTTIMO"
    if secondi <= 3:
        return "BUONO"
    return "PESSIMO"


def costruisci_sostituzioni(dati):
    """Costruisce il dizionario {placeholder: valore} a partire dal payload n8n."""
    pagespeed = {(r["label"], r["strategy"]): r for r in dati.get("pagespeed", [])}

    def val(label, strategy, campo, default="N/D"):
        r = pagespeed.get((label, strategy))
        if r is None or r.get("errore"):
            return default
        v = r.get(campo)
        return default if v is None else str(v)

    cliente_desktop = pagespeed.get(("cliente", "desktop"))
    competitor_desktop = pagespeed.get(("competitor", "desktop"))

    sostituzioni = {
        "PSI_MOBILE_CLIENTE": val("cliente", "mobile", "punteggio"),
        "PSI_DESKTOP_CLIENTE": val("cliente", "desktop", "punteggio"),
        "PSI_MOBILE_COMPETITOR": val("competitor", "mobile", "punteggio"),
        "PSI_DESKTOP_COMPETITOR": val("competitor", "desktop", "punteggio"),
        "GRADO_CLIENTE": val("cliente", "desktop", "grado"),
        "GRADO_COMPETITOR": val("competitor", "desktop", "grado"),
        "PAGESIZE_CLIENTE": val("cliente", "desktop", "pageSizeMB"),
        "PAGESIZE_COMPETITOR": val("competitor", "desktop", "pageSizeMB"),
        "LOADTIME_CLIENTE": val("cliente", "desktop", "loadTimeSec"),
        "LOADTIME_COMPETITOR": val("competitor", "desktop", "loadTimeSec"),
        "REQUESTS_CLIENTE": val("cliente", "desktop", "numeroRichieste"),
        "REQUESTS_COMPETITOR": val("competitor", "desktop", "numeroRichieste"),
        "LIVELLO_CLIENTE": livello_per_tempo(
            cliente_desktop.get("loadTimeSec") if cliente_desktop and not cliente_desktop.get("errore") else None
        ),
        "LIVELLO_COMPETITOR": livello_per_tempo(
            competitor_desktop.get("loadTimeSec") if competitor_desktop and not competitor_desktop.get("errore") else None
        ),
    }

    return sostituzioni, pagespeed


def sostituisci_in_run(run, mappa):
    cambiato = False
    testo = run.text

    for placeholder, valore in [
        ("WWW.DOMINIOCLIENTE.IT", mappa.get("_nomeCliente")),
        ("WWW.DOMINIOCOMPETITOR.IT", mappa.get("_nomeCompetitor")),
    ]:
        if valore and placeholder.lower() in testo.lower():
            pattern = re.compile(re.escape(placeholder), re.IGNORECASE)
            testo = pattern.sub(valore, testo)
            cambiato = True

    for token, valore in mappa.items():
        if token.startswith("_"):
            continue
        marcatore = "{{" + token + "}}"
        if marcatore in testo:
            testo = testo.replace(marcatore, valore)
            cambiato = True

    if cambiato:
        run.text = testo
    return cambiato


def processa_text_frame(text_frame, mappa):
    n = 0
    for paragraph in text_frame.paragraphs:
        for run in paragraph.runs:
            if sostituisci_in_run(run, mappa):
                n += 1
    return n


def processa_tabella(table, mappa):
    n = 0
    for row in table.rows:
        for cell in row.cells:
            n += processa_text_frame(cell.text_frame, mappa)
    return n


def trova_descr(shape):
    el = shape._element.find(f".//{NS_P}cNvPr")
    return el.get("descr") if el is not None else None


def colora_cerchi(slide, pagespeed):
    mappa_forme = {
        "circle_psi_mobile_cliente": ("cliente", "mobile"),
        "circle_psi_desktop_cliente": ("cliente", "desktop"),
        "circle_psi_mobile_competitor": ("competitor", "mobile"),
        "circle_psi_desktop_competitor": ("competitor", "desktop"),
    }
    mappa_box_loadtime = {
        "loadtime_box_cliente": "cliente",
        "loadtime_box_competitor": "competitor",
    }
    colorati = 0

    def ricorsione(shapes):
        nonlocal colorati
        for shape in shapes:
            if shape.shape_type == 6:  # GROUP
                ricorsione(shape.shapes)
                continue
            d = trova_descr(shape)
            if d in mappa_forme:
                label, strategy = mappa_forme[d]
                r = pagespeed.get((label, strategy))
                punteggio = r.get("punteggio") if r and not r.get("errore") else None
                colore = colore_per_punteggio(punteggio)
                try:
                    shape.fill.solid()
                    shape.fill.fore_color.rgb = colore
                    colorati += 1
                except Exception:
                    pass
            elif d in mappa_box_loadtime:
                label = mappa_box_loadtime[d]
                r = pagespeed.get((label, "desktop"))
                secondi = r.get("loadTimeSec") if r and not r.get("errore") else None
                colore = colore_per_secondi(secondi)
                try:
                    shape.line.color.rgb = colore
                    colorati += 1
                except Exception:
                    pass

    ricorsione(slide.shapes)
    return colorati


@app.post("/compila")
async def compila(
    file: UploadFile = File(...),
    dati: str = Form(...),
):
    mime_pptx = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    nome_ok = (file.filename or "").endswith(".pptx")
    mime_ok = file.content_type == mime_pptx
    if not nome_ok and not mime_ok:
        raise HTTPException(
            400,
            f"Il file deve essere .pptx (ricevuto filename={file.filename!r}, "
            f"content_type={file.content_type!r})",
        )

    try:
        payload = json.loads(dati)
    except json.JSONDecodeError:
        raise HTTPException(400, "Campo 'dati' non è JSON valido")

    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
        shutil.copyfileobj(file.file, tmp)
        tmp_path = tmp.name

    prs = Presentation(tmp_path)

    sostituzioni, pagespeed = costruisci_sostituzioni(payload)
    sostituzioni["_nomeCliente"] = payload.get("nomeCliente")
    sostituzioni["_nomeCompetitor"] = payload.get("nomeCompetitor")

    totale_testo = 0
    totale_colori = 0
    totale_livello = 0
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                totale_testo += processa_text_frame(shape.text_frame, sostituzioni)
                totale_livello += colora_testo_livello(shape.text_frame)
            if shape.has_table:
                totale_testo += processa_tabella(shape.table, sostituzioni)
        totale_colori += colora_cerchi(slide, pagespeed)

    if totale_testo == 0:
        raise HTTPException(
            422,
            "Nessuna sostituzione di testo effettuata. Verifica che il template "
            "contenga ancora i placeholder attesi.",
        )

    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)

    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={
            "Content-Disposition": 'attachment; filename="output.pptx"',
            "X-Sostituzioni-Testo": str(totale_testo),
            "X-Cerchi-Colorati": str(totale_colori),
        },
    )


@app.get("/health")
async def health():
    return {"status": "ok"}
