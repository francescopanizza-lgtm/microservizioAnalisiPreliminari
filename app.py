"""
Microservizio per la compilazione del template pptx (Valutazione Preliminare).
Gestisce:
  1. Sostituzione placeholder di testo (WWW.DOMINIOCLIENTE.IT, {{TOKEN}})
  2. Colore dinamico dei 4 cerchi punteggio PageSpeed (slide 3) e dei bordi
     Load time (slide 4), individuati tramite il campo "descr" (testo
     alternativo impostato in Google Slides)
  3. Analisi on-page gratuita (title/description mancanti, troppo corti/
     lunghi, duplicati) come sostituto minimale di Screaming Frog

Endpoint:
  POST /compila         -> compila il template (invariato)
  POST /analizza-onpage -> crawla un sito e restituisce errori title/description
  GET  /health
"""

import io
import json
import re
import shutil
import tempfile
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup
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


# ---------------------------------------------------------------------------
# Analisi on-page gratuita (sostituto minimale di Screaming Frog)
# ---------------------------------------------------------------------------

USER_AGENT = "Mozilla/5.0 (compatible; MPQuadroValutazioneBot/1.0)"
TIMEOUT_SEC = 10
TITLE_MIN, TITLE_MAX = 30, 60
DESC_MIN, DESC_MAX = 70, 155


def stesso_dominio(url, dominio_base):
    try:
        return urlparse(url).netloc.lower().lstrip("www.") == dominio_base.lower().lstrip("www.")
    except Exception:
        return False


def scopri_pagine(base_url, max_pagine=25):
    """Prova prima la sitemap.xml; se assente/vuota, fa un crawl BFS semplice."""
    dominio = urlparse(base_url).netloc
    pagine = []

    # Tentativo 1: sitemap.xml
    try:
        r = requests.get(urljoin(base_url, "/sitemap.xml"), headers={"User-Agent": USER_AGENT}, timeout=TIMEOUT_SEC)
        if r.status_code == 200 and "xml" in r.headers.get("Content-Type", ""):
            soup = BeautifulSoup(r.content, "xml")
            locs = [loc.text.strip() for loc in soup.find_all("loc")]
            pagine = [u for u in locs if stesso_dominio(u, dominio)][:max_pagine]
    except Exception:
        pass

    if pagine:
        return pagine

    # Tentativo 2: crawl BFS semplice partendo dalla homepage
    visitate = set()
    da_visitare = [base_url]
    while da_visitare and len(visitate) < max_pagine:
        url = da_visitare.pop(0)
        if url in visitate:
            continue
        visitate.add(url)
        try:
            r = requests.get(url, headers={"User-Agent": USER_AGENT}, timeout=TIMEOUT_SEC)
            if r.status_code != 200:
                continue
            soup = BeautifulSoup(r.content, "html.parser")
            for a in soup.find_all("a", href=True):
                link = urljoin(url, a["href"]).split("#")[0]
                if stesso_dominio(link, dominio) and link not in visitate and link not in da_visitare:
                    da_visitare.append(link)
        except Exception:
            continue

    return list(visitate)[:max_pagine]


def analizza_pagina(url):
    try:
        r = requests.get(url, headers={"User-Agent": USER_AGENT}, timeout=TIMEOUT_SEC, allow_redirects=True)
    except Exception as e:
        return {"url": url, "status": None, "errore_richiesta": str(e), "title": None, "description": None}

    risultato = {"url": url, "status": r.status_code, "errore_richiesta": None}

    if r.status_code >= 400:
        risultato["title"] = None
        risultato["description"] = None
        return risultato

    soup = BeautifulSoup(r.content, "html.parser")
    title_tag = soup.find("title")
    desc_tag = soup.find("meta", attrs={"name": "description"})

    risultato["title"] = title_tag.text.strip() if title_tag and title_tag.text else None
    risultato["description"] = desc_tag.get("content", "").strip() if desc_tag and desc_tag.get("content") else None

    return risultato


def valuta_title(title):
    if not title:
        return "mancante"
    if len(title) < TITLE_MIN:
        return "troppo corto"
    if len(title) > TITLE_MAX:
        return "troppo lungo"
    return None


def valuta_description(desc):
    if not desc:
        return "mancante"
    if len(desc) < DESC_MIN:
        return "troppo corta"
    if len(desc) > DESC_MAX:
        return "troppo lunga"
    return None


@app.post("/analizza-onpage")
async def analizza_onpage(url: str = Form(...), max_pagine: int = Form(25)):
    if not url.startswith(("http://", "https://")):
        url = "https://" + url

    pagine = scopri_pagine(url, max_pagine)
    if not pagine:
        raise HTTPException(422, "Nessuna pagina raggiungibile per questo dominio")

    risultati = [analizza_pagina(p) for p in pagine]

    titoli = [r["title"] for r in risultati if r["title"]]
    descrizioni = [r["description"] for r in risultati if r["description"]]
    duplicati_titoli = {t for t in titoli if titoli.count(t) > 1}
    duplicati_descrizioni = {d for d in descrizioni if descrizioni.count(d) > 1}

    pagine_con_errore_title = []
    pagine_con_errore_description = []
    pagine_con_errore_status = []

    for r in risultati:
        if r["status"] is None or r["status"] >= 400:
            pagine_con_errore_status.append({"url": r["url"], "status": r["status"], "dettaglio": r.get("errore_richiesta")})
            continue
        problema_title = valuta_title(r["title"])
        if not problema_title and r["title"] in duplicati_titoli:
            problema_title = "duplicato"
        if problema_title:
            pagine_con_errore_title.append({"url": r["url"], "problema": problema_title, "valore": r["title"]})

        problema_desc = valuta_description(r["description"])
        if not problema_desc and r["description"] in duplicati_descrizioni:
            problema_desc = "duplicata"
        if problema_desc:
            pagine_con_errore_description.append({"url": r["url"], "problema": problema_desc, "valore": r["description"]})

    return {
        "dominio": url,
        "pagine_analizzate": len(risultati),
        "conteggio_errori_title": len(pagine_con_errore_title),
        "conteggio_errori_description": len(pagine_con_errore_description),
        "conteggio_errori_status": len(pagine_con_errore_status),
        "dettaglio_errori_title": pagine_con_errore_title,
        "dettaglio_errori_description": pagine_con_errore_description,
        "dettaglio_errori_status": pagine_con_errore_status,
        "nota": (
            "Analisi indicativa: copertura limitata a max_pagine, non esegue JavaScript "
            "(possibili falsi positivi su siti che generano title/description via JS), "
            "soglie title 30-60 caratteri e description 70-155 caratteri."
        ),
    }
