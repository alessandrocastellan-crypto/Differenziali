
import io
import json
import re
import zipfile
from pathlib import Path
from datetime import datetime

import pandas as pd
import streamlit as st

try:
    import xlrd
except Exception:
    xlrd = None

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
except Exception:
    colors = None


APP_TITLE = "Gestionale schede differenziali"


def normalize_text(value):
    if value is None:
        return ""
    return str(value).replace("\n", " ").strip()


def clean_label(value):
    return normalize_text(value).lower().replace(":", "").strip()


def read_xls_bytes(file_bytes, filename):
    """
    Legge .xls / .xlsx in modo semplice.
    Per .xls usa xlrd, per .xlsx usa pandas/openpyxl.
    Restituisce una lista di fogli, ognuno come lista di liste.
    """
    suffix = Path(filename).suffix.lower()
    sheets = []

    if suffix == ".xls":
        if xlrd is None:
            raise RuntimeError("xlrd non installato")
        book = xlrd.open_workbook(file_contents=file_bytes)
        for sh in book.sheets():
            rows = []
            for r in range(sh.nrows):
                rows.append([sh.cell_value(r, c) for c in range(sh.ncols)])
            sheets.append(rows)
    elif suffix == ".xlsx":
        xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None, dtype=str)
            sheets.append(df.fillna("").values.tolist())
    return sheets


def find_value_after_labels(rows, labels):
    """
    Cerca una cella con etichetta tipo 'Nome quadro' e restituisce
    il primo valore utile nelle celle a destra o immediatamente sotto.
    """
    labels = [l.lower() for l in labels]
    for r, row in enumerate(rows):
        for c, cell in enumerate(row):
            txt = clean_label(cell)
            if any(lbl in txt for lbl in labels):
                # prima celle a destra
                for cc in range(c + 1, min(c + 8, len(row))):
                    val = normalize_text(row[cc])
                    if val and clean_label(val) not in labels:
                        return val
                # poi riga sotto stessa colonna/destra
                for rr in range(r + 1, min(r + 4, len(rows))):
                    for cc in range(c, min(c + 8, len(rows[rr]))):
                        val = normalize_text(rows[rr][cc])
                        if val and clean_label(val) not in labels:
                            return val
    return ""


def row_contains_any(row, words):
    joined = " ".join(clean_label(x) for x in row)
    return any(w.lower() in joined for w in words)


def extract_switch_rows(rows):
    """
    Estrae righe interruttori cercando intestazioni compatibili.
    Colonne attese: n° interruttore, circuito, tipo differenziale, dati nominali.
    È una funzione volutamente tollerante perché le schede possono variare.
    """
    header_idx = None
    colmap = {}

    for i, row in enumerate(rows):
        lowered = [clean_label(x) for x in row]
        joined = " ".join(lowered)
        if ("circuit" in joined or "utenza" in joined) and ("differenzial" in joined or "idn" in joined or "nominal" in joined):
            header_idx = i
            for c, txt in enumerate(lowered):
                if re.search(r'\bn\b|n°|numero|interrutt', txt):
                    colmap["numero"] = c
                if "circuit" in txt or "utenza" in txt or "descrizione" in txt:
                    colmap["circuito"] = c
                if "tipo" in txt and ("diff" in txt or "interrutt" in txt):
                    colmap["tipoDifferenziale"] = c
                if "dati" in txt or "nominal" in txt or "in " in txt or "idn" in txt or "sensibil" in txt:
                    colmap["datiNominali"] = c
            break

    extracted = []

    if header_idx is not None:
        for row in rows[header_idx + 1:]:
            vals = [normalize_text(x) for x in row]
            if not any(vals):
                continue
            joined = " ".join(vals).lower()
            if "note" in joined or "firma" in joined or "manutentore" in joined:
                break

            numero = vals[colmap.get("numero", 0)] if colmap.get("numero", 0) < len(vals) else ""
            circuito = vals[colmap.get("circuito", 1)] if colmap.get("circuito", 1) < len(vals) else ""
            tipo = vals[colmap.get("tipoDifferenziale", 2)] if colmap.get("tipoDifferenziale", 2) < len(vals) else ""
            dati = vals[colmap.get("datiNominali", 3)] if colmap.get("datiNominali", 3) < len(vals) else ""

            if circuito or tipo or dati:
                extracted.append({
                    "numero": numero,
                    "circuito": circuito,
                    "tipoDifferenziale": tipo,
                    "datiNominali": dati
                })

    # fallback: cerca righe dove compaiono valori tipici dei differenziali
    if not extracted:
        for row in rows:
            vals = [normalize_text(x) for x in row]
            joined = " ".join(vals).lower()
            if ("0,03" in joined or "0.03" in joined or "30ma" in joined or "diff" in joined) and len([v for v in vals if v]) >= 2:
                nonempty = [v for v in vals if v]
                extracted.append({
                    "numero": nonempty[0] if len(nonempty) > 0 else "",
                    "circuito": nonempty[1] if len(nonempty) > 1 else "",
                    "tipoDifferenziale": nonempty[2] if len(nonempty) > 2 else "",
                    "datiNominali": " ".join(nonempty[3:]) if len(nonempty) > 3 else ""
                })

    return extracted


def extract_sheet_from_file(file_bytes, filename):
    sheets = read_xls_bytes(file_bytes, filename)
    all_rows = []
    for s in sheets:
        all_rows.extend(s)
        all_rows.append([])

    blocco = find_value_after_labels(all_rows, ["blocco"])
    piano = find_value_after_labels(all_rows, ["piano"])
    nome = find_value_after_labels(all_rows, ["nome quadro", "quadro"])
    reparto = find_value_after_labels(all_rows, ["reparto", "utenza"])

    switches = extract_switch_rows(all_rows)

    if not nome:
        nome = Path(filename).stem

    return {
        "id": f"{Path(filename).stem}-{abs(hash(filename))}",
        "fileOrigine": filename,
        "blocco": blocco,
        "piano": piano,
        "nomeQuadro": nome,
        "reparto": reparto,
        "note": "",
        "interruttori": switches or [{"numero": "", "circuito": "", "tipoDifferenziale": "", "datiNominali": ""}]
    }


def import_zip(uploaded_file):
    data = []
    errors = []

    with zipfile.ZipFile(uploaded_file) as z:
        names = [n for n in z.namelist() if n.lower().endswith((".xls", ".xlsx")) and not Path(n).name.startswith("~$")]
        for name in names:
            try:
                file_bytes = z.read(name)
                sheet = extract_sheet_from_file(file_bytes, Path(name).name)
                data.append(sheet)
            except Exception as e:
                errors.append({"file": name, "errore": str(e)})

    # evita duplicati per Nome Quadro + Piano + Blocco
    seen = set()
    unique = []
    for s in data:
        key = (s.get("nomeQuadro","").strip().lower(), s.get("piano","").strip().lower(), s.get("blocco","").strip().lower())
        if key not in seen:
            seen.add(key)
            unique.append(s)

    unique.sort(key=lambda x: (x.get("blocco",""), x.get("nomeQuadro","")))
    return unique, errors


def flatten_data(sheets):
    rows = []
    for s in sheets:
        for sw in s.get("interruttori", []):
            rows.append({
                "Blocco": s.get("blocco", ""),
                "Piano": s.get("piano", ""),
                "Nome Quadro": s.get("nomeQuadro", ""),
                "Reparto": s.get("reparto", ""),
                "N° interruttore": sw.get("numero", ""),
                "Circuito": sw.get("circuito", ""),
                "Tipo differenziale": sw.get("tipoDifferenziale", ""),
                "Dati nominali": sw.get("datiNominali", ""),
                "File origine": s.get("fileOrigine", "")
            })
    return pd.DataFrame(rows)


def make_excel(sheets):
    df = flatten_data(sheets)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Tutti i dati", index=False)
        for blocco, group in df.groupby(df["Blocco"].fillna("")):
            name = str(blocco)[:25] if str(blocco).strip() else "Senza blocco"
            safe = re.sub(r'[\[\]\:\*\?\/\\]', "_", name)
            group.to_excel(writer, sheet_name=safe[:31], index=False)
    output.seek(0)
    return output.getvalue()


def make_pdf(sheets):
    if colors is None:
        raise RuntimeError("reportlab non installato")

    output = io.BytesIO()
    doc = SimpleDocTemplate(
        output,
        pagesize=A4,
        rightMargin=10*mm,
        leftMargin=10*mm,
        topMargin=8*mm,
        bottomMargin=8*mm
    )

    styles = getSampleStyleSheet()
    small = ParagraphStyle("small", parent=styles["Normal"], fontSize=7, leading=8)
    normal = ParagraphStyle("normal", parent=styles["Normal"], fontSize=8, leading=9)
    title = ParagraphStyle("title", parent=styles["Normal"], fontSize=12, leading=14, alignment=1, spaceAfter=4)

    story = []

    for idx, s in enumerate(sorted(sheets, key=lambda x: (x.get("blocco",""), x.get("nomeQuadro","")))):
        story.append(Table([
            ["LOGO / ENTE", Paragraph("<b>SCHEDA VERIFICA INTERRUTTORI DIFFERENZIALI</b>", title), "Rev. 00"]
        ], colWidths=[35*mm, 120*mm, 35*mm], style=[
            ("GRID", (0,0), (-1,-1), 1, colors.black),
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("ALIGN", (0,0), (-1,-1), "CENTER"),
            ("FONTNAME", (0,0), (-1,-1), "Helvetica-Bold"),
        ]))

        story.append(Table([
            [Paragraph(f"<b>Blocco:</b> {s.get('blocco','')}", normal),
             Paragraph(f"<b>Piano:</b> {s.get('piano','')}", normal),
             Paragraph(f"<b>Nome quadro:</b> {s.get('nomeQuadro','')}", normal),
             Paragraph(f"<b>Reparto:</b> {s.get('reparto','')}", normal)]
        ], colWidths=[35*mm, 30*mm, 65*mm, 60*mm], style=[
            ("GRID", (0,0), (-1,-1), 0.8, colors.black),
            ("VALIGN", (0,0), (-1,-1), "TOP"),
        ]))

        table_data = [[
            Paragraph("<b>N° INT.</b>", small),
            Paragraph("<b>CIRCUITO</b>", small),
            Paragraph("<b>TIPO DIFFERENZIALE</b>", small),
            Paragraph("<b>DATI NOMINALI</b>", small)
        ]]

        switches = s.get("interruttori", [])
        min_rows = max(18, len(switches))
        for r in range(min_rows):
            sw = switches[r] if r < len(switches) else {}
            table_data.append([
                Paragraph(str(sw.get("numero","")), small),
                Paragraph(str(sw.get("circuito","")), small),
                Paragraph(str(sw.get("tipoDifferenziale","")), small),
                Paragraph(str(sw.get("datiNominali","")), small),
            ])

        story.append(Table(table_data, colWidths=[18*mm, 82*mm, 45*mm, 45*mm], repeatRows=1, style=[
            ("GRID", (0,0), (-1,-1), 0.6, colors.black),
            ("VALIGN", (0,0), (-1,-1), "TOP"),
            ("ALIGN", (0,0), (0,-1), "CENTER"),
            ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
            ("FONTSIZE", (0,0), (-1,-1), 7),
            ("ROWHEIGHT", (0,1), (-1,-1), 8*mm),
        ]))

        story.append(Table([
            [Paragraph(f"<b>Note:</b> {s.get('note','')}", normal),
             Paragraph("<b>Data:</b>", normal),
             Paragraph("<b>Firma:</b>", normal)]
        ], colWidths=[95*mm, 45*mm, 50*mm], style=[
            ("GRID", (0,0), (-1,-1), 0.8, colors.black),
            ("VALIGN", (0,0), (-1,-1), "TOP"),
            ("MINROWHEIGHT", (0,0), (-1,-1), 18*mm),
        ]))

        if idx < len(sheets) - 1:
            story.append(PageBreak())

    doc.build(story)
    output.seek(0)
    return output.getvalue()


def ensure_state():
    if "sheets" not in st.session_state:
        st.session_state.sheets = []
    if "selected_index" not in st.session_state:
        st.session_state.selected_index = 0


def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    ensure_state()

    st.title(APP_TITLE)
    st.caption("Importa ZIP di schede, modifica i dati, esporta Excel e genera PDF stampabile.")

    with st.sidebar:
        uploaded = st.file_uploader("Carica ZIP schede originali", type=["zip"])
        if uploaded is not None:
            if st.button("Importa ZIP", type="primary"):
                sheets, errors = import_zip(uploaded)
                st.session_state.sheets = sheets
                st.session_state.selected_index = 0
                st.success(f"Importate {len(sheets)} schede.")
                if errors:
                    st.warning(f"{len(errors)} file non letti.")
                    st.json(errors)

        st.divider()

        if st.button("Nuova scheda"):
            st.session_state.sheets.append({
                "id": f"manuale-{datetime.now().timestamp()}",
                "fileOrigine": "manuale",
                "blocco": "",
                "piano": "",
                "nomeQuadro": "",
                "reparto": "",
                "note": "",
                "interruttori": [{"numero": "", "circuito": "", "tipoDifferenziale": "", "datiNominali": ""}]
            })
            st.session_state.selected_index = len(st.session_state.sheets) - 1

        json_upload = st.file_uploader("Importa database JSON", type=["json"])
        if json_upload is not None:
            try:
                st.session_state.sheets = json.load(json_upload)
                st.session_state.selected_index = 0
                st.success("Database JSON importato.")
            except Exception:
                st.error("JSON non valido.")

        if st.session_state.sheets:
            st.download_button(
                "Scarica database JSON",
                data=json.dumps(st.session_state.sheets, indent=2, ensure_ascii=False).encode("utf-8"),
                file_name="database_schede_differenziali.json",
                mime="application/json"
            )

    sheets = st.session_state.sheets

    if not sheets:
        st.info("Carica uno ZIP oppure crea una nuova scheda dalla barra laterale.")
        return

    col_left, col_right = st.columns([1, 2])

    with col_left:
        q = st.text_input("Cerca quadro / reparto / blocco")
        options = []
        index_map = []
        for i, s in enumerate(sheets):
            label = f"{s.get('nomeQuadro','Senza nome')} | Blocco {s.get('blocco','-')} | Piano {s.get('piano','-')} | {s.get('reparto','')}"
            if not q or q.lower() in label.lower():
                options.append(label)
                index_map.append(i)

        if options:
            current_label_index = 0
            if st.session_state.selected_index in index_map:
                current_label_index = index_map.index(st.session_state.selected_index)
            chosen = st.selectbox("Scheda", options, index=current_label_index)
            st.session_state.selected_index = index_map[options.index(chosen)]

        if st.button("Elimina scheda selezionata"):
            if sheets:
                sheets.pop(st.session_state.selected_index)
                st.session_state.selected_index = max(0, min(st.session_state.selected_index, len(sheets)-1))
                st.rerun()

        st.divider()
        st.subheader("Esporta")
        excel_bytes = make_excel(sheets)
        st.download_button("Scarica Excel dati", excel_bytes, "schede_differenziali.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        try:
            pdf_bytes = make_pdf(sheets)
            st.download_button("Scarica PDF stampabile", pdf_bytes, "schede_differenziali_stampabili.pdf", "application/pdf")
        except Exception as e:
            st.error(f"PDF non disponibile: {e}")

    with col_right:
        s = sheets[st.session_state.selected_index]
        st.subheader("Modifica scheda")

        c1, c2, c3, c4 = st.columns(4)
        s["blocco"] = c1.text_input("Blocco", value=s.get("blocco",""))
        s["piano"] = c2.text_input("Piano", value=s.get("piano",""))
        s["nomeQuadro"] = c3.text_input("Nome quadro", value=s.get("nomeQuadro",""))
        s["reparto"] = c4.text_input("Reparto", value=s.get("reparto",""))
        s["note"] = st.text_area("Note", value=s.get("note",""), height=80)

        st.markdown("### Interruttori")
        df = pd.DataFrame(s.get("interruttori", []))
        if df.empty:
            df = pd.DataFrame([{"numero": "", "circuito": "", "tipoDifferenziale": "", "datiNominali": ""}])

        edited = st.data_editor(
            df,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "numero": "N° interruttore",
                "circuito": "Circuito",
                "tipoDifferenziale": "Tipo differenziale",
                "datiNominali": "Dati nominali"
            }
        )

        s["interruttori"] = edited.fillna("").to_dict("records")

        st.markdown("### Anteprima dati")
        st.dataframe(flatten_data([s]), use_container_width=True)


if __name__ == "__main__":
    main()
