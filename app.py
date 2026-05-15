
import io
import re
import json
from datetime import datetime

import pandas as pd
import streamlit as st

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, PageBreak, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
except Exception:
    colors = None


APP_TITLE = "Gestionale schede differenziali"


EXPECTED_COLUMNS = [
    "Sede",
    "Blocco",
    "Piano",
    "Nome Quadro",
    "Reparto",
    "N° interruttore",
    "Circuito",
    "Tipo differenziale",
    "Dati nominali",
    "Note",
    "File origine",
]


def norm_col(c):
    return str(c).strip().lower().replace("\n", " ")


def normalize_database(df):
    rename_map = {}
    aliases = {
        "sede": "Sede",
        "blocco": "Blocco",
        "zona": "Blocco",
        "piano": "Piano",
        "nome quadro": "Nome Quadro",
        "quadro": "Nome Quadro",
        "reparto": "Reparto",
        "utenza": "Reparto",
        "reparto / utenza": "Reparto",
        "n° interruttore": "N° interruttore",
        "n interruttore": "N° interruttore",
        "n. interruttore": "N° interruttore",
        "numero interruttore": "N° interruttore",
        "circuito": "Circuito",
        "tipo differenziale": "Tipo differenziale",
        "dati nominali": "Dati nominali",
        "note": "Note",
        "file origine": "File origine",
    }

    for col in df.columns:
        key = norm_col(col)
        if key in aliases:
            rename_map[col] = aliases[key]

    df = df.rename(columns=rename_map)

    for col in EXPECTED_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    df = df[EXPECTED_COLUMNS].fillna("").astype(str)

    # pulizia numeri tipo "1.0"
    df["N° interruttore"] = df["N° interruttore"].str.replace(r"\.0$", "", regex=True)

    return df


def load_database(uploaded):
    xls = pd.ExcelFile(uploaded)
    # Preferisce "Tutti i dati" se presente, altrimenti primo foglio
    sheet = "Tutti i dati" if "Tutti i dati" in xls.sheet_names else xls.sheet_names[0]
    df = pd.read_excel(uploaded, sheet_name=sheet, dtype=str)
    return normalize_database(df)


def make_default_db():
    return pd.DataFrame(columns=EXPECTED_COLUMNS)


def sheet_key(row):
    return (
        str(row.get("Sede", "")),
        str(row.get("Blocco", "")),
        str(row.get("Piano", "")),
        str(row.get("Nome Quadro", "")),
        str(row.get("Reparto", "")),
    )


def get_sheets(df):
    if df.empty:
        return pd.DataFrame(columns=["Sede", "Blocco", "Piano", "Nome Quadro", "Reparto", "Interruttori"])
    g = (
        df.groupby(["Sede", "Blocco", "Piano", "Nome Quadro", "Reparto"], dropna=False)
        .size()
        .reset_index(name="Interruttori")
        .sort_values(["Sede", "Blocco", "Nome Quadro", "Piano", "Reparto"], kind="stable")
    )
    return g


def filter_df(df, sede, blocco, query):
    out = df.copy()
    if sede and sede != "Tutte":
        out = out[out["Sede"] == sede]
    if blocco and blocco != "Tutti":
        out = out[out["Blocco"] == blocco]
    if query:
        q = query.lower()
        mask = out.apply(lambda r: q in " ".join(str(v).lower() for v in r.values), axis=1)
        out = out[mask]
    return out


def update_sheet(df, old_key, header, switches):
    mask = (
        (df["Sede"] == old_key[0]) &
        (df["Blocco"] == old_key[1]) &
        (df["Piano"] == old_key[2]) &
        (df["Nome Quadro"] == old_key[3]) &
        (df["Reparto"] == old_key[4])
    )
    remaining = df.loc[~mask].copy()

    rows = []
    for _, sw in switches.fillna("").iterrows():
        # evita righe completamente vuote
        if not any(str(sw.get(c, "")).strip() for c in ["N° interruttore", "Circuito", "Tipo differenziale", "Dati nominali", "Note"]):
            continue
        row = {c: "" for c in EXPECTED_COLUMNS}
        row.update(header)
        row["N° interruttore"] = str(sw.get("N° interruttore", ""))
        row["Circuito"] = str(sw.get("Circuito", ""))
        row["Tipo differenziale"] = str(sw.get("Tipo differenziale", ""))
        row["Dati nominali"] = str(sw.get("Dati nominali", ""))
        row["Note"] = str(sw.get("Note", ""))
        row["File origine"] = str(sw.get("File origine", ""))
        rows.append(row)

    if not rows:
        row = {c: "" for c in EXPECTED_COLUMNS}
        row.update(header)
        rows.append(row)

    new_df = pd.concat([remaining, pd.DataFrame(rows)], ignore_index=True)
    return normalize_database(new_df)


def delete_sheet(df, key):
    mask = (
        (df["Sede"] == key[0]) &
        (df["Blocco"] == key[1]) &
        (df["Piano"] == key[2]) &
        (df["Nome Quadro"] == key[3]) &
        (df["Reparto"] == key[4])
    )
    return df.loc[~mask].reset_index(drop=True)


def make_excel(df):
    df = normalize_database(df)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Tutti i dati", index=False)
        sheets = get_sheets(df)
        for sede in sorted([x for x in df["Sede"].unique() if str(x).strip()]):
            df_sede = df[df["Sede"] == sede]
            safe_sede = re.sub(r'[\[\]\:\*\?\/\\]', "_", str(sede))[:31]
            df_sede.to_excel(writer, sheet_name=safe_sede or "Sede", index=False)

        for blocco in sorted([x for x in df["Blocco"].unique() if str(x).strip()]):
            df_b = df[df["Blocco"] == blocco]
            safe = re.sub(r'[\[\]\:\*\?\/\\]', "_", f"Blocco {blocco}")[:31]
            # evita duplicati nome foglio
            if safe not in writer.book.sheetnames:
                df_b.to_excel(writer, sheet_name=safe, index=False)

    output.seek(0)
    return output.getvalue()


def para(text, style):
    return Paragraph(str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"), style)


def make_pdf(df, layout_title="SCHEDA VERIFICA INTERRUTTORI DIFFERENZIALI", only_key=None):
    if colors is None:
        raise RuntimeError("ReportLab non installato.")

    df = normalize_database(df)
    if only_key is not None:
        df = df[
            (df["Sede"] == only_key[0]) &
            (df["Blocco"] == only_key[1]) &
            (df["Piano"] == only_key[2]) &
            (df["Nome Quadro"] == only_key[3]) &
            (df["Reparto"] == only_key[4])
        ]

    output = io.BytesIO()
    doc = SimpleDocTemplate(
        output,
        pagesize=A4,
        rightMargin=8*mm,
        leftMargin=8*mm,
        topMargin=8*mm,
        bottomMargin=8*mm,
    )

    styles = getSampleStyleSheet()
    small = ParagraphStyle("small", parent=styles["Normal"], fontSize=7, leading=8)
    normal = ParagraphStyle("normal", parent=styles["Normal"], fontSize=8, leading=9)
    bold = ParagraphStyle("bold", parent=styles["Normal"], fontSize=8, leading=9, fontName="Helvetica-Bold")
    title = ParagraphStyle("title", parent=styles["Normal"], fontSize=12, leading=14, alignment=1, fontName="Helvetica-Bold")

    story = []
    sheet_rows = get_sheets(df)

    for idx, sh in sheet_rows.iterrows():
        key = sheet_key(sh)
        rows = df[
            (df["Sede"] == key[0]) &
            (df["Blocco"] == key[1]) &
            (df["Piano"] == key[2]) &
            (df["Nome Quadro"] == key[3]) &
            (df["Reparto"] == key[4])
        ].copy()

        story.append(Table([
            [para("LOGO / ENTE", bold), Paragraph(layout_title, title), para("Rev. 00", normal)]
        ], colWidths=[35*mm, 125*mm, 34*mm], style=[
            ("GRID", (0, 0), (-1, -1), 1.1, colors.black),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 7),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ]))

        story.append(Table([
            [
                para(f"<b>Sede:</b> {key[0]}", normal),
                para(f"<b>Blocco:</b> {key[1]}", normal),
                para(f"<b>Piano:</b> {key[2]}", normal),
                para(f"<b>Nome quadro:</b> {key[3]}", normal),
            ],
            [
                para(f"<b>Reparto:</b> {key[4]}", normal),
                para("", normal),
                para("", normal),
                para("", normal),
            ],
        ], colWidths=[45*mm, 35*mm, 30*mm, 84*mm], style=[
            ("GRID", (0, 0), (-1, -1), 0.8, colors.black),
            ("SPAN", (0, 1), (-1, 1)),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))

        table_data = [[
            para("<b>N° INT.</b>", small),
            para("<b>CIRCUITO</b>", small),
            para("<b>TIPO DIFFERENZIALE</b>", small),
            para("<b>DATI NOMINALI</b>", small),
            para("<b>NOTE</b>", small),
        ]]

        rows = rows.sort_values("N° interruttore", kind="stable")
        min_rows = max(18, len(rows))
        data_rows = rows.to_dict("records")
        for i in range(min_rows):
            r = data_rows[i] if i < len(data_rows) else {}
            table_data.append([
                para(r.get("N° interruttore", ""), small),
                para(r.get("Circuito", ""), small),
                para(r.get("Tipo differenziale", ""), small),
                para(r.get("Dati nominali", ""), small),
                para(r.get("Note", ""), small),
            ])

        story.append(Table(table_data, colWidths=[17*mm, 65*mm, 42*mm, 45*mm, 25*mm], repeatRows=1, style=[
            ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
            ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("ALIGN", (0, 0), (0, -1), "CENTER"),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]))

        story.append(Table([
            [para("<b>Data:</b>", normal), para("<b>Nome manutentore:</b>", normal), para("<b>Firma:</b>", normal)]
        ], colWidths=[45*mm, 80*mm, 69*mm], style=[
            ("GRID", (0, 0), (-1, -1), 0.8, colors.black),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("MINROWHEIGHT", (0, 0), (-1, -1), 16*mm),
        ]))

        if idx < len(sheet_rows) - 1:
            story.append(PageBreak())

    doc.build(story)
    output.seek(0)
    return output.getvalue()


def init_state():
    if "df" not in st.session_state:
        st.session_state.df = make_default_db()
    if "selected_key" not in st.session_state:
        st.session_state.selected_key = None
    if "last_saved" not in st.session_state:
        st.session_state.last_saved = ""


def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    init_state()

    st.title(APP_TITLE)
    st.caption("Gestione database, modifica schede, creazione nuove schede e stampa nel layout configurato.")

    with st.sidebar:
        st.header("Database")
        uploaded = st.file_uploader("Carica database Excel", type=["xlsx"])

        if uploaded is not None:
            if st.button("Apri database", type="primary"):
                st.session_state.df = load_database(uploaded)
                st.session_state.selected_key = None
                st.success(f"Database caricato: {len(st.session_state.df)} righe.")

        st.download_button(
            "Scarica database aggiornato",
            data=make_excel(st.session_state.df),
            file_name=f"database_differenziali_aggiornato_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            disabled=st.session_state.df.empty,
        )

        if not st.session_state.df.empty:
            try:
                st.download_button(
                    "Scarica PDF di tutte le schede",
                    data=make_pdf(st.session_state.df),
                    file_name=f"schede_differenziali_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                    mime="application/pdf",
                )
            except Exception as e:
                st.error(f"PDF non disponibile: {e}")

        st.divider()

        if st.button("Crea nuova scheda"):
            new_header = {
                "Sede": "",
                "Blocco": "",
                "Piano": "",
                "Nome Quadro": "NUOVO QUADRO",
                "Reparto": "",
            }
            new_row = {c: "" for c in EXPECTED_COLUMNS}
            new_row.update(new_header)
            new_row["N° interruttore"] = "1"
            st.session_state.df = normalize_database(pd.concat([st.session_state.df, pd.DataFrame([new_row])], ignore_index=True))
            st.session_state.selected_key = ("", "", "", "NUOVO QUADRO", "")
            st.rerun()

    df = st.session_state.df

    if df.empty:
        st.info("Carica il database Excel di Cittadella oppure crea una nuova scheda.")
        return

    top1, top2, top3, top4 = st.columns([1, 1, 2, 1])
    sedi = ["Tutte"] + sorted([x for x in df["Sede"].unique() if str(x).strip()])
    blocchi = ["Tutti"] + sorted([x for x in df["Blocco"].unique() if str(x).strip()])

    sede = top1.selectbox("Sede", sedi)
    blocco = top2.selectbox("Blocco", blocchi)
    query = top3.text_input("Cerca")
    metric_area = top4.empty()

    filtered_df = filter_df(df, sede, blocco, query)
    sheets = get_sheets(filtered_df)
    metric_area.metric("Schede", len(sheets))

    left, right = st.columns([1.05, 2])

    with left:
        st.subheader("Elenco schede")
        if sheets.empty:
            st.warning("Nessuna scheda trovata.")
            return

        labels = []
        keys = []
        for _, r in sheets.iterrows():
            key = sheet_key(r)
            keys.append(key)
            labels.append(f"{r['Nome Quadro']} | {r['Sede']} | Blocco {r['Blocco']} | Piano {r['Piano']} | {r['Reparto']} ({r['Interruttori']})")

        if st.session_state.selected_key not in keys:
            st.session_state.selected_key = keys[0]

        idx = keys.index(st.session_state.selected_key)
        chosen = st.selectbox("Seleziona", labels, index=idx)
        st.session_state.selected_key = keys[labels.index(chosen)]

        c1, c2 = st.columns(2)
        if c1.button("Duplica"):
            key = st.session_state.selected_key
            subset = df[
                (df["Sede"] == key[0]) &
                (df["Blocco"] == key[1]) &
                (df["Piano"] == key[2]) &
                (df["Nome Quadro"] == key[3]) &
                (df["Reparto"] == key[4])
            ].copy()
            subset["Nome Quadro"] = subset["Nome Quadro"] + " - copia"
            st.session_state.df = normalize_database(pd.concat([df, subset], ignore_index=True))
            st.session_state.selected_key = (key[0], key[1], key[2], key[3] + " - copia", key[4])
            st.rerun()

        if c2.button("Elimina"):
            st.session_state.df = delete_sheet(df, st.session_state.selected_key)
            st.session_state.selected_key = None
            st.rerun()

        st.divider()
        st.subheader("Esporta selezione")
        try:
            st.download_button(
                "PDF scheda selezionata",
                data=make_pdf(df, only_key=st.session_state.selected_key),
                file_name=f"{st.session_state.selected_key[3]}_scheda.pdf".replace("/", "_"),
                mime="application/pdf",
            )
        except Exception as e:
            st.error(str(e))

    with right:
        key = st.session_state.selected_key
        subset = df[
            (df["Sede"] == key[0]) &
            (df["Blocco"] == key[1]) &
            (df["Piano"] == key[2]) &
            (df["Nome Quadro"] == key[3]) &
            (df["Reparto"] == key[4])
        ].copy()

        st.subheader("Modifica scheda")

        h1, h2, h3, h4, h5 = st.columns([1, 1, 1, 2, 2])
        header = {
            "Sede": h1.text_input("Sede", key[0]),
            "Blocco": h2.text_input("Blocco", key[1]),
            "Piano": h3.text_input("Piano", key[2]),
            "Nome Quadro": h4.text_input("Nome quadro", key[3]),
            "Reparto": h5.text_input("Reparto", key[4]),
        }

        st.markdown("#### Interruttori")
        sw_cols = ["N° interruttore", "Circuito", "Tipo differenziale", "Dati nominali", "Note", "File origine"]
        switches = subset[sw_cols].reset_index(drop=True)
        edited = st.data_editor(
            switches,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            column_config={
                "N° interruttore": st.column_config.TextColumn("N° interruttore", width="small"),
                "Circuito": st.column_config.TextColumn("Circuito", width="large"),
                "Tipo differenziale": st.column_config.TextColumn("Tipo differenziale", width="medium"),
                "Dati nominali": st.column_config.TextColumn("Dati nominali", width="large"),
                "Note": st.column_config.TextColumn("Note", width="medium"),
                "File origine": st.column_config.TextColumn("File origine", width="medium"),
            },
        )

        save_col, preview_col = st.columns([1, 2])
        if save_col.button("Salva modifiche", type="primary"):
            st.session_state.df = update_sheet(df, key, header, edited)
            st.session_state.selected_key = sheet_key(header)
            st.success("Modifiche salvate nel database in memoria. Scarica il database aggiornato dalla barra laterale.")
            st.rerun()

        with preview_col.expander("Anteprima righe database"):
            preview_df = pd.DataFrame([{**header, **r} for _, r in edited.fillna("").iterrows()])
            st.dataframe(preview_df, use_container_width=True)


if __name__ == "__main__":
    main()
