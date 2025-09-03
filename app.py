import streamlit as st
import pandas as pd
import io
from utils import load_price_list_by_letter, load_invoice_by_letter, match_invoices_to_pricelist, to_excel_bytes

st.set_page_config(page_title="Verifica prezzi: Listino vs Fatture", layout="wide")

st.title("\ud83d\udd0e Verifica prezzi: Listino (PA) vs Fatture")
st.markdown("""
Carica un file di listino e una o più fatture (Excel).
Il matching avviene esclusivamente per Codice e Prezzo.
""")

with st.expander("\u2699\ufe0f Impostazioni colonne (default secondo istruzioni)"):
    st.write("Se i file hanno strutture diverse, puoi modificare le lettere di colonna qui sotto.")
    col1, col2 = st.columns(2)
    with col1:
        price_code_col_letter = st.text_input("Listino - Colonna Codice (lettera)", value="C")
        price_price_col_letter = st.text_input("Listino - Colonna Prezzo (lettera)", value="AK")
    with col2:
        inv_code_col_letter = st.text_input("Fatture - Colonna Codice (lettera)", value="A")
        inv_price_col_letter = st.text_input("Fatture - Colonna Prezzo (lettera)", value="D")

with st.expander("\ud83c\udfaf Tolleranza prezzo"):
    tol_abs = st.number_input("Tolleranza Assoluta (\u20ac)", value=0.01, min_value=0.0, step=0.01, format="%.2f")
    tol_pct = st.number_input("Tolleranza Percentuale (%)", value=0.0, min_value=0.0, step=0.1, format="%.1f")

price_file = st.file_uploader("\ud83d\udce5 Carica file Listino (PA) in Excel", type=["xlsx","xls"])
invoice_files = st.file_uploader("\ud83d\udce5 Carica Fatture (uno o più file Excel)", type=["xlsx","xls"], accept_multiple_files=True)

if st.button("Esegui verifica", type="primary"):
    if not price_file or not invoice_files:
        st.error("Carica sia il listino che almeno una fattura.")
        st.stop()
    try:
        price_df = load_price_list_by_letter(price_file, price_code_col_letter, price_price_col_letter)
        if price_df.empty:
            st.error("Il listino è vuoto o le colonne non sono state lette correttamente.")
            st.stop()

        results = []
        for f in invoice_files:
            inv_df = load_invoice_by_letter(f, inv_code_col_letter, inv_price_col_letter, source_name=f.name)
            if inv_df.empty:
                st.warning(f"Il file fattura '{f.name}' è vuoto o non leggibile. Saltato.")
                continue
            res = match_invoices_to_pricelist(inv_df, price_df, tol_abs=tol_abs, tol_pct=tol_pct)
            results.append(res)

        if not results:
            st.error("Nessun risultato: controlla i file di input.")
            st.stop()

        report = pd.concat(results, ignore_index=True)

        total = len(report)
        matched = int(report["price_match"].sum())
        missing_code = int((report["status"] == "CODICE_NON_TROVATO").sum())
        mismatch = int((report["status"] == "PREZZO_DIVERSO").sum())

        st.subheader("\ud83d\udcca Risultati")
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Righe fatture", f"{total}")
        k2.metric("Prezzi OK", f"{matched}")
        k3.metric("Codici mancanti", f"{missing_code}")
        k4.metric("Prezzi diversi", f"{mismatch}")

        with st.expander("\ud83d\udd0d Dettaglio risultati"):
            st.dataframe(report, use_container_width=True, hide_index=True)

        # Build nice DataFrame with human-friendly columns
        nice = report[["source_file","row_index","code","invoice_price","list_price","delta_abs","delta_pct","status"]].copy()
        nice = nice.rename(columns={
            "source_file":"Fattura",
            "row_index":"Riga fattura",
            "code":"Codice",
            "invoice_price":"Prezzo Fattura",
            "list_price":"Prezzo Listino",
            "delta_abs":"Delta \u20ac",
            "delta_pct":"Delta %",
            "status":"Esito"
        })

        csv_bytes = nice.to_csv(index=False).encode("utf-8")
        st.download_button("\u2b07\ufe0f Scarica CSV", data=csv_bytes, file_name="report_verifica_prezzi.csv", mime="text/csv")

        xls_bytes = to_excel_bytes(nice)
        st.download_button("\u2b07\ufe0f Scarica Excel", data=xls_bytes, file_name="report_verifica_prezzi.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.exception(e)
        st.stop()

st.markdown("---")
st.caption("Creato per controlli rapidi fra Listino PA e Fatture in ambito B2B Ho.Re.Ca.")
