from pathlib import Path
import pandas as pd

# Fichier MASTER (unique)
LRI = Path("data_in/fichiers_sources/LRI jeremy enrichi.xlsx")

# Noms des mois
mois_noms = {
    1: "Janvier",
    2: "Février",
    3: "Mars",
    4: "Avril",
    5: "Mai",
    6: "Juin",
    7: "Juillet",
    8: "Août",
    9: "Septembre",
    10: "Octobre",
    11: "Novembre",
    12: "Décembre",
}

# Colonnes numériques à sommer / formatter
cols_numeric_format = [
    "Commission du canal ",
    "Frais de nettoyage",
    "frais de ménage corrigé",
    "correction panier",
    "Taxe",
    "Facturé",
    "Versement OTA",
    "Montant pour le propriétaire",
    "Montant de l'agence",
    "prix proprio Brut / nuit",
    "prix nuité brut convenu",
    "réajustement",
    "Loyer Brut proprio",
    "Loyer brut réajusté",
    "TOTAL HT",
    "TOTAL TVA",
    "TOTAL TTC",
    "Commission au propriétaire",
]

def _to_float_local(x):
    s = str(x).strip()
    if not s or s.lower() == "nan":
        return float("nan")
    s = s.replace(" ", "").replace("\xa0", "")
    s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return float("nan")

def add_totals_row_top(df_base: pd.DataFrame) -> pd.DataFrame:
    df = df_base.copy()
    totals = {col: "" for col in df.columns}

    # Nombre de réservations dans le mois
    if "Réservation" in df.columns:
        try:
            totals["Réservation"] = float(len(df))
        except Exception:
            totals["Réservation"] = len(df)

    # Sommes sur les colonnes numériques
    for col in cols_numeric_format:
        if col in df.columns:
            vals = df[col].apply(_to_float_local)
            if vals.notna().any():
                totals[col] = float(vals.sum(skipna=True))
            else:
                totals[col] = ""

    totals_row = pd.DataFrame([totals])
    blank_row = pd.DataFrame([{col: "" for col in df.columns}])

    return pd.concat([totals_row, blank_row, df], ignore_index=True)

def format_sheet(writer, sheet_name: str, df_sheet: pd.DataFrame) -> None:
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]

    bold_fmt = workbook.add_format({"bold": True})
    num_fmt = workbook.add_format({"num_format": "0.00"})

    # Ligne 1 = TOTAL (ligne 0 = entêtes)
    worksheet.set_row(1, None, bold_fmt)

    col_index = {name: idx for idx, name in enumerate(df_sheet.columns)}

    for col in cols_numeric_format:
        idx = col_index.get(col)
        if idx is not None:
            worksheet.set_column(idx, idx, None, num_fmt)

def main():
    print(f"📥 Lecture LRI master pour format mensuel IN-PLACE : {LRI}")
    if not LRI.exists():
        raise SystemExit(f"⛔ Fichier introuvable : {LRI}")

    df = pd.read_excel(LRI, sheet_name=0)

    if "Arrivée" not in df.columns:
        raise SystemExit("⛔ Colonne 'Arrivée' absente, impossible de générer les onglets mensuels.")

    df["_date_arr"] = pd.to_datetime(df["Arrivée"], dayfirst=True, errors="coerce")
    df_valid = df.loc[df["_date_arr"].notna()].copy()

    df_valid["__year__"] = df_valid["_date_arr"].dt.year
    df_valid["__month__"] = df_valid["_date_arr"].dt.month
    df_valid = df_valid.sort_values(["__year__", "__month__", "_date_arr"])

    print("📊 Répartition par mois/année :")
    vc = df_valid["_date_arr"].dt.to_period("M").value_counts().sort_index()
    if not vc.empty:
        print(vc.to_string())
    else:
        print("(aucune date valide)")

    print(f"💾 Réécriture du LRI master avec onglets mensuels (Global + Mois AA) : {LRI}")
    with pd.ExcelWriter(LRI, engine="xlsxwriter") as writer:
        # Onglet Global = toutes les lignes, sans colonnes techniques
        df_global = df_valid.drop(columns=["_date_arr", "__year__", "__month__"], errors="ignore")
        df_global.to_excel(writer, sheet_name="Global", index=False)

        # Onglets mensuels
        for (year, month), df_group in df_valid.groupby(["__year__", "__month__"]):
            df_m = df_group.drop(columns=["_date_arr", "__year__", "__month__"], errors="ignore").copy()

            if "Arrivée" in df_m.columns:
                df_m = df_m.sort_values("Arrivée")

            df_m = add_totals_row_top(df_m)

            year_suffix = str(year)[-2:]
            mois_nom = mois_noms.get(month, f"M{month:02d}")
            sheet_name = f"{mois_nom} {year_suffix}"

            df_m.to_excel(writer, sheet_name=sheet_name, index=False)
            format_sheet(writer, sheet_name, df_m)

    print("✅ format_lri_master terminé : MASTER = Global + onglets mensuels.")

if __name__ == "__main__":
    main()
