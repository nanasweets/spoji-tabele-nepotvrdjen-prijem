import pandas as pd

FILE1 = "Otvoreni PLTS na dan 21.09.2026.xlsx"
FILE2 = "Nepotvrđen prijem.xlsx"

df1 = pd.read_excel(FILE1, sheet_name="Sheet1")
df2 = pd.read_excel(FILE2, sheet_name=0)

df1.columns = df1.columns.str.strip()
df2.columns = df2.columns.str.strip()

df1 = df1.rename(columns={"brojdolg": "PLTS", "dokdatum": "Datum dokumenta"})

df1 = df1[["PLTS", "Datum dokumenta", "artikal_sifra", "ean",
           "naziv_artikla", "kolicina", "iznos", "zabeleska"]].copy()

df2 = df2[["TIKET 2", "Datum PLTS", "Broj PLTS", "Sifra artikla",
           "Naziv artikla", "Brend", "Datum obrade", "Nedostaje potvrda",
           "Lokacija aparata", "TMS nalog", "Datum TMS naloga",
           "KOMENTAR", "NPT", "Kolona1"]].copy()

# Normalizuj ključeve u string (da ne pravi problem sa brojevima)
df1["PLTS"] = df1["PLTS"].astype(str).str.strip()
df2["Broj PLTS"] = df2["Broj PLTS"].astype(str).str.strip()

# Zadrži SAMO JEDAN red iz druge tabele po PLTS-u da ne bi duplirao redove u prvoj
df2 = df2.drop_duplicates(subset=["Broj PLTS"], keep="first")

# LEFT JOIN - svi redovi iz prve tabele ostaju, kolone iz druge se dodaju gde ima poklapanja
merged = pd.merge(df1, df2, left_on="PLTS", right_on="Broj PLTS", how="left")

final_cols = ["PLTS", "Datum dokumenta", "artikal_sifra", "ean",
              "naziv_artikla", "kolicina", "iznos", "zabeleska",
              "TIKET 2", "Datum PLTS", "Broj PLTS", "Sifra artikla",
              "Naziv artikla", "Brend", "Datum obrade", "Nedostaje potvrda",
              "Lokacija aparata", "TMS nalog", "Datum TMS naloga",
              "KOMENTAR", "NPT", "Kolona1"]

merged = merged.reindex(columns=final_cols)
merged.to_excel("spojena_tabela.xlsx", index=False)

print(f"Redova u prvoj tabeli:    {len(df1)}")
print(f"Redova u spojenoj tabeli: {len(merged)}")
print(f"Poklopljeno (popunjeno):  {merged['Broj PLTS'].notna().sum()}")
print(f"Bez poklapanja (prazno):  {merged['Broj PLTS'].isna().sum()}")
