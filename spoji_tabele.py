import pandas as pd

FILE1 = "Otvoreni PLTS na dan 21.09.2026.xlsx"
FILE2 = "Nepotvrđen prijem.xlsx"

SHEET1 = "Sheet1"
SHEET2 = 0

df1 = pd.read_excel(FILE1, sheet_name=SHEET1)
df2 = pd.read_excel(FILE2, sheet_name=SHEET2)

df1.columns = df1.columns.str.strip()
df2.columns = df2.columns.str.strip()

# Prva tabela (glavna, ostaje nepromenjena po broju redova)
df1 = df1.rename(columns={"brojdolg": "PLTS", "dokdatum": "Datum dokumenta"})
df1 = df1[["PLTS", "Datum dokumenta", "artikal_sifra", "ean",
           "naziv_artikla", "kolicina", "iznos", "zabeleska"]]

# Druga tabela
df2 = df2[["TIKET 2", "Datum PLTS", "Broj PLTS", "Sifra artikla",
           "Naziv artikla", "Brend", "Datum obrade", "Nedostaje potvrda",
           "Lokacija aparata", "TMS nalog", "Datum TMS naloga",
           "KOMENTAR", "NPT", "Kolona1"]]

# KLJUČNO: ako druga tabela ima više redova za isti PLTS, zadrži samo jedan
df2 = df2.drop_duplicates(subset=["Broj PLTS"], keep="first")

# Spajanje: svi redovi iz PRVE tabele, dodaju se kolone iz druge
merged = pd.merge(df1, df2, left_on="PLTS", right_on="Broj PLTS", how="left")

final_cols = ["PLTS", "Datum dokumenta", "artikal_sifra", "ean",
              "naziv_artikla", "kolicina", "iznos", "zabeleska",
              "TIKET 2", "Datum PLTS", "Broj PLTS", "Sifra artikla",
              "Naziv artikla", "Brend", "Datum obrade", "Nedostaje potvrda",
              "Lokacija aparata", "TMS nalog", "Datum TMS naloga",
              "KOMENTAR", "NPT", "Kolona1"]

merged = merged.reindex(columns=final_cols)
merged.to_excel("spojena_tabela.xlsx", index=False)

print("Gotovo: spojena_tabela.xlsx")
print(f"Redova u prvoj tabeli: {len(df1)}")
print(f"Redova u spojenoj tabeli: {len(merged)}")
