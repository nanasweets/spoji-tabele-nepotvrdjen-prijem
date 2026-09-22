import pandas as pd

FILE1 = "Otvoreni PLTS na dan 21.09.2026.xlsx"
FILE2 = "nepotvrđen prijem.xlsx"

SHEET1 = "Sheet1"
SHEET2 = 0  # ako list ima tačan naziv, npr. "nepotvrđen prijem", upiši ga ovde

df1 = pd.read_excel(FILE1, sheet_name=SHEET1)
df2 = pd.read_excel(FILE2, sheet_name=SHEET2)

# skini razmake iz naziva kolona
df1.columns = df1.columns.str.strip()
df2.columns = df2.columns.str.strip()

# Prva tabela
df1 = df1.rename(columns={
    "brojdolg": "PLTS",
    "dokdatum": "Datum dokumenta"
})

df1 = df1[[
    "PLTS",
    "Datum dokumenta",
    "artikal_sifra",
    "ean",
    "naziv_artikla",
    "kolicina",
    "iznos",
    "zabeleska"
]]

# Druga tabela - samo potrebne kolone
df2 = df2[[
    "TIKET 2",
    "Datum PLTS",
    "Broj PLTS",
    "Sifra artikla",
    "Naziv artikla",
    "Brend",
    "Datum obrade",
    "Nedostaje potvrda",
    "Lokacija aparata",
    "TMS nalog",
    "Datum TMS naloga",
    "KOMENTAR",
    "NPT",
    "Kolona1"
]]

# Spajanje isključivo po PLTS
merged = pd.merge(
    df1,
    df2,
    left_on="PLTS",
    right_on="Broj PLTS",
    how="left"
)

# Redosled kolona u finalnoj tabeli
final_cols = [
    "PLTS",
    "Datum dokumenta",
    "artikal_sifra",
    "ean",
    "naziv_artikla",
    "kolicina",
    "iznos",
    "zabeleska",
    "TIKET 2",
    "Datum PLTS",
    "Broj PLTS",
    "Sifra artikla",
    "Naziv artikla",
    "Brend",
    "Datum obrade",
    "Nedostaje potvrda",
    "Lokacija aparata",
    "TMS nalog",
    "Datum TMS naloga",
    "KOMENTAR",
    "NPT",
    "Kolona1"
]

merged = merged.reindex(columns=final_cols)

# Snimi finalnu tabelu
merged.to_excel("spojena_tabela.xlsx", index=False)

print("Gotovo: spojena_tabela.xlsx")