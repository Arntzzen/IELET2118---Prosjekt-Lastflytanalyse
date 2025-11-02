import numpy as np
from prettytable import PrettyTable
import pandas as pd

A = (0.04 + 0.43j, 8.7, 600)
B = (0.03 + 0.33j, 11.3, 900)
C = (0.02 + 0.32j, 11.4, 1000)

linjer = ("1-2", "2-3", "3-4", "4-5", "5-6", "6-7", "7-8", "1-8", "1-6")
lengder = np.array([22.9, 18.8, 47.8, 12.1, 71.2, 54.4, 49.4, 105.8, 116.7])
klasser = np.array([A, A, A, B, C, C, A, A, A], dtype=object)


impedanser = {linjer[i]: lengder[i] * klasser[i][0] for i in range(len(linjer))}
impedanser_pu = {linje: Z / 90 for linje, Z in impedanser.items()}


kapasitanser = {linjer[i]: lengder[i] * klasser[i][1] for i in range(len(linjer))}



################### DataFrame og Excel eksport ###################
def round_complex(z):
    return complex(round(z.real, 4), round(z.imag, 4))

# Lag DataFrame og rund alt
df_imp = pd.DataFrame({
    "Linje": list(impedanser.keys()),
    "Impedans (ohm)": [round_complex(z) for z in impedanser.values()],
    "Impedans (p.u.)": [round_complex(z) for z in impedanser_pu.values()],
    "Kapasitans (nF)": [round(c) for c in kapasitanser.values()]
})

# Skriv til Excel
df_imp.to_excel("linjeimpedanser.xlsx", index=False)
print(df_imp)



################### PrettyTable utskrift ###################
linjeimpedanser = PrettyTable()
linjeimpedanser.field_names = ["Linjeimpedanser", "Impedans (ohm)"]

for key, value in impedanser.items():
    linjeimpedanser.add_row([key, round(value, 3)])
print(linjeimpedanser)
df_imp.to_excel('Linjeimpedanser.xlsx', index=False)


linjekapasitanser = PrettyTable()
linjekapasitanser.field_names = ["Linjekapasitanser", "Kapasitans (F)"]

for key, value in kapasitanser.items():
    linjekapasitanser.add_row([key, round(value, 3)])

print(linjekapasitanser)

linjeimpedanser_pu = PrettyTable()
linjeimpedanser_pu.field_names = ["Linjeimpedanser i pu", "Impedans (pu)"]

for key, value in impedanser_pu.items():
    linjeimpedanser_pu.add_row([key, round(value, 4)])

print(linjeimpedanser_pu)