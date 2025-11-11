import numpy as np
from prettytable import PrettyTable
import pandas as pd


w = 2 * np.pi * 50
##################### Linjedata ###################
A = (0.04 + 0.43j, 8.7, 600)
B = (0.03 + 0.33j, 11.3, 900)
C = (0.02 + 0.32j, 11.4, 1000)

linjer = ("1-2", "2-3", "3-4", "4-5", "5-6", "6-7", "7-8", "1-8", "1-6")
lengder = np.array([22.9, 18.8, 47.8, 12.1, 71.2, 54.4, 49.4, 105.8, 116.7])
klasser = np.array([A, A, A, B, C, C, A, A, A], dtype=object)



################### Beregninger ###################
impedanser = {linjer[i]: lengder[i] * klasser[i][0] for i in range(len(linjer))}
impedanser_pu = {linje: Z / 90 for linje, Z in impedanser.items()}
admittanser_pu = {linje: 1 / Z for linje, Z in impedanser_pu.items()}
kapasitanser = {linjer[i]: lengder[i] * klasser[i][1] for i in range(len(linjer))}
shunt_impedanser = {linje: 1 / (1j * w * (C * 10**-9)) for linje, C in kapasitanser.items()}
shunt_admittanser = {linje: 1j * w * (C * 10**-9) for linje, C in kapasitanser.items()}
shunt_admittanser_pu = {linje: Y * 90 for linje, Y in shunt_admittanser.items()}
shunt_admittanser_pu_half = {linje: Y / 2 for linje, Y in shunt_admittanser_pu.items()}



################### DataFrame og Excel eksport ###################
def round_complex(z):
    return complex(round(z.real, 4), round(z.imag, 4))

df_imp = pd.DataFrame({
    "Linje": list(impedanser.keys()),
    "Impedans (ohm)": [round_complex(z) for z in impedanser.values()],
    "Impedans (p.u.)": [round_complex(z) for z in impedanser_pu.values()],
    "Admittans (p.u.)": [round_complex(y) for y in admittanser_pu.values()],
    "Kapasitans (nF)": [round(c) for c in kapasitanser.values()],
    "Shunt Impedans (ohm)": [round_complex(y) for y in shunt_impedanser.values()],
    "Shunt Admittans (S)": [round_complex(y) for y in shunt_admittanser.values()],
    "Shunt Admittans (p.u.)": [round_complex(y) for y in shunt_admittanser_pu.values()],
    "Shunt Admittans (p.u.) / 2": [round_complex(y) for y in shunt_admittanser_pu_half.values()]
})
df_imp.to_excel("linjeimpedanser.xlsx", index=False)
print(df_imp)



################## Excel eksport uten averunding ##################
df_imp_full = pd.DataFrame({
    "Linje": list(impedanser.keys()),
    "Impedans (ohm)": list(impedanser.values()),
    "Impedans (p.u.)": list(impedanser_pu.values()),
    "Admittans (p.u.)": list(admittanser_pu.values()),
    "Kapasitans (nF)": list(kapasitanser.values()),
    "Shunt Impedans (ohm)": list(shunt_impedanser.values()),
    "Shunt Admittans (ohm)": list(shunt_admittanser.values()),
    "Shunt Admittans (p.u.)": list(shunt_admittanser_pu.values()),
    "Shunt Admittans (p.u.) half": list(shunt_admittanser_pu_half.values())
})

df_imp_full.to_excel("linjeimpedanser_uten_avrunding.xlsx", index=False)
print(df_imp_full)



################### PrettyTable print ###################
linjeimpedanser = PrettyTable()
linjeimpedanser.field_names = ["Linje", "Impedans (ohm)"]
for key, value in impedanser.items():
    linjeimpedanser.add_row([key, round(value, 3)])
print(linjeimpedanser)


linjekapasitanser = PrettyTable()
linjekapasitanser.field_names = ["Linje", "Kapasitans (F)"]
for key, value in kapasitanser.items():
    linjekapasitanser.add_row([key, round(value, 3)])
print(linjekapasitanser)


linjeimpedanser_pu = PrettyTable()
linjeimpedanser_pu.field_names = ["Linje", "Impedans (pu)"]
for key, value in impedanser_pu.items():
    linjeimpedanser_pu.add_row([key, round(value, 4)])
print(linjeimpedanser_pu)


linjeadmittanser_pu = PrettyTable()
linjeadmittanser_pu.field_names = ["Linje", "Admittans (pu)"]
for key, value in admittanser_pu.items():
    linjeadmittanser_pu.add_row([key, round(value, 4)])
print(linjeadmittanser_pu)