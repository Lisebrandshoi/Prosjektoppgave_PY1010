# -*- coding: utf-8 -*-
"""
Prosjektoppgave PY1010
Lise Brændshøi (lisebrandshoi@gmail.com)
Oppdatert 23.03.2026
"""

#%% Del a)


import pandas as pd

filnavn = "support_uke_24.xlsx"  # leser inn excel-filen
data = pd.read_excel(filnavn)  


import numpy as np 

u_dag = data["Ukedag"].values  # vi gjør om kolonnene i filen til arrays
kl_slett = data["Klokkeslett"].values
varighet = data["Varighet"].values
score = data["Tilfredshet"].values


#%% Del b)

import matplotlib.pyplot as plt 

antall = {}  

for dag in u_dag:  # teller antall per dag
    if dag in antall:
        antall [dag] += 1
    else: 
        antall [dag] = 1

dager =list(antall.keys())
verdier = list(antall.values())

plt.bar(dager, verdier)
plt.xlabel("Ukedag")
plt.ylabel("Antall henvendelser")
plt.title("Antall henvendelser per ukedag")

plt.show()


#%% Del c) 

varighet = pd.to_timedelta(varighet)  # gjør om tidsformat

minste = pd.to_timedelta(varighet).min()  # finner korteste samtaletid
lengste = pd.to_timedelta(varighet).max()  # finner lengste samtaletid


def til_min_og_sek(tid):  # funksjon for å få pen utskrift av antall minutter og sekunder
    total_sek = int(tid.total_seconds())  # regner om til sekunder
    minutter = total_sek // 60  # finner antall minutter
    sekunder = total_sek % 60  # finner antall sekunder
    return minutter, sekunder

min_m, min_s = til_min_og_sek(minste)
max_m, max_s = til_min_og_sek(lengste)


print (f"Korteste samtaletid var {min_m} minutter og {min_s} sekunder.")
print(f"Lengste samtaletid var {max_m} minutter og {max_s} sekunder.")


#%% Del d) 

snitt = pd.to_timedelta(varighet).mean()  # finne gjennomsnittlig tid
snitt_m, snitt_s = til_min_og_sek(snitt)

print (f"Gjennomsnittlig samtaletid var {snitt_m} minutter og {snitt_s} sekunder.")


#%% Del e) 

kl_slett = pd.to_datetime(kl_slett)  # gjør til rett tidsformat

teller_08_10 = 0  # lager tellere 
teller_10_12 = 0
teller_12_14 = 0
teller_14_16 = 0

for tid in (kl_slett):  # henter timen fra klokkeslettet
    time = tid.hour

    if 8 <= time < 10:  # setter de ulike intervallene og teller antall henvendelser per intervall
        teller_08_10 += 1
    elif 10 <= time < 12:
        teller_10_12 += 1
    elif 12 <= time < 14: 
        teller_12_14 += 1
    elif 14 <= time < 16:
        teller_14_16 += 1

verdier = [teller_08_10, teller_10_12, teller_12_14, teller_14_16]
labels = ["08-10", "10-12", "12-14", "14-16"]

plt.pie(verdier, labels=labels, autopct='%1.1f%%')
plt.title("Fordeling av henvendelser gjennom dagen")

plt.show


#%% Del f)

tellende_score = data["Tilfredshet"].dropna()  # fjerner de tomme verdiene

positive = 0  # lager tellere
negative = 0

for h in tellende_score:  # går gjennnom alle svarene
    if 1 <= h <= 6:
        negative += 1
    elif 9 <= h <= 10:
        positive += 1

totalt = len(tellende_score)  # totalt antall svar

prosent_negative = (negative/totalt) * 100  # regner ut prosent
prosent_positive = (positive/totalt) * 100

nps = prosent_positive - prosent_negative  # regner ut NPS

print (f"Supportavdelingens NPS er {nps:.1f}.")