
import os
import csv
from itertools import combinations
from collections import defaultdict
import pandas as pd

# ---------------------------- КОНФИГ ------------------------------------ #
KATALOG = "start"

IMENA = {
    1: "Список ПМ на 01.08.xlsx",
    2: "Список ИВТ на 01.08.xlsx",
    3: "Список ИТСС на 01.08.xlsx",
    4: "Список ИБ на 01.08.xlsx",
    5: "Список ПМ на 02.08.xlsx",
    6: "Список ИВТ на 02.08.xlsx",
    7: "Список ИТСС на 02.08.xlsx",
    8: "Список ИБ на 02.08.xlsx",
    9: "Список ПМ на 03.08.xlsx",
    10: "Список ИВТ на 03.08.xlsx",
    11: "Список ИТСС на 03.08.xlsx",
    12: "Список ИБ на 03.08.xlsx",
    13: "Список ПМ на 04.08.xlsx",
    14: "Список ИВТ на 04.08.xlsx",
    15: "Список ИТСС на 04.08.xlsx",
    16: "Список ИБ на 04.08.xlsx",
}

GRUPPY = {
    1: [1, 2, 3, 4],    # 01.08
    2: [5, 6, 7, 8],    # 02.08
    3: [9, 10, 11, 12], # 03.08
    4: [13, 14, 15, 16] # 04.08
}


PORYADOK_TROEK = [
    (0, 1, 2),  # ПМ-ИВТ-ИТСС
    (0, 1, 3),  # ПМ-ИВТ-ИБ
    (1, 2, 3),  # ИВТ-ИТСС-ИБ
    (0, 2, 3),  # ПМ-ИТСС-ИБ
]
PORYADOK_PAR = [
    (0, 1),  # ПМ-ИВТ
    (0, 2),  # ПМ-ИТСС
    (0, 3),  # ПМ-ИБ
    (1, 2),  # ИВТ-ИТСС
    (1, 3),  # ИВТ-ИБ
    (2, 3),  # ИТСС-ИБ
]

print("Читаю файлы…")
identifikatory: dict[int, set[int]] = {}
prioritety: defaultdict[int, dict[int, int]] = defaultdict(dict)
kolichestva_strok: dict[int, int] = {}
schetchiki_prioritetov: dict[int, dict[int, int]] = {}

for nomer_fayla, imya_fayla in IMENA.items():
    put = os.path.join(KATALOG, imya_fayla)
    if not os.path.exists(put):
        put = os.path.join(KATALOG.capitalize(), imya_fayla)
    if not os.path.exists(put):
        raise FileNotFoundError(put)

    dataframe = pd.read_excel(put)


    identifikatory[nomer_fayla] = set(dataframe["ID"].astype(int))
    kolichestva_strok[nomer_fayla] = len(dataframe)


    karta_prioritetov = dataframe.set_index("ID")["Приоритет"].to_dict()
    for identifikator, prioritet in karta_prioritetov.items():
        prioritety[identifikator][nomer_fayla] = prioritet
    

    seriya_prioritetov = pd.to_numeric(dataframe["Приоритет"], errors="coerce").dropna().astype(int)
    schetchiki_prioritetov[nomer_fayla] = {prioritet: int((seriya_prioritetov == prioritet).sum()) for prioritet in [1, 2, 3, 4]}



def razmer_peresecheniya(fayly):
    obshchie = identifikatory[fayly[0]].copy()
    for fayl in fayly[1:]:
        obshchie &= identifikatory[fayl]
    return len(obshchie)

def dubli_id_prioriteta(spisok_faylov):
    dublikaty = 0
    identifikatory_v_fayly = defaultdict(list)
    for fayl in spisok_faylov:
        for identifikator in identifikatory[fayl]:
            identifikatory_v_fayly[identifikator].append(fayl)
    for identifikator, spisok_faylov_id in identifikatory_v_fayly.items():
        if len(spisok_faylov_id) < 2:
            continue
        prioritety_spisok = [prioritety[identifikator][fayl] for fayl in spisok_faylov_id]
        if len(prioritety_spisok) != len(set(prioritety_spisok)):
            dublikaty += 1
    return dublikaty

def formatirovat(nomera_faylov):
    return ", ".join(IMENA[nomer] for nomer in nomera_faylov)

zapisi: list[tuple[str, int]] = []


for nomer_fayla in range(1, 17):
    zapisi.append((f"Число строк в \"{IMENA[nomer_fayla]}\"", kolichestva_strok[nomer_fayla]))


for gruppa, spisok_faylov in GRUPPY.items():

    zapisi.append((f"Количество пересечений ID в файлах {formatirovat(spisok_faylov)}", razmer_peresecheniya(spisok_faylov)))

    for indeksy in PORYADOK_TROEK:
        kombinatsiya = [spisok_faylov[indeks] for indeks in indeksy]
        zapisi.append((f"Количество пересечений ID в файлах {formatirovat(kombinatsiya)}", razmer_peresecheniya(kombinatsiya)))

    for indeksy in PORYADOK_PAR:
        kombinatsiya = [spisok_faylov[indeks] for indeks in indeksy]
        zapisi.append((f"Количество пересечений ID в файлах {formatirovat(kombinatsiya)}", razmer_peresecheniya(kombinatsiya)))

    zapisi.append((f"Число совпадеений Приоритет в группе файлов {gruppa}", dubli_id_prioriteta(spisok_faylov)))


    for fayl in spisok_faylov:
        for prioritet in [1, 2, 3, 4]:
            kolichestvo_prioriteta = schetchiki_prioritetov[fayl][prioritet]
            tekst = f"Количество совпадений Приоретет = {prioritet} в файле {IMENA[fayl]} = "
            zapisi.append((tekst, kolichestvo_prioriteta))


katalog_vyhoda = "otchet"
os.makedirs(katalog_vyhoda, exist_ok=True)
put_vyhoda = os.path.join(katalog_vyhoda, "otchet2.csv")

with open(put_vyhoda, "w", encoding="utf-8-sig", newline="") as fayl:
    pisatel = csv.writer(fayl, delimiter=";")
    pisatel.writerow(["Показатель", "Значение"])
    pisatel.writerows(zapisi)

print("Отчёт создан:", put_vyhoda)
