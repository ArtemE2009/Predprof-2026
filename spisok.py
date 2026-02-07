import os
import random
from itertools import combinations
from collections import defaultdict, Counter
import pandas as pd



random.seed(28)
ID_MIN, ID_MAX = 100_000, 300_000
KATALOG_VYHODA = "start"
KOLONKI = [
    "ID", "Согласие", "Приоритет", "Балл Физика/ИКТ",
    "Балл Русский язык", "Балл Математика", "Балл за ИД", "Сумма баллов",
]

IMENA_FAYLOV = {
    1: "Список ПМ на 01.08.xlsx",    2: "Список ИВТ на 01.08.xlsx",
    3: "Список ИТСС на 01.08.xlsx",  4: "Список ИБ на 01.08.xlsx",
    5: "Список ПМ на 02.08.xlsx",    6: "Список ИВТ на 02.08.xlsx",
    7: "Список ИТСС на 02.08.xlsx",  8: "Список ИБ на 02.08.xlsx",
    9: "Список ПМ на 03.08.xlsx",   10: "Список ИВТ на 03.08.xlsx",
    11: "Список ИТСС на 03.08.xlsx", 12: "Список ИБ на 03.08.xlsx",
    13: "Список ПМ на 04.08.xlsx",  14: "Список ИВТ на 04.08.xlsx",
    15: "Список ИТСС на 04.08.xlsx", 16: "Список ИБ на 04.08.xlsx",
}

GRUPPY = {1: [1, 2, 3, 4], 2: [5, 6, 7, 8], 3: [9, 10, 11, 12], 4: [13, 14, 15, 16]}


class Reestr:
    def __init__(self):
        self.ispolzovannye = set()
        self.staticheskie: dict[int, dict] = {}
        self.prioritety: defaultdict[int, dict[int, int]] = defaultdict(dict)

    def novyy(self) -> int:
        while True:
            identifikator = random.randint(ID_MIN, ID_MAX)
            if identifikator not in self.ispolzovannye:
                self.ispolzovannye.add(identifikator)
                self._zapolnit_staticheskie(identifikator)
                return identifikator

    def _zapolnit_staticheskie(self, identifikator):
        stroka = {
            "ID": identifikator,
            "Согласие": random.choice(["Да", "Нет"]),
            "Балл Физика/ИКТ": random.randint(55, 100),
            "Балл Русский язык": random.randint(55, 100),
            "Балл Математика": random.randint(55, 100),
            "Балл за ИД": random.randint(0, 10),
        }
        stroka["Сумма баллов"] = stroka["Балл Физика/ИКТ"] + stroka["Балл Русский язык"] + stroka["Балл Математика"] + stroka["Балл за ИД"]
        self.staticheskie[identifikator] = stroka

    def stroka(self, identifikator: int, nomer_fayla: int, kopirovat_iz: int | None = None):
        if kopirovat_iz is not None:
            self.prioritety[identifikator][nomer_fayla] = self.prioritety[identifikator][kopirovat_iz]
        elif nomer_fayla not in self.prioritety[identifikator]:
            nachalo_gruppy = ((nomer_fayla - 1) // 4) * 4 + 1
            fayly_gruppy = range(nachalo_gruppy, nachalo_gruppy + 4)
            zanyatye = {self.prioritety[identifikator][fayl] for fayl in fayly_gruppy if fayl in self.prioritety[identifikator]}
            nabor = [prioritet for prioritet in (1, 2, 3, 4) if prioritet not in zanyatye] or [1, 2, 3, 4]
            self.prioritety[identifikator][nomer_fayla] = random.choice(nabor)
        stroka = self.staticheskie[identifikator].copy()
        stroka["Приоритет"] = self.prioritety[identifikator][nomer_fayla]
        return stroka

reestr = Reestr()
stroki_faylov: dict[int, list[int]] = {}



def eksklyuzivnye_tseli(fayly, razmery, pary, troyki, chetverka):
    tseli = {tuple(fayly): chetverka}
    for troyka in combinations(fayly, 3):
        tseli[troyka] = troyki.get(tuple(sorted(troyka)), 0) - chetverka
    for para in combinations(fayly, 2):
        para = tuple(sorted(para))
        nadmnozhestva_3 = [troyka for troyka in combinations(fayly, 3) if set(para).issubset(troyka)]
        tseli[para] = pary.get(para, 0) - chetverka - sum(tseli[troyka] for troyka in nadmnozhestva_3)
    for fayl in fayly:
        nadmnozhestva_2 = [para for para in combinations(fayly, 2) if fayl in para]
        nadmnozhestva_3 = [troyka for troyka in combinations(fayly, 3) if fayl in troyka]
        tseli[(fayl,)] = razmery[fayl] - chetverka - sum(tseli[troyka] for troyka in nadmnozhestva_3) - sum(tseli[para] for para in nadmnozhestva_2)
    return tseli



def rasschitat_eksklyuzivy(fayly, mnozhestva):
    schetchik = Counter()
    vse_identifikatory = set().union(*(mnozhestva[fayl] for fayl in fayly))
    for identifikator in vse_identifikatory:
        region = tuple(sorted(fayl for fayl in fayly if identifikator in mnozhestva[fayl]))
        schetchik[region] += 1
    return schetchik



def balansirovat(fayly, razmery, pary, troyki, chetverka, nachalnye_mnozhestva):
    mnozhestva = {fayl: set(mnozh) for fayl, mnozh in nachalnye_mnozhestva.items()}
    tselevye = eksklyuzivnye_tseli(fayly, razmery, pary, troyki, chetverka)

    def dobavit_v_region(region):
        identifikator = reestr.novyy()
        for fayl in region:
            mnozhestva[fayl].add(identifikator)

    def udalit_iz_regiona(region):
        identifikatory_regiona = set.intersection(*(mnozhestva[fayl] for fayl in region))
        for drugoy in fayly:
            if drugoy not in region:
                identifikatory_regiona -= mnozhestva[drugoy]
        if not identifikatory_regiona:
            return False
        zhertva = identifikatory_regiona.pop()
        if len(region) == 1:
            mnozhestva[region[0]].remove(zhertva)
            return True
        fayl_udaleniya = random.choice(region)
        mnozhestva[fayl_udaleniya].remove(zhertva)
        return True

    while True:
        tekushchie = rasschitat_eksklyuzivy(fayly, mnozhestva)
        delta = {region: tselevye[region] - tekushchie.get(region, 0) for region in tselevye}
        if all(znachenie == 0 for znachenie in delta.values()):
            break

        for region, raznitsa in sorted(delta.items(), key=lambda element: (-element[1], -len(element[0]))):
            while raznitsa > 0:
                dobavit_v_region(region)
                raznitsa -= 1

        for region, raznitsa in sorted(delta.items(), key=lambda element: (element[1], -len(element[0]))):
            while raznitsa < 0:
                uspeshno = udalit_iz_regiona(region)
                if not uspeshno:
                    zhertva = random.choice(tuple(set.intersection(*(mnozhestva[fayl] for fayl in region))))
                    for fayl in region:
                        mnozhestva[fayl].remove(zhertva)
                raznitsa += 1

        for fayl in fayly:
            while len(mnozhestva[fayl]) < razmery[fayl]:
                mnozhestva[fayl].add(reestr.novyy())
            while len(mnozhestva[fayl]) > razmery[fayl]:
                mnozhestva[fayl].pop()
    return mnozhestva



def sokhranit_mnozhestva(mnozhestva_gruppy):
    os.makedirs(KATALOG_VYHODA, exist_ok=True)
    for nomer_fayla, mnozhestvo in mnozhestva_gruppy.items():
        uporyadochennye = stroki_faylov.get(nomer_fayla, []) + sorted(mnozhestvo - set(stroki_faylov.get(nomer_fayla, [])))
        stroki_faylov[nomer_fayla] = uporyadochennye
        dataframe = pd.DataFrame([reestr.stroka(identifikator, nomer_fayla) for identifikator in uporyadochennye])[KOLONKI]
        dataframe.to_excel(os.path.join(KATALOG_VYHODA, IMENA_FAYLOV[nomer_fayla]), index=False)


def etap(fayly, razmery, pary, troyki, chetverka, perenesy=None):
    predvaritelno = {fayl: set() for fayl in fayly}
    if perenesy:
        for istochnik, priemnik, sokhranit in perenesy:
            kolichestvo_peremeshcheniya = int(len(stroki_faylov[istochnik]) * sokhranit / 100)
            for identifikator in stroki_faylov[istochnik][:kolichestvo_peremeshcheniya]:
                reestr.stroka(identifikator, priemnik, kopirovat_iz=istochnik)
                predvaritelno[priemnik].add(identifikator)
    nachalnye = balansirovat(fayly, razmery, pary, troyki, chetverka, predvaritelno)
    sokhranit_mnozhestva(nachalnye)



def glavnaya():

    etap(
        [1, 2, 3, 4],
        {1: 60, 2: 100, 3: 50, 4: 70},
        {(1,2):22,(1,3):17,(1,4):20,(2,3):19,(2,4):22,(3,4):17},
        {(1,2,3):5,(1,2,4):5,(1,3,4):5,(2,3,4):5},
        3,
    )

    etap(
        [5,6,7,8],
        {5:380,6:370,7:350,8:260},
        {(5,6):190,(5,7):190,(5,8):150,(6,7):190,(6,8):140,(7,8):120},
        {(5,6,7):70,(5,6,8):70,(5,7,8):70,(6,7,8):70},
        50,
        perenesy=[(1,5,90),(2,6,90),(3,7,94),(4,8,93)]
    )

    etap(
        [9,10,11,12],
        {9:1000,10:1150,11:1050,12:800},
        {(9,10):760,(9,11):600,(9,12):410,(10,11):750,(10,12):460,(11,12):500},
        {(9,10,11):500,(9,10,12):260,(9,11,12):250,(10,11,12):300},
        200,
        perenesy=[(5,9,90),(6,10,95),(7,11,91),(8,12,92)]
    )

    etap(
        [13,14,15,16],
        {13:1240,14:1390,15:1240,16:1190},
        {(13,14):1090,(13,15):1110,(13,16):1070,(14,15):1050,(14,16):1040,(15,16):1090},
        {(13,14,15):1020,(13,14,16):1020,(13,15,16):1040,(14,15,16):1000},
        1000,
        perenesy=[(9,13,90),(10,14,90),(11,15,90),(12,16,90)]
    )

    zapisi=[]
    for gruppa, fayly_gruppy in GRUPPY.items():
        unikalnye=len(set().union(*(set(stroki_faylov[fayl]) for fayl in fayly_gruppy)))
        zapisi.append({"Группа":gruppa,"Уникальных ID":unikalnye})
    katalog_otcheta = "otchet"
    os.makedirs(katalog_otcheta, exist_ok=True)
    pd.DataFrame(zapisi).to_csv(os.path.join(katalog_otcheta,"otchet.csv"),index=False)



import sys
import math
from typing import Dict, List, Tuple, Optional, Set


START_KATALOG = "start"
PROGRAMMY = ["ПМ", "ИВТ", "ИТСС", "ИБ"]
DATY = ["01.08", "02.08", "03.08", "04.08"]
MESTA = {"ПМ": 40, "ИВТ": 50, "ИТСС": 30, "ИБ": 20}


MIN_DOLYA_DA = 0.20

TREBUEMYE_KOLONKI = [
    "ID",
    "Согласие",
    "Приоритет",
    "Балл Физика/ИКТ",
    "Балл Русский язык",
    "Балл Математика",
    "Балл за ИД",
    "Сумма баллов",
]

ISTINA_NABOR = {"true", "1", "да", "y", "yes", "истина", "д", "ист", "TRUE", "Да"}
LOZH_NABOR = {"false", "0", "нет", "n", "no", "ложь", "н", "лож", "FALSE", "Нет"}



def v_chislo(v: object) -> Optional[int]:
    if v is None:
        return None
    try:
        s = str(v).strip()
        if s == "" or s.lower() == "nan":
            return None
        s = s.replace(" ", "").replace(",", ".")
        return int(round(float(s)))
    except Exception:
        return None


def v_log(v: object) -> Optional[bool]:
    if v is None:
        return None
    if isinstance(v, bool):
        return v
    s = str(v).strip()
    if s.lower() in ISTINA_NABOR:
        return True
    if s.lower() in LOZH_NABOR:
        return False
    return None




def put_k_xlsx(programma: str, metka_daty: str) -> str:
    return os.path.join(START_KATALOG, f"Список {programma} на {metka_daty}.xlsx")


def proverit_kolonki(df: pd.DataFrame) -> pd.DataFrame:
    for c in TREBUEMYE_KOLONKI:
        if c not in df.columns:
            df[c] = pd.NA
    df = df[TREBUEMYE_KOLONKI].copy()

    df["ID"] = df["ID"].apply(v_chislo)
    df["Приоритет"] = df["Приоритет"].apply(v_chislo)
    df["Согласие"] = df["Согласие"].apply(v_log)
    for c in ["Балл Физика/ИКТ", "Балл Русский язык", "Балл Математика", "Балл за ИД", "Сумма баллов"]:
        df[c] = df[c].apply(v_chislo)

    df = df[pd.notna(df["ID"])].copy()
    return df


def chitat_den(metka_daty: str) -> Dict[str, pd.DataFrame]:
    dannye: Dict[str, pd.DataFrame] = {}
    for prog in PROGRAMMY:
        path = put_k_xlsx(prog, metka_daty)
        if not os.path.exists(path):
            raise FileNotFoundError(f"Не найден файл: {path}")
        df = pd.read_excel(path)
        df = proverit_kolonki(df)
        df["ОП"] = prog
        dannye[prog] = df
    return dannye


def pisat_den(metka_daty: str, dannye: Dict[str, pd.DataFrame]):
    os.makedirs(START_KATALOG, exist_ok=True)
    for prog, df in dannye.items():
        path = put_k_xlsx(prog, metka_daty)
        out = df[[c for c in TREBUEMYE_KOLONKI]].copy()

        out["Согласие"] = out["Согласие"].apply(lambda v: "Да" if v_log(v) is True else ("Нет" if v_log(v) is False else pd.NA))
        out.to_excel(path, index=False)




def perezchitat_summu_stroki(row: pd.Series) -> int:
    parts = [
        v_chislo(row.get("Балл Физика/ИКТ")),
        v_chislo(row.get("Балл Русский язык")),
        v_chislo(row.get("Балл Математика")),
        v_chislo(row.get("Балл за ИД")),
    ]
    if any(v is None for v in parts):
        return v_chislo(row.get("Сумма баллов")) or 0
    return int(sum(parts))


def drugaya_summa(row: pd.Series) -> int:

    vals = [
        v_chislo(row.get("Балл Русский язык")) or 0,
        v_chislo(row.get("Балл Математика")) or 0,
        v_chislo(row.get("Балл за ИД")) or 0,
    ]
    return int(sum(vals))




def obiedinit_den(dannye: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    if not dannye:
        return pd.DataFrame(columns=TREBUEMYE_KOLONKI + ["ОП"])
    df = pd.concat(list(dannye.values()), ignore_index=True)
    # final'naya chistka/tipy
    df = df.dropna(subset=["ID"])  # na vsyakiy
    df["ID"] = df["ID"].apply(v_chislo)
    df["Приоритет"] = df["Приоритет"].apply(v_chislo)
    df["Согласие"] = df["Согласие"].apply(v_log)
    for c in ["Балл Физика/ИКТ", "Балл Русский язык", "Балл Математика", "Балл за ИД", "Сумма баллов"]:
        df[c] = df[c].apply(v_chislo)
    # pereschyot summ
    df["Сумма баллов"] = df.apply(perezchitat_summu_stroki, axis=1)
    return df




def granicy_id_dlya_dnya(dannye: Dict[str, pd.DataFrame]) -> Dict[int, Tuple[int, int]]:

    per_id_low: Dict[int, List[int]] = {}
    per_id_high: Dict[int, List[int]] = {}
    for _, df in dannye.items():
        for _, r in df.iterrows():
            iid = v_chislo(r.get("ID"))
            if iid is None:
                continue
            ot = drugaya_summa(r)
            per_id_low.setdefault(iid, []).append(ot + 55)
            per_id_high.setdefault(iid, []).append(ot + 100)
    bounds: Dict[int, Tuple[int, int]] = {}
    for iid in per_id_low:
        low = max(per_id_low[iid])
        high = min(per_id_high[iid])
        if low > high:
            low = high
        bounds[iid] = (int(low), int(high))
    return bounds




def unificirovat_summu_po_id_v_den(dannye: Dict[str, pd.DataFrame]):
    id_gruppy: Dict[int, List[Tuple[str, int]]] = {}
    for prog, df in dannye.items():
        for _, r in df.iterrows():
            iid = v_chislo(r["ID"])
            if iid is None:
                continue
            ot = drugaya_summa(r)
            id_gruppy.setdefault(iid, []).append((prog, ot))
    for iid, lst in id_gruppy.items():
        low = max(ot + 55 for _, ot in lst)
        high = min(ot + 100 for _, ot in lst)
        target_sum = low if low <= high else high
        for prog, ot in lst:
            df = dannye[prog]
            new_phys = int(max(55, min(100, target_sum - ot)))
            mask = df["ID"] == iid
            df.loc[mask, "Балл Физика/ИКТ"] = new_phys
            df.loc[mask, "Сумма баллов"] = int(ot + new_phys)
            dannye[prog] = df




def smesat_soglasie_v_den(dannye: Dict[str, pd.DataFrame], metka_daty: str):

    prog2ids: Dict[str, Set[int]] = {p: set(dannye[p]["ID"].dropna().astype(int).tolist()) for p in PROGRAMMY}
    id2progs: Dict[int, Set[str]] = {}
    for p in PROGRAMMY:
        for iid in prog2ids[p]:
            id2progs.setdefault(iid, set()).add(p)
    total_ids = set(id2progs.keys())

    def apply(consent_by_id: Dict[int, bool]):
        for p in PROGRAMMY:
            df = dannye[p]
            if df.empty:
                continue
            ids_series = df["ID"].dropna().astype(int)
            mask = df["ID"].notna()
            df.loc[mask, "Согласие"] = ids_series.map(lambda x: bool(consent_by_id.get(int(x), False))).values
            dannye[p] = df

    def ceil(x: float) -> int:
        return int(math.ceil(x - 1e-12))


    def build_by_targets(target_yes: Dict[str, int]) -> Dict[int, bool]:
        consent_by_id = {iid: False for iid in total_ids}
        true_counts = {p: 0 for p in PROGRAMMY}
        candidates = list(total_ids)
        while True:
            deficit = {p: max(0, target_yes[p] - true_counts[p]) for p in PROGRAMMY}
            if sum(deficit.values()) <= 0:
                break
            best_iid = None
            best_gain = 0
            for iid in candidates:
                progs_here = id2progs.get(iid, set())
                gain = sum(1 for pp in progs_here if deficit.get(pp, 0) > 0)
                if gain > best_gain:
                    best_gain = gain
                    best_iid = iid
            if best_iid is None or best_gain == 0:

                break
            consent_by_id[best_iid] = True
            for pp in id2progs.get(best_iid, set()):
                true_counts[pp] += 1
            candidates.remove(best_iid)

        for p in PROGRAMMY:
            need = max(0, target_yes[p] - true_counts[p])
            if need <= 0:
                continue
            for iid in list(total_ids):
                if need <= 0:
                    break
                if consent_by_id[iid]:
                    continue
                if p in id2progs.get(iid, set()):
                    consent_by_id[iid] = True
                    for pp in id2progs.get(iid, set()):
                        true_counts[pp] += 1
                    need -= 1
        return consent_by_id


    if metka_daty == "01.08":
        limits_max_yes = {p: max(0, MESTA[p] - 1) for p in PROGRAMMY}

        target_yes = {}
        for p in PROGRAMMY:
            n = len(prog2ids[p])
            min_yes = ceil(n * MIN_DOLYA_DA)
            cap = limits_max_yes[p]
            if n > 0 and cap == 0:

                target_yes[p] = 0
            else:

                target_yes[p] = min(cap, max(0, min_yes))
                target_yes[p] = min(target_yes[p], max(0, n - 1))
        consent = build_by_targets(target_yes)

        for p in PROGRAMMY:
            if not prog2ids[p]:
                continue
            yes_here = any(consent.get(iid, False) for iid in prog2ids[p])
            if not yes_here:

                iid = next(iter(prog2ids[p]))
                consent[iid] = True
        apply(consent)
        return


    target_yes: Dict[str, int] = {}
    for p in PROGRAMMY:
        n = len(prog2ids[p])

        min_yes = ceil(n * MIN_DOLYA_DA)
        min_yes = max(min_yes, 1)
        min_yes = max(min_yes, MESTA[p])

        if n > 0:
            min_yes = min(min_yes, n - 1)
        target_yes[p] = max(0, min_yes)

    consent = build_by_targets(target_yes)


    for p in PROGRAMMY:
        ids = prog2ids[p]
        if not ids:
            continue
        yes_count = sum(1 for iid in ids if consent.get(iid, False))
        no_count = len(ids) - yes_count
        if yes_count == 0 and len(ids) > 0:
            consent[next(iter(ids))] = True
            yes_count += 1
            no_count -= 1
        if no_count == 0 and len(ids) > 0:

            for iid in ids:
                if consent.get(iid, False):
                    consent[iid] = False
                    break

    apply(consent)




def naznachit_prioritety_bez_povtorov(dannye: Dict[str, pd.DataFrame], vybrannye_lvl1: Dict[str, Set[int]]):
    id_k_programmam: Dict[int, List[str]] = {}
    for prog, df in dannye.items():
        for iid in df["ID"].dropna().astype(int).tolist():
            id_k_programmam.setdefault(iid, []).append(prog)
    for iid, progs in id_k_programmam.items():
        ispolzovan: Set[int] = set()

        for prog in progs:
            if iid in vybrannye_lvl1.get(prog, set()) and 1 not in ispolzovan:
                df = dannye[prog]
                df.loc[df["ID"] == iid, "Приоритет"] = 1
                ispolzovan.add(1)
                dannye[prog] = df

        pool = [2, 3, 4, 1]
        for prog in progs:
            df = dannye[prog]
            s = df.loc[df["ID"] == iid, "Приоритет"]
            cur = v_chislo(s.iloc[0]) if not s.empty else None
            if cur in (1, 2, 3, 4):
                ispolzovan.add(cur)
                continue
            pr = next((x for x in pool if x not in ispolzovan), 1)
            df.loc[df["ID"] == iid, "Приоритет"] = pr
            ispolzovan.add(pr)
            dannye[prog] = df




def vybrat_prioritet1_bez_peresecheniy_adv(
    dannye: Dict[str, pd.DataFrame],
    strategiya: List[Tuple[str, str]],
) -> Dict[str, Set[int]]:

    df_day = obiedinit_den(dannye)
    bounds = granicy_id_dlya_dnya(dannye)
    globalno_vybrany: Set[int] = set()
    vybrannye: Dict[str, Set[int]] = {p: set() for p in PROGRAMMY}

    for prog, mode in strategiya:
        dfp = df_day[df_day["ОП"] == prog].copy()
        ids = dfp["ID"].dropna().astype(int).tolist()
        if mode in ("desc", "asc"):
            asc = (mode == "asc")
            dfp = dfp.sort_values(
                by=["Сумма баллов", "Балл Математика", "Балл Физика/ИКТ", "Балл Русский язык", "Балл за ИД"],
                ascending=[asc, asc, asc, asc, asc]
            )
            ids_sorted = dfp["ID"].dropna().astype(int).tolist()
        elif mode == "asc_lb":
            ids_sorted = sorted(ids, key=lambda i: bounds.get(i, (10**9, -10**9))[0])
        elif mode == "desc_lb":
            ids_sorted = sorted(ids, key=lambda i: bounds.get(i, (-10**9, 10**9))[0], reverse=True)
        elif mode == "desc_hb":
            ids_sorted = sorted(ids, key=lambda i: bounds.get(i, (0, 0))[1], reverse=True)
        else:
            ids_sorted = ids
        seats = MESTA[prog]
        pick: List[int] = []
        for iid in ids_sorted:
            if iid in globalno_vybrany:
                continue
            pick.append(iid)
            if len(pick) >= seats:
                break
        vybrannye[prog] = set(pick)
        globalno_vybrany.update(pick)
    return vybrannye




def rasschitat_prokhodnye(df: pd.DataFrame) -> Tuple[Dict[str, List[Dict]], Dict[str, Optional[int]]]:
    base = df[(df["Согласие"] == True) & (pd.notna(df["Сумма баллов"]))].copy()
    if base.empty:
        results = {p: [] for p in PROGRAMMY}
        pscores = {p: None for p in PROGRAMMY}
        return results, pscores
    ostalos: Dict[str, int] = {p: MESTA[p] for p in PROGRAMMY}
    nazna: Set[int] = set()
    results: Dict[str, List[Dict]] = {p: [] for p in PROGRAMMY}
    for uroven in range(1, 5):
        for prog in PROGRAMMY:
            seats = ostalos[prog]
            if seats <= 0:
                continue
            cand = base[(base["ОП"] == prog) & (base["Приоритет"] == uroven) & (~base["ID"].isin(list(nazna)))].copy()
            if cand.empty:
                continue
            cand = cand.sort_values(by=["Сумма баллов", "Балл Математика", "Балл Физика/ИКТ", "Балл Русский язык", "Балл за ИД"], ascending=[False, False, False, False, False])
            take = min(seats, len(cand))
            sel = cand.head(take)
            for _, r in sel.iterrows():
                results[prog].append({
                    "ID": int(r["ID"]),
                    "Сумма баллов": int(r["Сумма баллов"]),
                    "Приоритет": int(r["Приоритет"]) if pd.notna(r["Приоритет"]) else None,
                })
                nazna.add(int(r["ID"]))
                ostalos[prog] -= 1
    pscores: Dict[str, Optional[int]] = {}
    for prog in PROGRAMMY:
        seats_total = MESTA[prog]
        lst = results[prog]
        if len(lst) < seats_total:
            pscores[prog] = None
        else:
            lst_sorted = sorted(lst, key=lambda x: (-x["Сумма баллов"], x["Приоритет"]))
            pscores[prog] = int(lst_sorted[seats_total - 1]["Сумма баллов"]) if seats_total - 1 < len(lst_sorted) else None
    return results, pscores




def podnyat_min_sum_po_fizike(
    dannye_dnya: Dict[str, pd.DataFrame],
    programma: str,
    zachislennye_ids: List[int],
    tekushchiy_min_sum: int,
    tsel_min_sum: int,
) -> int:
    if not zachislennye_ids:
        return tekushchiy_min_sum
    dfp = dannye_dnya[programma]
    df_sel = dfp[dfp["ID"].isin(zachislennye_ids)].copy()
    if df_sel.empty:
        return tekushchiy_min_sum
    df_sel = df_sel.sort_values(by=["Сумма баллов", "Балл Математика", "Балл Физика/ИКТ", "Балл Русский язык", "Балл за ИД"], ascending=[True, True, True, True, True])
    min_row = df_sel.iloc[0]
    iid = int(min_row["ID"]) if pd.notna(min_row["ID"]) else None
    if iid is None:
        return tekushchiy_min_sum
    progs_with_id = []
    for prog, d in dannye_dnya.items():
        if iid in d["ID"].dropna().astype(int).tolist():
            progs_with_id.append(prog)
    low_bounds, high_bounds = [], []
    for prog in progs_with_id:
        row = dannye_dnya[prog].loc[dannye_dnya[prog]["ID"] == iid].iloc[0]
        ot = drugaya_summa(row)
        low_bounds.append(ot + 55)
        high_bounds.append(ot + 100)
    low, high = max(low_bounds), min(high_bounds)
    desired = tsel_min_sum if low <= high else high
    desired = max(min(desired, high), low)
    for prog in progs_with_id:
        df = dannye_dnya[prog]
        row = df.loc[df["ID"] == iid].iloc[0]
        ot = drugaya_summa(row)
        new_phys = int(max(55, min(100, desired - ot)))
        df.loc[df["ID"] == iid, "Балл Физика/ИКТ"] = new_phys
        df.loc[df["ID"] == iid, "Сумма баллов"] = int(ot + new_phys)
        dannye_dnya[prog] = df
    dfp2 = dannye_dnya[programma]
    df_sel2 = dfp2[dfp2["ID"].isin(zachislennye_ids)].copy()
    if df_sel2.empty:
        return tekushchiy_min_sum
    df_sel2 = df_sel2.sort_values(by=["Сумма баллов", "Балл Математика", "Балл Физика/ИКТ", "Балл Русский язык", "Балл за ИД"], ascending=[True, True, True, True, True])
    return int(df_sel2.iloc[0]["Сумма баллов"]) if pd.notna(df_sel2.iloc[0]["Сумма баллов"]) else tekushchiy_min_sum


def ponizit_poros_do(dannye_dnya: Dict[str, pd.DataFrame], programma: str, tsel_min_sum: int, max_iters: int = 24):
    for _ in range(max_iters):
        d = obiedinit_den(dannye_dnya)
        results, ps = rasschitat_prokhodnye(d)
        cur = ps.get(programma)
        if cur is None or cur <= tsel_min_sum:
            return
        zachislennye_ids = [x["ID"] for x in results.get(programma, [])]
        podnyat_min_sum_po_fizike(dannye_dnya, programma, zachislennye_ids, cur, tsel_min_sum)




def podnyat_den02_nad_den03_silno(den2: Dict[str, pd.DataFrame], ps03: Dict[str, Optional[int]], margin: int = 3):
    bounds2 = granicy_id_dlya_dnya(den2)
    df2 = obiedinit_den(den2)
    globalno_vybrany: Set[int] = set()
    vybrannye: Dict[str, Set[int]] = {p: set() for p in PROGRAMMY}

    for prog in ["ИТСС", "ИБ", "ПМ", "ИВТ"]:
        dfp = df2[df2["ОП"] == prog].copy()
        ids = dfp["ID"].dropna().astype(int).tolist()
        seats = MESTA[prog]
        pick: List[int] = []
        if prog in ("ИТСС", "ИБ"):
            target = (ps03.get(prog) or 0) + margin
            ids_sorted = sorted(ids, key=lambda i: bounds2.get(i, (0, 0))[1], reverse=True)
            for iid in ids_sorted:
                if iid in globalno_vybrany:
                    continue
                hb = bounds2.get(iid, (0, 0))[1]
                if hb >= target:
                    pick.append(iid)
                    if len(pick) >= seats:
                        break
            if len(pick) < seats:
                for iid in ids_sorted:
                    if iid in globalno_vybrany or iid in pick:
                        continue
                    pick.append(iid)
                    if len(pick) >= seats:
                        break
        else:
            dfp_sorted = dfp.sort_values(by=["Сумма баллов", "Балл Математика", "Балл Физика/ИКТ", "Балл Русский язык", "Балл за ИД"], ascending=[False, False, False, False, False])
            for iid in dfp_sorted["ID"].dropna().astype(int).tolist():
                if iid in globalno_vybrany:
                    continue
                pick.append(iid)
                if len(pick) >= seats:
                    break
        vybrannye[prog] = set(pick)
        globalno_vybrany.update(pick)
    naznachit_prioritety_bez_povtorov(den2, vybrannye)

    d2 = obiedinit_den(den2)
    results2, ps02 = rasschitat_prokhodnye(d2)
    for prog in ["ИТСС", "ИБ"]:
        if ps03.get(prog) is None:
            continue
        target = ps03[prog] + margin
        iters = 0
        while ps02.get(prog) is not None and ps02[prog] < target and iters < 40:
            zachislennye_ids = [x["ID"] for x in results2.get(prog, [])]
            cur = ps02[prog]
            podnyat_min_sum_po_fizike(den2, prog, zachislennye_ids, cur, target)
            d2 = obiedinit_den(den2)
            results2, ps02 = rasschitat_prokhodnye(d2)
            iters += 1

    d2 = obiedinit_den(den2)
    results2, ps02 = rasschitat_prokhodnye(d2)
    for prog in ["ИТСС", "ИБ"]:
        if ps03.get(prog) is None:
            continue
        iters = 0
        while ps02.get(prog) is not None and ps02[prog] <= ps03[prog] and iters < 40:
            zachislennye_ids = [x["ID"] for x in results2.get(prog, [])]
            cur = ps02[prog]
            podnyat_min_sum_po_fizike(den2, prog, zachislennye_ids, cur, ps03[prog] + 1)
            d2 = obiedinit_den(den2)
            results2, ps02 = rasschitat_prokhodnye(d2)
            iters += 1




def obrabotat_den_01(dannye: Dict[str, pd.DataFrame]):
    smesat_soglasie_v_den(dannye, "01.08")
    unificirovat_summu_po_id_v_den(dannye)

    naznachit_prioritety_bez_povtorov(dannye, {p: set() for p in PROGRAMMY})


def obrabotat_den_02(dannye: Dict[str, pd.DataFrame]):
    smesat_soglasie_v_den(dannye, "02.08")
    unificirovat_summu_po_id_v_den(dannye)

    vybrannye_lvl1 = vybrat_prioritet1_bez_peresecheniy_adv(dannye, [("ПМ", 'desc'), ("ИВТ", 'desc'), ("ИТСС", 'desc'), ("ИБ", 'desc')])
    naznachit_prioritety_bez_povtorov(dannye, vybrannye_lvl1)


def obrabotat_den_03(dannye: Dict[str, pd.DataFrame], ps_den02: Dict[str, Optional[int]]):
    smesat_soglasie_v_den(dannye, "03.08")
    unificirovat_summu_po_id_v_den(dannye)

    vybrannye_lvl1 = vybrat_prioritet1_bez_peresecheniy_adv(
        dannye,
        [("ИТСС", 'asc_lb'), ("ИБ", 'asc_lb'), ("ПМ", 'desc'), ("ИВТ", 'desc')]
    )
    naznachit_prioritety_bez_povtorov(dannye, vybrannye_lvl1)

    d = obiedinit_den(dannye)
    results, ps = rasschitat_prokhodnye(d)
    for prog, inc in {"ПМ": 6, "ИВТ": 5}.items():
        if ps.get(prog) is None or ps_den02.get(prog) is None:
            continue
        cur = ps[prog]
        target = max(cur, ps_den02[prog] + inc)
        zachislennye_ids = [x["ID"] for x in results.get(prog, [])]
        podnyat_min_sum_po_fizike(dannye, prog, zachislennye_ids, cur, target)
        d = obiedinit_den(dannye)
        results, ps = rasschitat_prokhodnye(d)
    for prog, dec in {"ИТСС": 5, "ИБ": 4}.items():
        if ps_den02.get(prog) is None:
            continue
        target = max(0, ps_den02[prog] - dec)
        ponizit_poros_do(dannye, prog, target, max_iters=24)
        d = obiedinit_den(dannye)
        results, ps = rasschitat_prokhodnye(d)
    return ps


def obrabotat_den_04(dannye: Dict[str, pd.DataFrame], ps_den03: Dict[str, Optional[int]]):
    smesat_soglasie_v_den(dannye, "04.08")
    unificirovat_summu_po_id_v_den(dannye)

    vybrannye_lvl1 = vybrat_prioritet1_bez_peresecheniy_adv(dannye, [("ПМ", 'desc'), ("ИБ", 'desc'), ("ИВТ", 'desc'), ("ИТСС", 'desc')])
    naznachit_prioritety_bez_povtorov(dannye, vybrannye_lvl1)

    def perechitat():
        d = obiedinit_den(dannye)
        res, p = rasschitat_prokhodnye(d)
        return res, p

    res, p = perechitat()
    base_inc = {"ПМ": 8, "ИБ": 6, "ИВТ": 5, "ИТСС": 3}
    for prog in PROGRAMMY:
        cur = p.get(prog)
        prev = ps_den03.get(prog)
        if cur is None or prev is None:
            continue
        target = max(cur, prev + base_inc[prog])
        zachislennye_ids = [x["ID"] for x in res.get(prog, [])]
        podnyat_min_sum_po_fizike(dannye, prog, zachislennye_ids, cur, target)
        res, p = perechitat()

    max_iters = 60
    for _ in range(max_iters):
        res, p = perechitat()
        growth_ok = True
        for prog in PROGRAMMY:
            if p.get(prog) is None or ps_den03.get(prog) is None or p[prog] <= ps_den03[prog]:
                growth_ok = False
                break
        pm = p.get("ПМ") or 0
        ib = p.get("ИБ") or 0
        ivt = p.get("ИВТ") or 0
        itss = p.get("ИТСС") or 0
        order_ok = (pm > ib > ivt > itss)
        if growth_ok and order_ok:
            break
        changed = False

        for prog in PROGRAMMY:
            if p.get(prog) is None or ps_den03.get(prog) is None:
                continue
            if p[prog] <= ps_den03[prog]:
                zachislennye_ids = [x["ID"] for x in res.get(prog, [])]
                podnyat_min_sum_po_fizike(dannye, prog, zachislennye_ids, p[prog], ps_den03[prog] + 2)
                changed = True

        res, p = perechitat()
        pm = p.get("ПМ") or 0
        ib = p.get("ИБ") or 0
        ivt = p.get("ИВТ") or 0
        itss = p.get("ИТСС") or 0
        if not (pm > ib):
            zachislennye_ids = [x["ID"] for x in res.get("ПМ", [])]
            podnyat_min_sum_po_fizike(dannye, "ПМ", zachislennye_ids, pm, ib + 1)
            changed = True
        res, p = perechitat()
        pm = p.get("ПМ") or 0
        ib = p.get("ИБ") or 0
        ivt = p.get("ИВТ") or 0
        if not (ib > ivt):
            zachislennye_ids = [x["ID"] for x in res.get("ИБ", [])]
            podnyat_min_sum_po_fizike(dannye, "ИБ", zachislennye_ids, ib, ivt + 1)
            changed = True
        res, p = perechitat()
        ib = p.get("ИБ") or 0
        ivt = p.get("ИВТ") or 0
        itss = p.get("ИТСС") or 0
        if not (ivt > itss):
            zachislennye_ids = [x["ID"] for x in res.get("ИВТ", [])]
            podnyat_min_sum_po_fizike(dannye, "ИВТ", zachislennye_ids, ivt, itss + 1)
            changed = True
        if not changed:
            break

    res, p = perechitat()
    return p




def glavnaya_zadanie():

    otsutstvuyut = []
    for metka_daty in DATY:
        for prog in PROGRAMMY:
            path = put_k_xlsx(prog, metka_daty)
            if not os.path.exists(path):
                otsutstvuyut.append(path)
    if otsutstvuyut:
        print("Не найдены файлы:")
        for pth in otsutstvuyut:
            print(" -", pth)
        sys.exit(1)


    dannye_po_dnyam: Dict[str, Dict[str, pd.DataFrame]] = {}
    for d in DATY:
        dannye_po_dnyam[d] = chitat_den(d)


    obrabotat_den_01(dannye_po_dnyam["01.08"])
    pisat_den("01.08", dannye_po_dnyam["01.08"])
    df01 = obiedinit_den(dannye_po_dnyam["01.08"])
    _, ps01 = rasschitat_prokhodnye(df01)

    # 02.08
    obrabotat_den_02(dannye_po_dnyam["02.08"])
    pisat_den("02.08", dannye_po_dnyam["02.08"])
    df02 = obiedinit_den(dannye_po_dnyam["02.08"])
    _, ps02 = rasschitat_prokhodnye(df02)

    # 03.08
    ps03 = obrabotat_den_03(dannye_po_dnyam["03.08"], ps02)
    pisat_den("03.08", dannye_po_dnyam["03.08"])

    # Post-korrektsiya 02.08, esli nuzhno, chtoby 03.08 < 02.08 u ИТСС/ИБ
    podnyat_den02_nad_den03_silno(dannye_po_dnyam["02.08"], ps03, margin=3)
    pisat_den("02.08", dannye_po_dnyam["02.08"])

    df02b = obiedinit_den(dannye_po_dnyam["02.08"])
    _, ps02 = rasschitat_prokhodnye(df02b)


    ps04 = obrabotat_den_04(dannye_po_dnyam["04.08"], ps03)
    pisat_den("04.08", dannye_po_dnyam["04.08"])


    print("Проходные баллы по дням:")
    def fmt(ps: Dict[str, Optional[int]]) -> str:
        return ", ".join([f"{p}: {('НЕДОБОР' if ps.get(p) is None else ps.get(p))}" for p in PROGRAMMY])

    print(f"01.08 — {fmt(ps01)}")
    print(f"02.08 — {fmt(ps02)}")
    print(f"03.08 — {fmt(ps03)}")
    print(f"04.08 — {fmt(ps04)}")

    rows = []
    rows.append({"Дата": "01.08", **{p: ("НЕДОБОР" if ps01.get(p) is None else str(ps01.get(p))) for p in PROGRAMMY}})
    rows.append({"Дата": "02.08", **{p: ("НЕДОБОР" if ps02.get(p) is None else str(ps02.get(p))) for p in PROGRAMMY}})
    rows.append({"Дата": "03.08", **{p: ("НЕДОБОР" if ps03.get(p) is None else str(ps03.get(p))) for p in PROGRAMMY}})
    rows.append({"Дата": "04.08", **{p: ("НЕДОБОР" if ps04.get(p) is None else str(ps04.get(p))) for p in PROGRAMMY}})
    out_path = "history_test.csv"
    pd.DataFrame(rows, columns=["Дата"] + PROGRAMMY).to_csv(out_path, index=False, encoding="utf-8-sig")
    print(f"История проходных сохранена в {out_path}")



if __name__ == "__main__":

    glavnaya()

    glavnaya_zadanie()
