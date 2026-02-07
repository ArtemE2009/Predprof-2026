

import os
import sys
import csv
import subprocess
from datetime import datetime
from typing import Dict, List, Optional, Tuple
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import tkinter.font as tkfont


try:
    import pandas as pd
except Exception:
    pd = None

try:
    from PIL import Image, ImageTk
    PIL_DOSTUPEN = True
except Exception:
    PIL_DOSTUPEN = False

MATPLOT_DOSTUPEN = True
try:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
except Exception:
    MATPLOT_DOSTUPEN = False

REPORTLAB_DOSTUPEN = True
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas as pdfcanvas
    from reportlab.lib.units import cm
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
except Exception:
    REPORTLAB_DOSTUPEN = False


ZAGOLOVOK_PRILOZHENIYA = "Конкурсные списки"

FAYL_BD = "абитуриенты.csv"
FAYL_ISTORII = "history.csv"
FAYL_ZAPUSKA = "zapusk.csv"
KATALOG_SPISKOV = "spiski"
KATALOG_OTCHETOV = "report"
KATALOG_SPISKOV_ZACHISLENIYA = "lists"
KATALOG_ANALIZA = "otchet"
KARTINKA_FONA_PO_UMOLCHANIYU = "1.png"
KATALOG_VHODNYKH_FAYLOV = "start"

FAYL_KONFIG_PROGRAMM = "programs.csv"
FAYL_KONFIG_DAT = "date_list.csv"

PROGRAMMY_PO_UMOLCHANIYU = ["ПМ", "ИВТ", "ИТСС", "ИБ"]
MESTA_PO_UMOLCHANIYU = {"ПМ": 40, "ИВТ": 50, "ИТСС": 30, "ИБ": 20}
DATY_PO_UMOLCHANIYU = ["01.08", "02.08", "03.08", "04.08"]


def zagruzit_konfig_programm() -> Tuple[List[str], Dict[str, int]]:
    if not os.path.exists(FAYL_KONFIG_PROGRAMM):
        return PROGRAMMY_PO_UMOLCHANIYU.copy(), MESTA_PO_UMOLCHANIYU.copy()
    programmy, mesta = [], {}
    try:
        with open(FAYL_KONFIG_PROGRAMM, "r", encoding="utf-8-sig", newline="") as f:
            chitatel = csv.DictReader(f)
            for stroka in chitatel:
                prog = stroka.get("ОП", "").strip()
                mest = stroka.get("Количество мест", "").strip()
                if prog:
                    try:
                        mesta[prog] = int(mest)
                        programmy.append(prog)
                    except Exception:
                        pass
    except Exception:
        return PROGRAMMY_PO_UMOLCHANIYU.copy(), MESTA_PO_UMOLCHANIYU.copy()
    return (programmy if programmy else PROGRAMMY_PO_UMOLCHANIYU.copy(), 
            mesta if programmy else MESTA_PO_UMOLCHANIYU.copy())


def sokhranit_konfig_programm(programmy: List[str], mesta: Dict[str, int]):
    with open(FAYL_KONFIG_PROGRAMM, "w", encoding="utf-8-sig", newline="") as f:
        pisatel = csv.writer(f)
        pisatel.writerow(["ОП", "Количество мест"])
        for prog in programmy:
            pisatel.writerow([prog, mesta.get(prog, "")])


def zagruzit_konfig_dat() -> List[str]:
    if not os.path.exists(FAYL_KONFIG_DAT):
        return DATY_PO_UMOLCHANIYU.copy()
    daty: List[str] = []
    try:
        with open(FAYL_KONFIG_DAT, "r", encoding="utf-8-sig") as f:
            for stroka in f:
                stroka = stroka.strip()
                if stroka:
                    daty.append(stroka)
    except Exception:
        return DATY_PO_UMOLCHANIYU.copy()
    return daty if daty else DATY_PO_UMOLCHANIYU.copy()


def sokhranit_konfig_dat(daty: List[str]):
    with open(FAYL_KONFIG_DAT, "w", encoding="utf-8-sig") as f:
        for data in daty:
            f.write(data + "\n")


PROGRAMMY, MESTA = zagruzit_konfig_programm()
SPISOK_DAT = zagruzit_konfig_dat()


KOLONKI_BD = [
    "ID",
    "Согласие",
    "Приоритет",
    "Балл Физика/ИКТ",
    "Балл Русский язык",
    "Балл Математика",
    "Балл за ИД",
    "Сумма баллов",
    "ОП",
]
KOLONKI_PROSMOTRA = KOLONKI_BD.copy()

SINONIMY_ZAGOLOVKOV = {
    "id": "ID",
    "ид": "ID",
    "идентификатор": "ID",
    "согласие": "Согласие",
    "наличие согласия": "Согласие",
    "согласие о зачислении": "Согласие",
    "приоритет": "Приоритет",
    "физика": "Балл Физика/ИКТ",
    "физика/икт": "Балл Физика/ИКТ",
    "икт": "Балл Физика/ИКТ",
    "русский": "Балл Русский язык",
    "русский язык": "Балл Русский язык",
    "математика": "Балл Математика",
    "индивидуальные достижения": "Балл за ИД",
    "индивидуальные_достижения": "Балл за ИД",
    "сумма": "Сумма баллов",
    "сумма баллов": "Сумма баллов",
}

MNOZHESTVO_ISTINA = {"true", "1", "да", "y", "yes", "истина", "д", "ист"}
MNOZHESTVO_LOZH = {"false", "0", "нет", "n", "no", "ложь", "н", "лож"}


def proverit_pandas():
    if pd is None:
        messagebox.showerror("Зависимость", "Требуется pandas (pip install pandas openpyxl)")
        return False
    return True


def proverit_matplotlib():
    if not MATPLOT_DOSTUPEN:
        messagebox.showerror("Зависимость", "Требуется matplotlib (pip install matplotlib)")
        return False
    return True


def proverit_reportlab():
    if not REPORTLAB_DOSTUPEN:
        messagebox.showerror("Зависимость", "Требуется reportlab (pip install reportlab)")
        return False
    return True


def zaregistrirovat_kirillicheskiy_shrift() -> str:
    imya_shrifta = "Helvetica"
    try:
        arial_windows = os.path.join(os.environ.get("WINDIR", "C:/Windows"), "Fonts", "arial.ttf")
        if os.path.exists(arial_windows):
            pdfmetrics.registerFont(TTFont("Arial", arial_windows))
            return "Arial"
        dejavu_lokalnyy = "DejaVuSans.ttf"
        if os.path.exists(dejavu_lokalnyy):
            pdfmetrics.registerFont(TTFont("DejaVuSans", dejavu_lokalnyy))
            return "DejaVuSans"
    except Exception:
        pass
    return imya_shrifta


def v_bool(znachenie):
    if znachenie is None:
        return None
    if isinstance(znachenie, bool):
        return znachenie
    stroka = str(znachenie).strip().lower()
    if stroka in MNOZHESTVO_ISTINA:
        return True
    if stroka in MNOZHESTVO_LOZH:
        return False
    try:
        if stroka.replace(".", "", 1).isdigit():
            return float(stroka) != 0.0
    except Exception:
        pass
    return None


def v_tseloe(znachenie):
    if znachenie is None:
        return None
    if isinstance(znachenie, int):
        return znachenie
    try:
        stroka = str(znachenie).strip().replace(" ", "")
        if stroka == "" or stroka.lower() == "nan":
            return None
        stroka2 = stroka.replace(",", ".")
        return int(round(float(stroka2)))
    except Exception:
        return None


def ogranichit_prioritet(znachenie: Optional[int]):
    if znachenie is None:
        return None
    return znachenie if 1 <= znachenie <= 4 else None


def pereschitat_summu(stroka: Dict):
    klyuchi = ["Балл Физика/ИКТ", "Балл Русский язык", "Балл Математика", "Балл за ИД"]
    znacheniya = []
    for klyuch in klyuchi:
        znach = v_tseloe(stroka.get(klyuch))
        if znach is None:
            return None
        znacheniya.append(znach)
    return sum(znacheniya)


def normalizovat_zagolovki(df: "pd.DataFrame") -> "pd.DataFrame":
    slovar_pereimenovaniya = {}
    for kolonka in df.columns:
        klyuch = str(kolonka).strip()
        klyuch_nizh = klyuch.lower()
        if klyuch in KOLONKI_BD:
            slovar_pereimenovaniya[kolonka] = klyuch
        elif klyuch_nizh in SINONIMY_ZAGOLOVKOV:
            slovar_pereimenovaniya[kolonka] = SINONIMY_ZAGOLOVKOV[klyuch_nizh]
    df = df.rename(columns=slovar_pereimenovaniya)
    for kol in [k for k in KOLONKI_BD if k != "ОП"]:
        if kol not in df.columns:
            df[kol] = None
    df = df[[k for k in KOLONKI_BD if k != "ОП"]]
    zapisi = []
    for _, strok in df.iterrows():
        stroka = dict(strok)
        stroka["ID"] = v_tseloe(stroka.get("ID"))
        stroka["Согласие"] = v_bool(stroka.get("Согласие"))
        stroka["Приоритет"] = ogranichit_prioritet(v_tseloe(stroka.get("Приоритет")))
        stroka["Балл Физика/ИКТ"] = v_tseloe(stroka.get("Балл Физика/ИКТ"))
        stroka["Балл Русский язык"] = v_tseloe(stroka.get("Балл Русский язык"))
        stroka["Балл Математика"] = v_tseloe(stroka.get("Балл Математика"))
        stroka["Балл за ИД"] = v_tseloe(stroka.get("Балл за ИД"))
        summa = pereschitat_summu(stroka)
        stroka["Сумма баллов"] = summa if summa is not None else v_tseloe(stroka.get("Сумма баллов"))
        if stroka["ID"] is not None:
            zapisi.append(stroka)
    vyhod = pd.DataFrame(zapisi, columns=[k for k in KOLONKI_BD if k != "ОП"])
    if not vyhod.empty:
        vyhod = vyhod.drop_duplicates(subset=["ID"], keep="last")
    vyhod.reset_index(drop=True, inplace=True)
    return vyhod


def primenit_tipy(df: "pd.DataFrame") -> "pd.DataFrame":
    vyhod = df.copy()
    for kol in KOLONKI_BD:
        if kol not in vyhod.columns:
            vyhod[kol] = pd.NA
    try:
        vyhod["ID"] = vyhod["ID"].astype("Int64")
        vyhod["Приоритет"] = vyhod["Приоритет"].astype("Int64")
        for klyuch in ["Балл Физика/ИКТ", "Балл Русский язык", "Балл Математика", "Балл за ИД", "Сумма баллов"]:
            vyhod[klyuch] = vyhod[klyuch].astype("Int64")
        vyhod["Согласие"] = vyhod["Согласие"].astype("boolean")
        vyhod["ОП"] = vyhod["ОП"].astype("string")
    except Exception:
        pass
    return vyhod[KOLONKI_BD]


def prochitat_bd():
    if not os.path.exists(FAYL_BD):
        return pd.DataFrame(columns=KOLONKI_BD)
    try:
        df = pd.read_csv(FAYL_BD, encoding="utf-8-sig")
    except Exception:
        return pd.DataFrame(columns=KOLONKI_BD)
    for kol in KOLONKI_BD:
        if kol not in df.columns:
            df[kol] = pd.NA
    df = df[KOLONKI_BD]
    if not df.empty:
        df["ID"] = df["ID"].apply(v_tseloe)
        df["Согласие"] = df["Согласие"].apply(v_bool)
        df["Приоритет"] = df["Приоритет"].apply(lambda x: ogranichit_prioritet(v_tseloe(x)))
        for klyuch in ["Балл Физика/ИКТ", "Балл Русский язык", "Балл Математика", "Балл за ИД", "Сумма баллов"]:
            df[klyuch] = df[klyuch].apply(v_tseloe)
    return primenit_tipy(df)


def zapisat_bd(df: "pd.DataFrame"):
    primenit_tipy(df).to_csv(FAYL_BD, index=False, encoding="utf-8-sig")

def obnovit_programmu(bazovaya_bd: "pd.DataFrame", df_novyy: "pd.DataFrame", programma: str):
    novyy_df = df_novyy.copy()
    novyy_df["ОП"] = programma
    novyy_df = novyy_df[KOLONKI_BD]

    staraya_prog = bazovaya_bd[bazovaya_bd["ОП"] == programma].copy() if not bazovaya_bd.empty else pd.DataFrame(columns=KOLONKI_BD)

    staraya_karta = {v_tseloe(strok["ID"]): strok for _, strok in staraya_prog.iterrows()}
    novaya_karta = {v_tseloe(strok["ID"]): strok for _, strok in novyy_df.iterrows()}

    starye_id = set([i for i in staraya_karta.keys() if i is not None])
    novye_id = set([i for i in novaya_karta.keys() if i is not None])

    dobavleno = len(novye_id - starye_id)
    udaleno = len(starye_id - novye_id)

    obnovleno = 0
    peresechenie = starye_id & novye_id
    kolonki_sravneniya = [k for k in KOLONKI_BD]
    for ident in peresechenie:
        staroe = staraya_karta[ident]
        novoe = novaya_karta[ident]
        razlichno = False
        for kol in kolonki_sravneniya:
            star_znach = staroe.get(kol)
            nov_znach = novoe.get(kol)
            if str(star_znach) != str(nov_znach):
                razlichno = True
                break
        if razlichno:
            obnovleno += 1

    ostatok = bazovaya_bd[bazovaya_bd["ОП"] != programma] if not bazovaya_bd.empty else pd.DataFrame(columns=KOLONKI_BD)
    bd2 = pd.concat([ostatok, novyy_df], ignore_index=True)
    bd2 = primenit_tipy(bd2)

    return bd2, dobavleno, obnovleno, udaleno

def prochitat_zapusk() -> Dict[Tuple[str, str], bool]:
    zagruzheno = {}
    if not os.path.exists(FAYL_ZAPUSKA):
        return zagruzheno
    try:
        with open(FAYL_ZAPUSKA, "r", encoding="utf-8-sig", newline="") as f:
            chitatel = csv.DictReader(f)
            for stroka in chitatel:
                data = stroka.get("Дата")
                prog = stroka.get("ОП")
                if data in SPISOK_DAT and prog in PROGRAMMY:
                    zagruzheno[(prog, data)] = True
    except Exception:
        pass
    return zagruzheno


def zapisat_zapusk(zapisi: List[Dict[str, str]]):
    with open(FAYL_ZAPUSKA, "w", encoding="utf-8-sig", newline="") as f:
        pisatel = csv.DictWriter(f, fieldnames=["Дата", "ОП", "Файл", "Время"])
        pisatel.writeheader()
        [pisatel.writerow(zapis) for zapis in zapisi]


def dobavit_zapusk(programma: str, metka_daty: str, put_fayla: str):
    zapisi = []
    if os.path.exists(FAYL_ZAPUSKA):
        with open(FAYL_ZAPUSKA, "r", encoding="utf-8-sig", newline="") as f:
            chitatel = csv.DictReader(f)
            zapisi = list(chitatel)
    zapisi = [zapis for zapis in zapisi if not (zapis.get("Дата") == metka_daty and zapis.get("ОП") == programma)]
    zapisi.append({
        "Дата": metka_daty,
        "ОП": programma,
        "Файл": put_fayla,
        "Время": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    })
    zapisat_zapusk(zapisi)

class Prilozhenie(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(ZAGOLOVOK_PRILOZHENIYA)
        self.minsize(1200, 700)


        self.sgenerovannye_daty: set = set()
        self.vse_rezultaty: Dict[str, Dict[str, List[Dict]]] = {}
        self.vse_prokhodnye_bally: Dict[str, Dict[str, Optional[int]]] = {}
        self.poslednyaya_sortirovannaya_kolonka: Optional[str] = None

        self.imya_shrifta_pdf = "Helvetica"
        self.tekushchaya_metka_daty: str = SPISOK_DAT[0]
        self.karta_zagruzhennyh = prochitat_zapusk()
        self.aktivnyy_df_istochnik: Optional[pd.DataFrame] = None
        self.aktivnyy_df_vid: Optional[pd.DataFrame] = None
        self.aktivnyy_kontekst: str = ""
        self.filtr_soglasie = "Все"
        self.filtr_prioritet = "Любой"


        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.holst_fona = tk.Canvas(self, highlightthickness=0, bd=0)
        self.holst_fona.grid(row=0, column=0, sticky="nsew")
        self.konteyner = ttk.Frame(self.holst_fona)
        self.konteyner.place(relx=0, rely=0, relwidth=1, relheight=1)
        self._put_kartinki_fona = None
        self._obekt_kartinki_fona = None
        self._pil_kartinka_fona = None
        self.bind("<Configure>", self._pri_izmenenii_razmera)

        self.peremennaya_zagolovka = tk.StringVar(value="")
        shapka = ttk.Frame(self.konteyner)
        shapka.pack(side=tk.TOP, fill=tk.X, padx=8, pady=(8, 2))
        ttk.Label(shapka, textvariable=self.peremennaya_zagolovka, font=("Segoe UI", 12, "bold")).pack(side=tk.LEFT)

        self._postroit_tablitsu()
        self._postroit_statusbar()


        self._obespechit_startovye_konfigi()

        self._postroit_menyu()


        for katalog in [KATALOG_SPISKOV, KATALOG_OTCHETOV, KATALOG_SPISKOV_ZACHISLENIYA, KATALOG_ANALIZA, KATALOG_VHODNYKH_FAYLOV]:
            os.makedirs(katalog, exist_ok=True)


        if os.path.exists(KARTINKA_FONA_PO_UMOLCHANIYU):
            self.ustanovit_fon(KARTINKA_FONA_PO_UMOLCHANIYU)
        if proverit_reportlab():
            self.imya_shrifta_pdf = zaregistrirovat_kirillicheskiy_shrift()

        self.pokazat_bd_vse()

    def _obespechit_startovye_konfigi(self):
        global PROGRAMMY, MESTA, SPISOK_DAT
        if not os.path.exists(FAYL_KONFIG_PROGRAMM):
            self._okno_konfiga_programm(nachalnoe=True)
        if not os.path.exists(FAYL_KONFIG_DAT):
            self._okno_konfiga_dat(nachalnoe=True)
        PROGRAMMY, MESTA = zagruzit_konfig_programm()
        SPISOK_DAT[:] = zagruzit_konfig_dat()

    def _pri_izmenenii_razmera(self, sobytie):
        if self._put_kartinki_fona and PIL_DOSTUPEN and self._pil_kartinka_fona is not None:
            self._obnovit_kartinku_fona()

    def _postroit_tablitsu(self):
        ramka = ttk.Frame(self.konteyner)
        ramka.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=8, pady=6)
        self.derevo = ttk.Treeview(ramka, columns=KOLONKI_PROSMOTRA, show="headings")
        vertikalnaya_polosa = ttk.Scrollbar(ramka, orient="vertical", command=self.derevo.yview)
        gorizontalnaya_polosa = ttk.Scrollbar(ramka, orient="horizontal", command=self.derevo.xview)
        self.derevo.configure(yscroll=vertikalnaya_polosa.set, xscroll=gorizontalnaya_polosa.set)
        for kol in KOLONKI_PROSMOTRA:
            self.derevo.heading(kol, text=kol, command=lambda k=kol: self._sortirovat_po_kolonke(k, False))
            self.derevo.column(kol, width=140, anchor=tk.CENTER)
        self.derevo.grid(row=0, column=0, sticky="nsew")
        vertikalnaya_polosa.grid(row=0, column=1, sticky="ns")
        gorizontalnaya_polosa.grid(row=1, column=0, sticky="ew")
        ramka.grid_rowconfigure(0, weight=1)
        ramka.grid_columnconfigure(0, weight=1)

    def _postroit_statusbar(self):
        self.peremennaya_statusa = tk.StringVar(value="Готово")
        stroka_statusa = ttk.Frame(self.konteyner)
        stroka_statusa.pack(side=tk.BOTTOM, fill=tk.X)
        ttk.Label(stroka_statusa, textvariable=self.peremennaya_statusa, anchor="w").pack(side=tk.LEFT, padx=8, pady=4)

    def _postroit_menyu(self):
        panel_menyu = tk.Menu(self)


        self.menyu_fayl = tk.Menu(panel_menyu, tearoff=0)
        self.menyu_fayl.add_command(label="Генератор списков…", command=self.menyu_generator)
        self.menyu_fayl.add_separator()
        for data in SPISOK_DAT:
            for prog in PROGRAMMY:
                nadpis = f"Загрузить список {prog} на {data}"
                self.menyu_fayl.add_command(
                    label=nadpis,
                    command=lambda p=prog, d=data: self.menyu_zagruzit_excel_dlya(p, d),
                    state=(tk.DISABLED if self.karta_zagruzhennyh.get((prog, data)) else tk.NORMAL),
                )
        self.menyu_fayl.add_separator()
        self.menyu_fayl.add_command(label="Выход", command=self.destroy)
        panel_menyu.add_cascade(label="Файл", menu=self.menyu_fayl)


        menyu_prosmotr = tk.Menu(panel_menyu, tearoff=0)
        menyu_prosmotr.add_command(label="Актуальный список", command=self.pokazat_bd_vse)
        menyu_prosmotr.add_separator()
        for data in SPISOK_DAT:
            for prog in PROGRAMMY:
                menyu_prosmotr.add_command(
                    label=f"Список {prog} на {data}",
                    command=lambda d=data, p=prog: self.posmotret_fayl_spiska(d, p),
                )
        menyu_prosmotr.add_separator()
        menyu_prosmotr.add_command(label="Открыть список (выбор ОП и даты)", command=self.menyu_prosmotr_vybor_spiska)
        panel_menyu.add_cascade(label="Просмотр", menu=menyu_prosmotr)


        self.menyu_spiski = tk.Menu(panel_menyu, tearoff=0)
        self.menyu_spiski.add_command(label="Сформировать списки…", command=self.menyu_rasschitat_spiski)
        self.menyu_spiski.add_separator()
        for prog in PROGRAMMY:
            self.menyu_spiski.add_command(label=f"Список {prog}…", command=lambda p=prog: self.menyu_pokazat_zapros_rezultata(p))
        panel_menyu.add_cascade(label="Списки", menu=self.menyu_spiski)


        menyu_filtr = tk.Menu(panel_menyu, tearoff=0)
        podmenyu_soglasie = tk.Menu(menyu_filtr, tearoff=0)
        podmenyu_soglasie.add_command(label="Все", command=lambda: self.ustanovit_filtr_soglasie("Все"))
        podmenyu_soglasie.add_command(label="Только с согласием", command=lambda: self.ustanovit_filtr_soglasie("Только с согласием"))
        podmenyu_soglasie.add_command(label="Без согласия", command=lambda: self.ustanovit_filtr_soglasie("Без согласия"))
        menyu_filtr.add_cascade(label="Фильтр по согласию", menu=podmenyu_soglasie)
        podmenyu_prioritet = tk.Menu(menyu_filtr, tearoff=0)
        podmenyu_prioritet.add_command(label="Любой", command=lambda: self.ustanovit_filtr_prioritet("Любой"))
        for nomer in [1, 2, 3, 4]:
            podmenyu_prioritet.add_command(label=f"Только приоритет = {nomer}", command=lambda v=nomer: self.ustanovit_filtr_prioritet(str(v)))
        menyu_filtr.add_cascade(label="Фильтр по приоритету", menu=podmenyu_prioritet)
        menyu_filtr.add_separator()
        menyu_filtr.add_command(label="Поиск по ID (из БД)", command=self.menyu_poisk_id_bd)
        menyu_filtr.add_command(label="Поиск по ID (везде)", command=self.menyu_poisk_id_vse)
        panel_menyu.add_cascade(label="Фильтр", menu=menyu_filtr)


        menyu_sortirovka = tk.Menu(panel_menyu, tearoff=0)
        podmenyu_vozrast = tk.Menu(menyu_sortirovka, tearoff=0)
        podmenyu_ubyvanie = tk.Menu(menyu_sortirovka, tearoff=0)
        kolonki_sortirovki = [
            "ID",
            "Согласие",
            "Приоритет",
            "Балл Физика/ИКТ",
            "Балл Русский язык",
            "Балл Математика",
            "Сумма баллов",
        ]
        for kol in kolonki_sortirovki:
            podmenyu_vozrast.add_command(label=kol, command=lambda k=kol: self.menyu_sortirovat_kolonku(k, True))
            podmenyu_ubyvanie.add_command(label=kol, command=lambda k=kol: self.menyu_sortirovat_kolonku(k, False))
        menyu_sortirovka.add_cascade(label="По возрастанию", menu=podmenyu_vozrast)
        menyu_sortirovka.add_cascade(label="По убыванию", menu=podmenyu_ubyvanie)
        panel_menyu.add_cascade(label="Сортировка", menu=menyu_sortirovka)


        menyu_otchety = tk.Menu(panel_menyu, tearoff=0)
        menyu_otchety.add_command(label="Статистика…", command=self.menyu_statistika_dlya_daty)
        menyu_otchety.add_command(label="К зачислению (в PDF)…", command=self.menyu_otchety_pdf)
        menyu_otchety.add_command(label="График", command=self.menyu_okno_grafika_istorii)
        menyu_otchety.add_command(label="Анализ генерации списков", command=self.menyu_pokazat_otchet2)
        menyu_otchety.add_command(label="Число уникальных ID", command=self.menyu_pokazat_otchet1)
        menyu_otchety.add_command(label="Динамика", command=self.menyu_pokazat_tablitsu_istorii)
        panel_menyu.add_cascade(label="Отчёты", menu=menyu_otchety)


        menyu_nastroyki = tk.Menu(panel_menyu, tearoff=0)
        menyu_nastroyki.add_command(label="ОП и Количество мест…", command=self.menyu_konfig_programm)
        menyu_nastroyki.add_command(label="Даты приёма…", command=self.menyu_konfig_dat)
        panel_menyu.add_cascade(label="Настройки", menu=menyu_nastroyki)

        self.config(menu=panel_menyu)

    def _vklyuchit_vse_punkty_zagruzki_menyu_fayl(self):
        try:
            konets = self.menyu_fayl.index("end") or -1
            for indeks in range(konets + 1):
                if self.menyu_fayl.type(indeks) == "command":
                    nadpis = self.menyu_fayl.entrycget(indeks, "label")
                    if isinstance(nadpis, str) and nadpis.startswith("Загрузить список "):
                        self.menyu_fayl.entryconfig(indeks, state=tk.NORMAL)
        except Exception:
            pass

    def menyu_generator(self):
        if not messagebox.askyesno("Генератор списков", "Все временные файлы и данные будут удалены. Продолжить?"):
            return
        for katalog in [KATALOG_SPISKOV, KATALOG_OTCHETOV, KATALOG_ANALIZA, KATALOG_SPISKOV_ZACHISLENIYA]:
            if os.path.isdir(katalog):
                for imya_fayla in os.listdir(katalog):
                    try:
                        os.remove(os.path.join(katalog, imya_fayla))
                    except Exception:
                        pass
        for fayl in [FAYL_BD, FAYL_ZAPUSKA, FAYL_ISTORII]:
            if os.path.exists(fayl):
                try:
                    with open(fayl, "r", encoding="utf-8-sig", newline="") as vhod:
                        chitatel = csv.reader(vhod)
                        zagolovok = next(chitatel, None)
                    if zagolovok:
                        with open(fayl, "w", encoding="utf-8-sig", newline="") as vyhod:
                            pisatel = csv.writer(vyhod)
                            pisatel.writerow(zagolovok)
                except Exception:
                    pass
        for skript in ["spisok.py", "proverka.py"]:
            if os.path.exists(skript):
                try:
                    subprocess.run([sys.executable, skript], check=False)
                except Exception as oshibka:
                    messagebox.showwarning("Генератор", f"Не удалось запустить {skript}: {oshibka}")
            else:
                messagebox.showwarning("Генератор", f"Файл {skript} не найден")
        self.karta_zagruzhennyh.clear()
        self._perestroit_menyu()
        self._vklyuchit_vse_punkty_zagruzki_menyu_fayl()
        self.sgenerovannye_daty.clear()
        self.vse_rezultaty.clear()
        self.vse_prokhodnye_bally.clear()
        self.pokazat_bd_vse()
        self._ustanovit_status("Генерация списков завершена")

    def ustanovit_fon(self, put: Optional[str]):
        self._put_kartinki_fona = put
        if put is None:
            self._pil_kartinka_fona = None
            self._obekt_kartinki_fona = None
            self.holst_fona.delete("FON")
            return
        try:
            if PIL_DOSTUPEN:
                self._pil_kartinka_fona = Image.open(put)
                self._obnovit_kartinku_fona()
            else:
                kartinka = tk.PhotoImage(file=put)
                self._obekt_kartinki_fona = kartinka
                self._otrisovat_fon(kartinka)
        except Exception as oshibka:
            messagebox.showwarning("Фон", f"Не удалось установить фон: {oshibka}")

    def _otrisovat_fon(self, tk_kartinka):
        self.holst_fona.delete("FON")
        shirina = self.holst_fona.winfo_width()
        vysota = self.holst_fona.winfo_height()
        shirina_kart = tk_kartinka.width()
        vysota_kart = tk_kartinka.height()
        x = max(0, (shirina - shirina_kart)//2)
        y = max(0, (vysota - vysota_kart)//2)
        self.holst_fona.create_image(x, y, image=tk_kartinka, anchor="nw", tags="FON")

    def _obnovit_kartinku_fona(self):
        if not self._put_kartinki_fona or not PIL_DOSTUPEN or self._pil_kartinka_fona is None:
            return
        shirina = max(1, self.holst_fona.winfo_width())
        vysota = max(1, self.holst_fona.winfo_height())
        try:
            kartinka = self._pil_kartinka_fona.copy().resize((shirina, vysota))
            self._obekt_kartinki_fona = ImageTk.PhotoImage(kartinka)
            self._otrisovat_fon(self._obekt_kartinki_fona)
        except Exception:
            pass

    def _pokazat_dataframe(self, df: "pd.DataFrame"):
        for kol in KOLONKI_PROSMOTRA:
            if kol not in df.columns:
                df[kol] = pd.NA
        df = df[KOLONKI_PROSMOTRA]
        self.derevo.delete(*self.derevo.get_children())
        for _, strok in df.iterrows():
            znacheniya = []
            for klyuch in KOLONKI_PROSMOTRA:
                znach = strok.get(klyuch)
                if isinstance(znach, bool):
                    znacheniya.append("TRUE" if znach else "FALSE")
                elif znach is None or (isinstance(znach, float) and pd.isna(znach)) or (str(znach) == "<NA>"):
                    znacheniya.append("")
                else:
                    znacheniya.append(str(znach))
            self.derevo.insert("", tk.END, values=znacheniya)

    def _sortirovat_po_kolonke(self, kolonka: str, po_ubyvaniyu: bool):
        self.sortirovat_derevo_po_kolonke(kolonka, po_vozrastaniyu=not po_ubyvaniyu)

    def sortirovat_derevo_po_kolonke(self, kolonka: str, po_vozrastaniyu: bool):
        self.poslednyaya_sortirovannaya_kolonka = kolonka
        dannye = []
        for identifikator in self.derevo.get_children(""):
            znacheniya = self.derevo.item(identifikator, "values")
            stroka = {k: znacheniya[i] for i, k in enumerate(KOLONKI_PROSMOTRA)}
            dannye.append(stroka)

        chislovye_kolonki = {
            "ID",
            "Приоритет",
            "Балл Физика/ИКТ",
            "Балл Русский язык",
            "Балл Математика",
            "Балл за ИД",
            "Сумма баллов",
        }

        def klyuch_sortirovki(strok):
            znach = strok.get(kolonka)
            if znach in ("", None):
                return (1, None)
            if kolonka == "Согласие":
                znachenie = 1 if str(znach).upper() == "TRUE" else 0
                return (0, znachenie)
            if kolonka in chislovye_kolonki:
                try:
                    return (0, float(znach))
                except Exception:
                    return (0, float("inf"))
            return (0, str(znach))

        dannye.sort(key=klyuch_sortirovki, reverse=not po_vozrastaniyu)
        self.derevo.delete(*self.derevo.get_children())
        for strok in dannye:
            self.derevo.insert("", tk.END, values=[strok.get(k, "") for k in KOLONKI_PROSMOTRA])

    def menyu_sortirovat_kolonku(self, kolonka: str, po_vozrastaniyu: bool):
        self.sortirovat_derevo_po_kolonke(kolonka, po_vozrastaniyu)

    def _ustanovit_status(self, tekst: str):
        self.peremennaya_statusa.set(tekst)

    def zaprosit_vybor_daty(self, zagolovok="Дата", priglashenie="Выберите дату:") -> Optional[str]:
        okno = tk.Toplevel(self)
        okno.title(zagolovok)
        okno.transient(self)
        okno.grab_set()
        peremennaya_vybora = tk.StringVar(value=self.tekushchaya_metka_daty if self.tekushchaya_metka_daty in SPISOK_DAT else SPISOK_DAT[0])
        ramka = ttk.Frame(okno, padding=10)
        ramka.pack(fill=tk.BOTH, expand=True)
        ttk.Label(ramka, text=priglashenie).pack(anchor="w", pady=(0,6))
        vypadayushchiy_spisok = ttk.Combobox(ramka, values=SPISOK_DAT, textvariable=peremennaya_vybora, state="readonly")
        vypadayushchiy_spisok.pack(fill=tk.X)
        vypadayushchiy_spisok.focus_set()
        knopki = ttk.Frame(ramka)
        knopki.pack(fill=tk.X, pady=(10,0))
        def ok():
            okno.vybrano = peremennaya_vybora.get()
            okno.destroy()
        def otmena():
            okno.vybrano = None
            okno.destroy()
        ttk.Button(knopki, text="OK", command=ok).pack(side=tk.RIGHT, padx=(6,0))
        ttk.Button(knopki, text="Отмена", command=otmena).pack(side=tk.RIGHT)
        okno.bind("<Return>", lambda e: ok())
        okno.bind("<Escape>", lambda e: otmena())
        okno.update_idletasks()
        x = self.winfo_rootx() + (self.winfo_width()-okno.winfo_width())//2
        y = self.winfo_rooty() + (self.winfo_height()-okno.winfo_height())//2
        okno.geometry(f"+{x}+{y}")
        self.wait_window(okno)
        return getattr(okno, "vybrano", None)

    def zaprosit_vybor_programmy_i_daty(self, zagolovok="Просмотр", priglashenie="Выберите ОП и дату:") -> Optional[Tuple[str,str]]:
        okno = tk.Toplevel(self)
        okno.title(zagolovok)
        okno.transient(self)
        okno.grab_set()
        peremennaya_prog = tk.StringVar(value=PROGRAMMY[0])
        peremennaya_daty = tk.StringVar(value=self.tekushchaya_metka_daty if self.tekushchaya_metka_daty in SPISOK_DAT else SPISOK_DAT[0])
        ramka = ttk.Frame(okno, padding=10)
        ramka.pack(fill=tk.BOTH, expand=True)
        ttk.Label(ramka, text=priglashenie).pack(anchor="w", pady=(0,6))
        stroka1 = ttk.Frame(ramka)
        stroka1.pack(fill=tk.X, pady=4)
        ttk.Label(stroka1, text="ОП:").pack(side=tk.LEFT)
        ttk.Combobox(stroka1, values=PROGRAMMY, textvariable=peremennaya_prog, state="readonly").pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(6,0))
        stroka2 = ttk.Frame(ramka)
        stroka2.pack(fill=tk.X, pady=4)
        ttk.Label(stroka2, text="Дата:").pack(side=tk.LEFT)
        ttk.Combobox(stroka2, values=SPISOK_DAT, textvariable=peremennaya_daty, state="readonly").pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(6,0))
        knopki = ttk.Frame(ramka)
        knopki.pack(fill=tk.X, pady=(10,0))
        def ok():
            okno.vybrano=(peremennaya_prog.get(), peremennaya_daty.get())
            okno.destroy()
        def otmena():
            okno.vybrano=None
            okno.destroy()
        ttk.Button(knopki, text="OK", command=ok).pack(side=tk.RIGHT, padx=(6,0))
        ttk.Button(knopki, text="Отмена", command=otmena).pack(side=tk.RIGHT)
        okno.bind("<Return>", lambda e: ok())
        okno.bind("<Escape>", lambda e: otmena())
        okno.update_idletasks()
        x=self.winfo_rootx()+(self.winfo_width()-okno.winfo_width())//2
        y=self.winfo_rooty()+(self.winfo_height()-okno.winfo_height())//2
        okno.geometry(f"+{x}+{y}")
        self.wait_window(okno)
        return getattr(okno,"vybrano",None)

    def pokazat_bd_vse(self):
        if not proverit_pandas():
            return
        df = prochitat_bd()
        self.aktivnyy_df_istochnik = df.copy()
        self.primenit_filtry()
        self.aktivnyy_kontekst = "Все записи БД"
        self.peremennaya_zagolovka.set(self.aktivnyy_kontekst)

    def menyu_prosmotr_vybor_spiska(self):
        vybor = self.zaprosit_vybor_programmy_i_daty(zagolovok="Просмотр списка", priglashenie="Выберите ОП и дату:")
        if not vybor:
            return
        prog, metka_daty = vybor
        self.posmotret_fayl_spiska(metka_daty, prog)

    def posmotret_fayl_spiska(self, metka_daty: str, programma: str):
        if not proverit_pandas():
            return
        put = os.path.join(KATALOG_SPISKOV, f"{metka_daty}_{programma}.csv")
        if not os.path.exists(put):
            messagebox.showinfo("Просмотр", f"Файл не найден: {put}")
            return
        try:
            df = pd.read_csv(put, encoding="utf-8-sig")
        except Exception as oshibka:
            messagebox.showerror("Просмотр", f"Ошибка чтения: {oshibka}")
            return
        for kol in [k for k in KOLONKI_BD if k!="ОП"]:
            if kol not in df.columns:
                df[kol]=pd.NA
        df = df[[k for k in KOLONKI_BD if k!="ОП"]]
        self.aktivnyy_df_istochnik = df
        self.primenit_filtry()
        self.aktivnyy_kontekst = f"Список поступающих на {programma} от {metka_daty}"
        self.peremennaya_zagolovka.set(self.aktivnyy_kontekst)

    def ustanovit_filtr_soglasie(self, rezhim:str):
        self.filtr_soglasie=rezhim
        self.primenit_filtry()

    def ustanovit_filtr_prioritet(self, rezhim:str):
        self.filtr_prioritet=rezhim
        self.primenit_filtry()

    def primenit_filtry(self):
        if not proverit_pandas():
            return
        if self.aktivnyy_df_istochnik is None:
            df = pd.DataFrame(columns=KOLONKI_PROSMOTRA)
        else:
            df = self.aktivnyy_df_istochnik.copy()
        if "Согласие" in df.columns:
            if self.filtr_soglasie == "Только с согласием":
                df = df[df_bool(df["Согласие"])].copy()
            elif self.filtr_soglasie == "Без согласия":
                try:
                    df = df[df["Согласие"] == False].copy()
                except Exception:
                    df = df[df["Согласие"].astype(str).str.lower().isin(["false", "0", "нет"])].copy()
        if self.filtr_prioritet not in ("Любой", None) and "Приоритет" in df.columns:
            try:
                prioritet = int(str(self.filtr_prioritet).strip())
                df = df[df["Приоритет"] == prioritet].copy()
            except Exception:
                pass
        self.aktivnyy_df_vid = df
        self._pokazat_dataframe(df)
        self._ustanovit_status(f"Строк: {len(df)}")

    def menyu_poisk_id_bd(self):
        if not proverit_pandas():
            return
        zapros = simpledialog.askstring("Поиск", "Введите ID:")
        if not zapros:
            return
        try:
            identifikator_zaprosa=int(zapros.strip())
        except Exception:
            messagebox.showinfo("Поиск","ID должен быть числом")
            return
        bd = prochitat_bd()
        df = bd[bd["ID"]==identifikator_zaprosa].copy()
        if df.empty:
            messagebox.showinfo("Поиск", f"ID {identifikator_zaprosa} не найден в БД")
            return
        self.aktivnyy_df_istochnik = df
        self.primenit_filtry()
        self.aktivnyy_kontekst = f"Данные БД по ID {identifikator_zaprosa}"
        self.peremennaya_zagolovka.set(self.aktivnyy_kontekst)

    def _prochitat_stroki_csv(self, put: str) -> List[List[str]]:
        with open(put, "r", encoding="utf-8-sig", newline="") as f:
            obrazets = f.read(4096)
            f.seek(0)
            try:
                dialekt = csv.Sniffer().sniff(obrazets, delimiters=",;")
                chitatel = csv.reader(f, dialekt)
            except Exception:
                razdelitel = ";" if obrazets.count(";") > obrazets.count(",") else ","
                chitatel = csv.reader(f, delimiter=razdelitel)
            return list(chitatel)

    def menyu_poisk_id_vse(self):
        zapros = simpledialog.askstring("Поиск ID", "Введите ID:")
        if not zapros:
            return
        try:
            identifikator_zaprosa = int(zapros.strip())
        except Exception:
            messagebox.showinfo("Поиск", "ID должен быть числом")
            return
        rezultaty = []
        if os.path.exists(FAYL_BD):
            stroki = self._prochitat_stroki_csv(FAYL_BD)
            zagolovok = stroki[0] if stroki else []
            for stroka in stroki[1:]:
                if stroka and str(stroka[0]).strip() == str(identifikator_zaprosa):
                    rezultaty.append(("Актуальный список", "БД", zagolovok, stroka))
                    break
        for papka, metka in [(KATALOG_SPISKOV, "Список поступающих"), (KATALOG_SPISKOV_ZACHISLENIYA, "Рекомендованные к зачислению")]:
            if not os.path.isdir(papka):
                continue
            for imya_fayla in os.listdir(papka):
                if not imya_fayla.lower().endswith(".csv"):
                    continue
                put = os.path.join(papka, imya_fayla)
                stroki = self._prochitat_stroki_csv(put)
                if not stroki:
                    continue
                zagolovok = stroki[0]
                for stroka in stroki[1:]:
                    if stroka and str(stroka[0]).strip() == str(identifikator_zaprosa):
                        rezultaty.append((metka, os.path.splitext(imya_fayla)[0], zagolovok, stroka))
                        break
        if not rezultaty:
            messagebox.showinfo("Поиск", "ID не найден")
            return
        okno = tk.Toplevel(self)
        okno.title(f"Результаты ID {identifikator_zaprosa}")
        okno.minsize(700, 450)
        bloknot = ttk.Notebook(okno)
        bloknot.pack(fill=tk.BOTH, expand=True)
        for metka, imya_fayla, zagolovok, stroka in rezultaty:
            vkladka = ttk.Frame(bloknot)
            bloknot.add(vkladka, text=f"{metka}: {imya_fayla}")
            derevo = ttk.Treeview(vkladka, columns=zagolovok, show="headings")
            vertikalnaya_polosa = ttk.Scrollbar(vkladka, orient="vertical", command=derevo.yview)
            gorizontalnaya_polosa = ttk.Scrollbar(vkladka, orient="horizontal", command=derevo.xview)
            derevo.configure(yscroll=vertikalnaya_polosa.set, xscroll=gorizontalnaya_polosa.set)
            for kol in zagolovok:
                derevo.heading(kol, text=kol)
                derevo.column(kol, width=140, anchor=tk.CENTER)
            derevo.grid(row=0, column=0, sticky="nsew")
            vertikalnaya_polosa.grid(row=0, column=1, sticky="ns")
            gorizontalnaya_polosa.grid(row=1, column=0, sticky="ew")
            vkladka.grid_rowconfigure(0, weight=1)
            vkladka.grid_columnconfigure(0, weight=1)
            znacheniya = stroka + [""] * (len(zagolovok) - len(stroka))
            znacheniya = znacheniya[:len(zagolovok)]
            derevo.insert("", tk.END, values=znacheniya)
        ramka_knopok = ttk.Frame(okno)
        ramka_knopok.pack(fill=tk.X)
        ttk.Button(ramka_knopok, text="Закрыть", command=okno.destroy).pack(side=tk.RIGHT, padx=8, pady=8)

    def menyu_rasschitat_spiski(self):
        vybor=self.zaprosit_vybor_daty(zagolovok="Дата",priglashenie="Выберите дату:")
        if vybor is None:
            return
        if vybor in self.sgenerovannye_daty:
            messagebox.showinfo("Списки",f"Список на {vybor} уже сформирован")
            return
        self.rasschitat_zachislenie_dlya_daty(vybor)
        self._sokhranit_spiski_dlya_daty(vybor)
        self.pokazat_rezultat_programmy(PROGRAMMY[0], vybor)

    def rasschitat_zachislenie_dlya_daty(self, metka_daty: str):
        if not proverit_pandas():
            return
        bd=prochitat_bd()
        if bd.empty:
            messagebox.showinfo("Списки","БД пуста")
            return
        baza=bd[(df_bool(bd["Согласие"])) & (pd.notna(bd["Сумма баллов"]))].copy()
        if baza.empty:
            messagebox.showinfo("Списки","Нет согласий")
            return
        ostavshiesya_mesta={prog:MESTA[prog] for prog in PROGRAMMY}
        naznachennye=set()
        rezultaty={prog:[] for prog in PROGRAMMY}
        for uroven in range(1,5):
            for prog in PROGRAMMY:
                mesta=ostavshiesya_mesta[prog]
                if mesta<=0:
                    continue
                kandidaty=baza[(baza["ОП"]==prog)&(baza["Приоритет"]==uroven)&(~baza["ID"].isin(list(naznachennye)))]
                if kandidaty.empty:
                    continue
                kolonki_sortirovki=["Сумма баллов","Балл Математика","Балл Физика/ИКТ","Балл Русский язык","Балл за ИД"]
                kandidaty=kandidaty.sort_values(by=kolonki_sortirovki,ascending=[False]*len(kolonki_sortirovki))
                vzyat=min(mesta,len(kandidaty))
                vybrannye=kandidaty.head(vzyat)
                for _,strok in vybrannye.iterrows():
                    rezultaty[prog].append({"ID":int(strok.ID),"Сумма баллов":int(strok["Сумма баллов"]),"Приоритет":int(strok.Приоритет)})
                    naznachennye.add(int(strok.ID))
                    ostavshiesya_mesta[prog]-=1
        prokhodnye_bally={}
        for prog in PROGRAMMY:
            obshchee_mest=MESTA[prog]
            spisok=rezultaty[prog]
            if len(spisok)<obshchee_mest:
                prokhodnye_bally[prog]=None
            else:
                spisok_sortirovannyy=sorted(spisok,key=lambda x:(-x["Сумма баллов"],x["Приоритет"]))
                prokhodnye_bally[prog]=spisok_sortirovannyy[obshchee_mest-1]["Сумма баллов"]
        self.vse_rezultaty[metka_daty]=rezultaty
        self.vse_prokhodnye_bally[metka_daty]=prokhodnye_bally
        self.sgenerovannye_daty.add(metka_daty)
        self._dobavit_istoriyu(metka_daty, prokhodnye_bally)
        self.tekushchaya_metka_daty = metka_daty
        self._ustanovit_status(f"Списки сформированы на {metka_daty}")

    def _sokhranit_spiski_dlya_daty(self, metka_daty: str):
        if not proverit_pandas():
            return
        os.makedirs(KATALOG_SPISKOV_ZACHISLENIYA, exist_ok=True)
        bd = prochitat_bd()
        for prog in PROGRAMMY:
            identifikatory = [int(element["ID"]) for element in self.vse_rezultaty.get(metka_daty, {}).get(prog, [])]
            if not identifikatory:
                put_vyhoda = os.path.join(KATALOG_SPISKOV_ZACHISLENIYA, f"Список {prog} на {metka_daty}.csv")
                pd.DataFrame(columns=[k for k in KOLONKI_BD if k != "ОП"]).to_csv(put_vyhoda, index=False, encoding="utf-8-sig")
                continue
            df = bd[(bd["ОП"] == prog) & (bd["ID"].isin(identifikatory))].copy()
            if not df.empty:
                kolonki_sortirovki = ["Сумма баллов", "Балл Математика", "Балл Физика/ИКТ", "Балл Русский язык", "Балл за ИД"]
                df = df.sort_values(by=kolonki_sortirovki, ascending=[False]*len(kolonki_sortirovki))
            put_vyhoda = os.path.join(KATALOG_SPISKOV_ZACHISLENIYA, f"Список {prog} на {metka_daty}.csv")
            df_vyhod = df[[k for k in KOLONKI_BD if k != "ОП"]]
            df_vyhod.to_csv(put_vyhoda, index=False, encoding="utf-8-sig")

    def _dobavit_istoriyu(self, metka_daty:str, prokhodnye_bally:Dict[str,Optional[int]]):
        stroki=[]
        if os.path.exists(FAYL_ISTORII):
            with open(FAYL_ISTORII,"r",encoding="utf-8-sig",newline="") as f:
                stroki=list(csv.DictReader(f))
        stroki=[strok for strok in stroki if strok.get("Дата")!=metka_daty]
        stroki.append({"Дата":metka_daty, **{prog:("НЕДОБОР" if prokhodnye_bally.get(prog) is None else str(prokhodnye_bally.get(prog))) for prog in PROGRAMMY}})
        with open(FAYL_ISTORII,"w",encoding="utf-8-sig",newline="") as f:
            pisatel=csv.DictWriter(f,fieldnames=["Дата"]+PROGRAMMY)
            pisatel.writeheader()
            stroki_sortirovannye=sorted(stroki,key=lambda strok: SPISOK_DAT.index(strok["Дата"]) if strok["Дата"] in SPISOK_DAT else 999)
            for strok in stroki_sortirovannye:
                pisatel.writerow(strok)

    def menyu_pokazat_zapros_rezultata(self, programma: str):
        if not self.sgenerovannye_daty:
            messagebox.showinfo("Списки","Списки ещё не сформированы")
            return
        daty_dostupnye=[data for data in SPISOK_DAT if data in self.sgenerovannye_daty]
        if not daty_dostupnye:
            messagebox.showinfo("Списки","Списки не сформированы")
            return
        vybor=self.zaprosit_vybor_daty()
        if vybor is None:
            return
        if vybor not in self.sgenerovannye_daty:
            messagebox.showinfo("Списки","Список не сформирован")
            return
        self.pokazat_rezultat_programmy(programma, vybor)

    def pokazat_rezultat_programmy(self, programma:str, metka_daty:str):
        if metka_daty not in self.sgenerovannye_daty:
            messagebox.showinfo("Списки","Список не сформирован")
            return
        identifikatory=[int(element["ID"]) for element in self.vse_rezultaty.get(metka_daty,{}).get(programma,[])]
        bd=prochitat_bd()
        df=bd[(bd["ОП"]==programma)&(bd["ID"].isin(identifikatory))].copy()
        if not df.empty:
            kolonki_sortirovki=["Сумма баллов","Балл Математика","Балл Физика/ИКТ","Балл Русский язык","Балл за ИД"]
            df=df.sort_values(by=kolonki_sortirovki,ascending=[False]*len(kolonki_sortirovki))
        self.aktivnyy_df_istochnik=df
        self.primenit_filtry()
        self.aktivnyy_kontekst=f"Зачисленные на {programma} — {metka_daty}"
        self.peremennaya_zagolovka.set(self.aktivnyy_kontekst)
        self.tekushchaya_metka_daty = metka_daty

    def menyu_zagruzit_excel_dlya(self, programma: str, metka_daty: str):
        if not proverit_pandas():
            return
        put=filedialog.askopenfilename(
            title=f"Выберите XLSX для {programma} на {metka_daty}",
            initialdir=KATALOG_VHODNYKH_FAYLOV,
            filetypes=[("Excel (XLSX)", "*.xlsx"), ("Все файлы", "*.*")]
        )
        if not put:
            return
        rasshirenie=os.path.splitext(put)[1].lower()
        if rasshirenie != ".xlsx":
            messagebox.showinfo("Загрузка","Ожидается файл формата .xlsx")
            return
        dvizhok="openpyxl"
        try:
            excel_fayl=pd.ExcelFile(put,engine=dvizhok)
            frejmy=[excel_fayl.parse(list) for list in excel_fayl.sheet_names]
            frejmy=[frejm for frejm in frejmy if not frejm.empty]
            if not frejmy:
                messagebox.showinfo("Загрузка","Файл пустой")
                return
            df=pd.concat(frejmy,ignore_index=True)
        except Exception as oshibka:
            messagebox.showerror("Excel",f"Ошибка: {oshibka}")
            return
        df=normalizovat_zagolovki(df)
        if df.empty:
            messagebox.showinfo("Загрузка","После нормализации нет валидных строк")
            return
        bd=prochitat_bd()
        bd2,dobavleno,obnovleno,udaleno=obnovit_programmu(bd,df,programma)
        zapisat_bd(bd2)
        os.makedirs(KATALOG_SPISKOV,exist_ok=True)
        put_vyhoda=os.path.join(KATALOG_SPISKOV,f"{metka_daty}_{programma}.csv")
        df_vyhod=df[[k for k in KOLONKI_BD if k!="ОП"]]
        df_vyhod.to_csv(put_vyhoda,index=False,encoding="utf-8-sig")
        dobavit_zapusk(programma,metka_daty,put_vyhoda)
        self.karta_zagruzhennyh[(programma,metka_daty)] = True
        self._otklyuchit_punkt_menyu_fayl(programma,metka_daty)
        self._ustanovit_status(f"Загружено {programma} {metka_daty}: +{dobavleno}/~{obnovleno}/-{udaleno}")
        self.posmotret_fayl_spiska(metka_daty,programma)

    def _otklyuchit_punkt_menyu_fayl(self, programma:str, metka_daty:str):
        try:
            konets=self.menyu_fayl.index("end") or -1
            tselevaya_nadpis=f"Загрузить список {programma} на {metka_daty}"
            for indeks in range(konets+1):
                if self.menyu_fayl.type(indeks)=="command" and self.menyu_fayl.entrycget(indeks,"label")==tselevaya_nadpis:
                    self.menyu_fayl.entryconfig(indeks,state=tk.DISABLED)
                    break
        except Exception:
            pass

    def _rasschitat_statistiku(self, bazovyy_df: "pd.DataFrame") -> Dict[str, Dict[str, int]]:
        statistika={}
        for prog in PROGRAMMY:
            stat={
                "Obshchee_kol_zayavleniy":int(bazovyy_df[bazovyy_df["ОП"]==prog].shape[0]),
                "Mesta":MESTA[prog],
                "Zayavleniya_prior_1":int(((bazovyy_df["ОП"]==prog)&(bazovyy_df["Приоритет"]==1)).sum()),
                "Zayavleniya_prior_2":int(((bazovyy_df["ОП"]==prog)&(bazovyy_df["Приоритет"]==2)).sum()),
                "Zayavleniya_prior_3":int(((bazovyy_df["ОП"]==prog)&(bazovyy_df["Приоритет"]==3)).sum()),
                "Zayavleniya_prior_4":int(((bazovyy_df["ОП"]==prog)&(bazovyy_df["Приоритет"]==4)).sum()),
                "Zachisleno_prior_1":0,
                "Z2":0,"Z3":0,"Z4":0,
                "Zachisleno_prior_2":0,"Zachisleno_prior_3":0,"Zachisleno_prior_4":0,
                "Minimalnyy_ball":"НЕДОБОР"}
            if self.vse_rezultaty.get(self.tekushchaya_metka_daty):
                for element in self.vse_rezultaty[self.tekushchaya_metka_daty].get(prog,[]):
                    prior=element["Приоритет"]
                    stat[f"Zachisleno_prior_{prior}"]+=1
                prokhodnoy_ball=self.vse_prokhodnye_bally[self.tekushchaya_metka_daty].get(prog)
                stat["Minimalnyy_ball"] = prokhodnoy_ball if prokhodnoy_ball is not None else "НЕДОБОР"
            statistika[prog]=stat
        return statistika

    def menyu_statistika_dlya_daty(self):
        vybor=self.zaprosit_vybor_daty(zagolovok="Дата статистики",priglashenie="Выберите дату:")
        if vybor is None:
            return
        if vybor not in self.sgenerovannye_daty:
            messagebox.showinfo("Статистика","На текущую дату нет информации")
            return
        self.tekushchaya_metka_daty = vybor
        self._pokazat_statistiku_dlya(vybor)

    def _pokazat_statistiku_dlya(self, metka_daty:str):
        if not proverit_pandas():
            return
        baza=prochitat_bd()
        statistika=self._rasschitat_statistiku(baza)
        okno=tk.Toplevel(self)
        okno.title("Статистика по ОП")
        okno.minsize(650,450)
        konteyner=ttk.Frame(okno)
        konteyner.pack(fill=tk.BOTH,expand=True)
        ttk.Label(konteyner,text=f"Минимальные проходные баллы на дату {metka_daty}",font=("Segoe UI",12,"bold")).pack(pady=6)
        ramka1=ttk.Frame(konteyner)
        ramka1.pack(fill=tk.X,padx=8)
        for prog in PROGRAMMY:
            ttk.Label(ramka1,text=f"{prog}: {statistika[prog]['Minimalnyy_ball']}").pack(side=tk.LEFT,padx=12)
        ttk.Label(konteyner,text="Сводная таблица",font=("Segoe UI",11,"bold")).pack(pady=(10,4))
        kolonki=["Показатель"]+PROGRAMMY
        derevo=ttk.Treeview(konteyner,columns=kolonki,show="headings")
        for kol in kolonki:
            derevo.heading(kol,text=kol)
            derevo.column(kol,width=130,anchor=tk.CENTER)
        derevo.pack(fill=tk.BOTH,expand=True,padx=8,pady=6)
        def dobavit(nazvanie,poluchit):
            znacheniya=[nazvanie]+[str(poluchit(statistika[prog])) for prog in PROGRAMMY]
            derevo.insert("",tk.END,values=znacheniya)
        dobavit("Общее кол-во заявлений",lambda stat:stat["Obshchee_kol_zayavleniy"])
        dobavit("Количество мест",lambda stat:stat["Mesta"])
        for nomer in range(1,5):
            dobavit(f"Заявления {nomer}-го приоритета",lambda stat,n=nomer: stat[f"Zayavleniya_prior_{n}"])
        for nomer in range(1,5):
            dobavit(f"Зачислено {nomer}-го приоритета",lambda stat,n=nomer: stat[f"Zachisleno_prior_{n}"])
        ttk.Button(konteyner,text="Выход",command=okno.destroy).pack(pady=6)

    def _kod_programmy(self, prog:str)->str:
        return {"ПМ":"PM","ИВТ":"IVT","ИТСС":"ITSS","ИБ":"IB"}.get(prog,prog)

    def _sgenerirovat_pdf_dlya_programmy_i_daty(self,prog:str,metka_daty:str):
        shrift=self.imya_shrifta_pdf
        put_pdf_vyhoda=os.path.join(KATALOG_OTCHETOV,f"{self._kod_programmy(prog)}_{metka_daty}.pdf")
        holst=pdfcanvas.Canvas(put_pdf_vyhoda,pagesize=A4)
        shirina,vysota=A4
        holst.setFont(shrift,12)
        holst.drawString(2*cm,vysota-2*cm,f"{prog} — {metka_daty}")
        holst.setFont(shrift,10)
        holst.drawString(2*cm,vysota-2.7*cm,f"Сформировано: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        prokhodnoy_ball=self.vse_prokhodnye_bally.get(metka_daty,{}).get(prog)
        holst.setFont(shrift,10)
        holst.drawString(2*cm,vysota-3.5*cm,"Проходной балл:")
        holst.drawString(2*cm,vysota-4*cm,str("НЕДОБОР" if prokhodnoy_ball is None else prokhodnoy_ball))
        y_pozitsiya=vysota-4.8*cm
        holst.setFont(shrift,8)
        holst.drawString(2*cm,y_pozitsiya,"Список зачисленных (ID, сумма, приор.):")
        y_pozitsiya-=0.6*cm
        holst.setFont(shrift,8)
        for element in self.vse_rezultaty.get(metka_daty,{}).get(prog,[]):
            holst.drawString(2*cm,y_pozitsiya,f"ID {element['ID']} — {element['Сумма баллов']} (приор. {element['Приоритет']})")
            y_pozitsiya-=0.45*cm
            if y_pozitsiya<2.0*cm:
                holst.showPage()
                y_pozitsiya=vysota-2*cm
                holst.setFont(shrift,8)
        holst.save()

    def menyu_otchety_pdf(self):
        if not proverit_reportlab():
            return
        vybor=self.zaprosit_vybor_daty(zagolovok="Дата отчёта",priglashenie="Выберите дату:")
        if vybor is None:
            return
        if vybor not in self.sgenerovannye_daty:
            messagebox.showinfo("PDF","На дату нет данных")
            return
        for prog in PROGRAMMY:
            try:
                self._sgenerirovat_pdf_dlya_programmy_i_daty(prog,vybor)
            except Exception as oshibka:
                messagebox.showwarning("PDF",f"Ошибка {prog}: {oshibka}")
        messagebox.showinfo("PDF",f"Отчёты сохранены в {KATALOG_OTCHETOV}")

    def menyu_okno_grafika_istorii(self):
        if not proverit_matplotlib():
            return
        if not os.path.exists(FAYL_ISTORII):
            messagebox.showinfo("График", "history.csv отсутствует")
            return


        daty = []
        znacheniya = {prog: [] for prog in PROGRAMMY}
        with open(FAYL_ISTORII, "r", encoding="utf-8-sig", newline="") as f:
            stroki = list(csv.DictReader(f))
        stroki = sorted(
            stroki,
            key=lambda s: SPISOK_DAT.index(s["Дата"]) if s["Дата"] in SPISOK_DAT else 999
        )

        for s in stroki:
            data = s["Дата"]
            daty.append(data)
            for prog in PROGRAMMY:
                znach = s.get(prog)

                if znach in (None, "", "НЕДОБОР"):
                    znacheniya[prog].append(None)
                else:
                    znacheniya[prog].append(int(znach))


        fig, ax = plt.subplots(figsize=(8, 5))
        for prog in PROGRAMMY:
            ax.plot(daty, znacheniya[prog], marker="o", label=prog)

        ax.set_xlabel("Дата")
        ax.set_ylabel("Проходной балл")
        ax.set_title("Динамика проходных баллов")
        ax.legend()
        fig.tight_layout()


        okno = tk.Toplevel(self)
        okno.title("График")
        canvas = FigureCanvasTkAgg(fig, master=okno)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)


        panel = ttk.Frame(okno)
        panel.pack(fill=tk.X)

        def save_pdf():
            try:

                os.makedirs("report", exist_ok=True)
                put = os.path.join("report", "grafik_istorii.pdf")
                fig.savefig(put, format="pdf", bbox_inches="tight")
                messagebox.showinfo("Сохранение", f"График сохранён:\n{put}")
            except Exception as e:
                messagebox.showerror("Сохранение", f"Не удалось сохранить:\n{e}")

        ttk.Button(panel, text="Сохранить график", command=save_pdf).pack(side=tk.LEFT, padx=8, pady=8)
        ttk.Button(panel, text="Закрыть", command=okno.destroy).pack(side=tk.RIGHT, padx=8, pady=8)

    def _pokazat_tablitsu_csv_v_okne(self, put: str, zagolovok: str):
        if not os.path.exists(put):
            messagebox.showinfo(zagolovok, f"Файл {put} не найден")
            return
        stroki = self._prochitat_stroki_csv(put)
        if not stroki:
            messagebox.showinfo(zagolovok, "Файл пустой")
            return
        zagolovki = stroki[0]
        stroki_dannyh = stroki[1:]
        okno = tk.Toplevel(self)
        okno.title(zagolovok)
        okno.minsize(700, 450)
        ramka = ttk.Frame(okno)
        ramka.pack(fill=tk.BOTH, expand=True)
        derevo = ttk.Treeview(ramka, columns=zagolovki, show="headings")
        vertikalnaya_polosa = ttk.Scrollbar(ramka, orient="vertical", command=derevo.yview)
        gorizontalnaya_polosa = ttk.Scrollbar(ramka, orient="horizontal", command=derevo.xview)
        derevo.configure(yscroll=vertikalnaya_polosa.set, xscroll=gorizontalnaya_polosa.set)
        for kol in zagolovki:
            derevo.heading(kol, text=kol)
            derevo.column(kol, width=140, anchor=tk.CENTER)
        derevo.grid(row=0, column=0, sticky="nsew")
        vertikalnaya_polosa.grid(row=0, column=1, sticky="ns")
        gorizontalnaya_polosa.grid(row=1, column=0, sticky="ew")
        ramka.grid_rowconfigure(0, weight=1)
        ramka.grid_columnconfigure(0, weight=1)
        for stroka in stroki_dannyh:
            znacheniya = stroka + [""] * (len(zagolovki) - len(stroka))
            znacheniya = znacheniya[:len(zagolovki)]
            derevo.insert("", tk.END, values=znacheniya)
        ramka_knopok = ttk.Frame(okno)
        ramka_knopok.pack(fill=tk.X)
        ttk.Button(ramka_knopok, text="Закрыть", command=okno.destroy).pack(side=tk.RIGHT, padx=8, pady=8)

    def menyu_pokazat_otchet2(self):
        self._pokazat_tablitsu_csv_v_okne(os.path.join(KATALOG_ANALIZA, "otchet2.csv"), "Анализ генерации списков поступающих")

    def menyu_pokazat_otchet1(self):
        self._pokazat_tablitsu_csv_v_okne(os.path.join(KATALOG_ANALIZA, "otchet.csv"), "Число сгенерированных уникальных ID")

    def menyu_pokazat_tablitsu_istorii(self):
        self._pokazat_tablitsu_csv_v_okne(FAYL_ISTORII, "Динамика проходных баллов на ОП по дням")

    def _perestroit_menyu(self):
        self.config(menu=None)
        self._postroit_menyu()

    def menyu_konfig_programm(self):
        self._okno_konfiga_programm(nachalnoe=False)

    def _okno_konfiga_programm(self, nachalnoe:bool=False):
        global PROGRAMMY, MESTA
        okno=tk.Toplevel(self)
        okno.title("ОП и количество мест")
        okno.transient(self)
        okno.grab_set()
        ramka=ttk.Frame(okno,padding=10)
        ramka.pack(fill=tk.BOTH,expand=True)
        ttk.Label(ramka,text="ОП").grid(row=0,column=0)
        ttk.Label(ramka,text="Количество мест").grid(row=0,column=1)
        polya_prog=[]
        polya_mest=[]
        maks_strok=max(4,len(PROGRAMMY))
        for nomer in range(maks_strok):
            peremennaya_prog=tk.StringVar(value=PROGRAMMY[nomer] if nomer<len(PROGRAMMY) else "")
            peremennaya_mest=tk.StringVar(value=str(MESTA.get(PROGRAMMY[nomer],"")) if nomer<len(PROGRAMMY) else "")
            ttk.Entry(ramka,textvariable=peremennaya_prog,width=20).grid(row=nomer+1,column=0,padx=4,pady=2)
            ttk.Entry(ramka,textvariable=peremennaya_mest,width=10).grid(row=nomer+1,column=1,padx=4,pady=2)
            polya_prog.append(peremennaya_prog)
            polya_mest.append(peremennaya_mest)
        def sokhranit():
            programmy=[]
            mesta={}
            for peremennaya_prog,peremennaya_mest in zip(polya_prog,polya_mest):
                prog=peremennaya_prog.get().strip()
                mest=peremennaya_mest.get().strip()
                if prog:
                    if not mest.isdigit():
                        messagebox.showwarning("Ошибка","Места – число")
                        return
                    programmy.append(prog)
                    mesta[prog]=int(mest)
            if not programmy:
                messagebox.showwarning("Ошибка","Нужно указать хотя бы одну ОП")
                return
            PROGRAMMY[:]=programmy
            MESTA.clear()
            MESTA.update(mesta)
            sokhranit_konfig_programm(PROGRAMMY,MESTA)
            okno.destroy()
            self._perestroit_menyu()
        ttk.Button(ramka,text="Сохранить",command=sokhranit).grid(row=maks_strok+1,column=0,columnspan=2,pady=6)
        if nachalnoe:
            self.wait_window(okno)

    def menyu_konfig_dat(self):
        self._okno_konfiga_dat(nachalnoe=False)

    def _okno_konfiga_dat(self, nachalnoe:bool=False):
        global SPISOK_DAT
        okno=tk.Toplevel(self)
        okno.title("Даты приёма")
        okno.transient(self)
        okno.grab_set()
        ramka=tk.Frame(okno)
        ramka.pack(fill=tk.BOTH, expand=True)
        vnutrennaya=ttk.Frame(ramka,padding=10)
        vnutrennaya.pack(fill=tk.BOTH,expand=True)
        ttk.Label(vnutrennaya,text="Введите даты (ДД.ММ)").pack(anchor="w")
        polya=[]
        maks_strok=max(4,len(SPISOK_DAT))
        for nomer in range(maks_strok):
            peremennaya=tk.StringVar(value=SPISOK_DAT[nomer] if nomer<len(SPISOK_DAT) else "")
            pole=ttk.Entry(vnutrennaya,textvariable=peremennaya,width=10)
            pole.pack(pady=2,anchor="w")
            polya.append(peremennaya)
        def sokhranit():
            daty=[peremennaya.get().strip() for peremennaya in polya if peremennaya.get().strip()]
            if not daty:
                messagebox.showwarning("Ошибка","Укажите хотя бы одну дату")
                return
            SPISOK_DAT[:]=daty
            sokhranit_konfig_dat(SPISOK_DAT)
            okno.destroy()
            self._perestroit_menyu()
        ttk.Button(vnutrennaya,text="Сохранить",command=sokhranit).pack(pady=6)
        if nachalnoe:
            self.wait_window(okno)

def df_bool(seriya: "pd.Series") -> "pd.Series":
    try:
        return seriya.astype("boolean") == True
    except Exception:
        return seriya == True

def privetstvennoe_okno():
    koren = tk.Tk()
    koren.title("Приемная компания")
    koren.resizable(False, False)

    zagolovok = ttk.Label(koren, text="Приемная компания", font=("Segoe UI", 30, "bold"))
    zagolovok.pack(side=tk.TOP, fill=tk.X, padx=20, pady=(16, 12))

    soderzhanie = ttk.Frame(koren)
    soderzhanie.pack(fill=tk.BOTH, expand=True, padx=16)

    try:
        vysota_stroki = tkfont.nametofont("TkDefaultFont").metrics("linespace")
        if not vysota_stroki:
            vysota_stroki = 18
    except Exception:
        vysota_stroki = 18
    maksimalnaya_vysota = int(vysota_stroki * 15)

    metka_kartinki = ttk.Label(soderzhanie)
    metka_kartinki.grid(row=0, column=0, sticky="n", padx=(0, 16))

    obekt_kartinki = None
    put = KARTINKA_FONA_PO_UMOLCHANIYU if os.path.exists(KARTINKA_FONA_PO_UMOLCHANIYU) else "1.png"
    if os.path.exists(put):
        try:
            if PIL_DOSTUPEN:
                kartinka = Image.open(put)
                shirina0, vysota0 = kartinka.size
                if vysota0 > maksimalnaya_vysota:
                    koeffitsient = maksimalnaya_vysota / float(vysota0)
                    novaya_shirina = max(1, int(shirina0 * koeffitsient))
                    kartinka = kartinka.resize((novaya_shirina, maksimalnaya_vysota))
                obekt_kartinki = ImageTk.PhotoImage(kartinka)
                metka_kartinki.configure(image=obekt_kartinki)
                metka_kartinki.image = obekt_kartinki
            else:
                kartinka = tk.PhotoImage(file=put)
                vysota0 = kartinka.height()
                if vysota0 > maksimalnaya_vysota:
                    koef = int((vysota0 + maksimalnaya_vysota - 1) // maksimalnaya_vysota)
                    if koef < 1:
                        koef = 1
                    kartinka = kartinka.subsample(koef, koef)
                metka_kartinki.configure(image=kartinka)
                metka_kartinki.image = kartinka
        except Exception as oshibka:
            metka_kartinki.configure(text=f"Не удалось загрузить 1.png: {oshibka}")
    else:
        metka_kartinki.configure(text="1.png не найден")

    pravaya_kolonka = ttk.Frame(soderzhanie)
    pravaya_kolonka.grid(row=0, column=1, sticky="n")

    ttk.Label(pravaya_kolonka, text="Университетский лицей 1511").pack(anchor="w")
    ttk.Label(pravaya_kolonka, text="Предуниверситария НИЯУ МИФИ").pack(anchor="w")
    ttk.Label(pravaya_kolonka, text="").pack(anchor="w")
    ttk.Label(pravaya_kolonka, text="Разработчики:").pack(anchor="w")
    ttk.Label(pravaya_kolonka, text="Артём Ефремцев  10Д").pack(anchor="w")
    ttk.Label(pravaya_kolonka, text="Артём Ананьев  11Л").pack(anchor="w")
    ttk.Label(pravaya_kolonka, text="Георгий Соловей 11Л").pack(anchor="w")
    ttk.Label(pravaya_kolonka, text="").pack(anchor="w")
    ttk.Label(pravaya_kolonka, text="").pack(anchor="w")
    ttk.Label(pravaya_kolonka, text="Москва").pack(anchor="w")
    ttk.Label(pravaya_kolonka, text="2026").pack(anchor="w")

    ramka_knopki = ttk.Frame(koren)
    ramka_knopki.pack(fill=tk.X, pady=(12, 16))
    def idti():
        try:
            koren.destroy()
        except Exception:
            pass
    knopka_ok = ttk.Button(ramka_knopki, text="ОК", command=idti)
    knopka_ok.pack()
    koren.bind("<Return>", lambda e: idti())

    koren.update_idletasks()
    shirina = koren.winfo_reqwidth()
    vysota = koren.winfo_reqheight()
    shirina_ekrana = koren.winfo_screenwidth()
    vysota_ekrana = koren.winfo_screenheight()
    x = (shirina_ekrana - shirina) // 2
    y = (vysota_ekrana - vysota) // 3
    koren.geometry(f"{shirina}x{vysota}+{x}+{y}")

    koren.mainloop()

def glavnyi():
    privetstvennoe_okno()
    prilozhenie = Prilozhenie()
    prilozhenie.mainloop()

if __name__=="__main__":
    glavnyi()
