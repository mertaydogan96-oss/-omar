# -*- coding: utf-8 -*-
"""
Başvuru XML + Excel Aktarıcı (Deneme Sürümü)

Amaç:
- Var olan XML giriş ekranını ana mantık olarak korumak
- Aynı kalem listesinden ikinci sekmede Excel çıktısı üretmek
- XML ve Excel için aynı veriyi ikinci kez yazma ihtiyacını azaltmak

Not:
- Excel tarafı ilk sürümdür; bazı kolon eşleşmeleri varsayımsaldır.
- Kullanıcı dostu temel akış hedeflenmiştir.
"""

import copy
import os
import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import tkinter.simpledialog as simpledialog
import xml.etree.ElementTree as ET
from xml.dom import minidom
from datetime import date, datetime

try:
    from openpyxl import Workbook
except Exception:
    Workbook = None


VARSAYILAN_TASLAK_ADI = "P-AQRO"

VARSAYILAN_GENEL = {
    "gondericiVergiNo": "",
    "gondericiUnvan": "P-AQRO MMC",
    "sevkUlkesi": "078",
    "ticaretYapilanUlke": "078",
    "cikisUlkesi": "078",
    "sinirdakiTasimaSekli": "30",
    "esyaninBulunduguYer": "MALATYA",
    "girisGumruk": "440100",
    "bosaltmaGumruk": "440100",
}

K_SABIT = {
    "gtip": "520100900019",
    "belgeTipi": "4",
    "menseUlke": "078",
    "degerBirimi": "USD",
    "miktarBirimi": "KGM",
    "urunGrubu": "ORTA ELYAFLI (UPLAND)",
    "grubu": "ORTA ELYAFLI (UPLAND)",
    "urunCinsi": "LİF PAMUK",
    "imalatciBaslik": "İMALATÇI UNVANI",
    "imalatciDeger": "P-AQRO MMC",
}

EXCEL_HEADERS = [
    "GTİP No", "Ticari Tanım", "Menşei", "Fatura Tutar", "Satış Miktarı", "Birim",
    "Brüt KG", "Net KG", "Tamamlayıcı Ölçü Miktar", "Tamamlayıcı Ölçü Birim",
    "Kap Marka", "Kap Numara", "Kap Adet", "Kap Cinsi", "DİR Satır Kodu",
    "Kul.Eşya Kod (K1/K2/K3)", "A-Damping Firma Kodu", "Özellik (02:Bedelsiz)",
    "Muafiyet1", "Muafiyet2", "Muafiyet3", "Anlaşma Kod",
    "Tercihli KDV Kod (KDV1/KDV8/KDV18)", "Tarım Payı Ek Kod",
    "Açıklama(44.hane ilave)", "Mahrece İade Mi?(E)", "İkincil Ürün Mü? (E)",
    "Algı.Birim-1", "Algı.Miktar-1", "Algı.Birim-2", "Algı.Miktar-2",
    "Algı.Birim-3", "Algı.Miktar-3", "Navlun", "Sigorta", "İlave Yurtdışı",
    "İlave Yurtdışı-Açıklama", "KKDF", "YİÇİ-Banka", "YİÇİ-Depolama",
    "YİÇİ-Tahliye", "YİÇİ-Liman", "KültürF", "ÇevreKP", "YİÇİ-Diğer",
    "YİÇİ-Diğer-Açıklama", "Konteyner No", "Konteyner Ülke Kodu", "Fatura No",
    "Fatura Tarihi", "Tamamlayıcı Bilgi Kod", "Tamamlayıcı Bilgi Değer",
    "TPS Belge Kodu", "TPS Belge No", "TPS Belge Tarihi", "Marka Tescilli (E)",
    "Marka Tescil No", "Marka Adı", "Kıymeti-TL", "Ref / Şasi No", "Model Yılı",
    "Model", "Motor Hacmi", "Silindir Adedi", "Renk",
    "Vites (1:Otomatik/2:Düz/3:Belirtilmemiş)", "Motor Gücü",
    "Motor Tipi (1:Akaryakıtlı/2:Elektrikli/3:Hibrit)", "Motor No",
    "Ödeme Şekli", "Tutar", "TBF ID", "TBF Tarih"
]

DB_DOSYA_ADI = "app_data.db"
BOOL_AYAR_ANAHTARLARI = {
    "genel_duzenle_var",
    "sabitleri_goster_var",
    "sabit_duzenle_var",
    "belge_fatura_esit_var",
    "no_otomatik_artsin_var",
    "tarih_sabit_kalsin_var",
}

# Excel tarafında kullanıcıya gösterilmeyen ama arka planda sabit doldurulan alanlar
EXCEL_SABIT_DOLU = {
    "GTİP No": "5201.00.90.00.19",
    "Ticari Tanım": "PAMUK",
    "Menşei": ("sabit_var", "menseUlke"),
    "Birim": ("sabit_var", "miktarBirimi"),
    "Tamamlayıcı Ölçü Birim": ("sabit_var", "miktarBirimi"),
    "Kap Marka": "ADDR",
    "Kap Numara": ".",
    "Kap Cinsi": "BL",
    "Kul.Eşya Kod (K1/K2/K3)": "K1",
    "Özellik (02:Bedelsiz)": "01",
    "Muafiyet1": "KKDFM",
    "Marka Adı": "",  
    "Konteyner Ülke Kodu": "",
    "İlave Yurtdışı-Açıklama": "KANTAR FARKI YURTDIŞI DİĞER DE BEYAN EDİLMİŞTİR.",
}

# Excel tarafında bilinçli olarak boş bırakılan kolonlar
EXCEL_SABIT_BOS = {
    "DİR Satır Kodu",
    "A-Damping Firma Kodu",
    "Muafiyet2",
    "Muafiyet3",
    "Anlaşma Kod",
    "Tercihli KDV Kod (KDV1/KDV8/KDV18)",
    "Tarım Payı Ek Kod",
    "Açıklama(44.hane ilave)",
    "Mahrece İade Mi?(E)",
    "İkincil Ürün Mü? (E)",
    "Algı.Birim-1",
    "Algı.Miktar-1",
    "Algı.Birim-2",
    "Algı.Miktar-2",
    "Algı.Birim-3",
    "Algı.Miktar-3",
    "KKDF",
    "YİÇİ-Banka",
    "YİÇİ-Depolama",
    "YİÇİ-Tahliye",
    "YİÇİ-Liman",
    "KültürF",
    "ÇevreKP",
    "YİÇİ-Diğer",
    "YİÇİ-Diğer-Açıklama",
    "Konteyner No",
    "Tamamlayıcı Bilgi Kod",
    "Tamamlayıcı Bilgi Değer",
    "TPS Belge Kodu",
    "TPS Belge No",
    "TPS Belge Tarihi",
    "Marka Tescilli (E)",
    "Marka Tescil No",
    "Kıymeti-TL",
    "Ref / Şasi No",
    "Model Yılı",
    "Model",
    "Motor Hacmi",
    "Silindir Adedi",
    "Renk",
    "Vites (1:Otomatik/2:Düz/3:Belirtilmemiş)",
    "Motor Gücü",
    "Motor Tipi (1:Akaryakıtlı/2:Elektrikli/3:Hibrit)",
    "Motor No",
    "Tutar",
    "TBF ID",
    "TBF Tarih",
}

EXCEL_MANUEL_ALANLAR = ("brutKg", "navlun", "sigorta", "ilaveYurtdisi")
BELGE_KOD_TANIMLARI = [
    ("0100", "Fatura", "faturaNo", "faturaTarihi", True),
    ("0101", "Navlun Fatura", "belge0101Ref", "belge0101Tarih", False),
    ("0102", "Sigorta Poliçe", "belge0102Ref", "belge0102Tarih", False),
    ("0979", "TAREKS", "belge0979Ref", "belge0979Tarih", False),
    ("0903", "TPS", "belge0903Ref", "belge0903Tarih", False),
    ("0877", "İTKİB", "belge0877Ref", "belge0877Tarih", False),
]

# Yeni yardımcı eşleme: belge koduna göre hangi DB alanlarının güncelleneceği
BELGE_ALAN_HARITASI = {
    "0100": ("faturaNo", "faturaTarihi"),
    "0101": ("belge0101Ref", "belge0101Tarih"),
    "0102": ("belge0102Ref", "belge0102Tarih"),
    "0979": ("belge0979Ref", "belge0979Tarih"),
    "0903": ("belge0903Ref", "belge0903Tarih"),
    "0877": ("belge0877Ref", "belge0877Tarih"),
}

# Hızlı lookup hangi belge kodu readonly (ör. fatura otomatik)
BELGE_READONLY = {kod: readonly for kod, _, _, _, readonly in BELGE_KOD_TANIMLARI}


def bugunun_tarihi_gosterim():
    return date.today().strftime("%d.%m.%Y")


def bugunun_tarihi_xml():
    return date.today().strftime("%Y-%m-%d")


def tarihi_xml_formatina_cevir(metin: str) -> str:
    metin = (metin or "").strip()
    if not metin:
        return ""
    parca = metin.split(".")
    if len(parca) != 3:
        return metin
    gun, ay, yil = parca
    return f"{yil}-{ay}-{gun}"


def tarihi_belge_excel_formatina_cevir(metin: str) -> str:
    metin = (metin or "").strip()
    if not metin:
        return ""
    parca = metin.split(".")
    if len(parca) == 3:
        gun, ay, yil = parca
        return f"{gun}/{ay}/{yil}"
    return metin.replace(".", "/")


def belge_numarasi_artir(metin: str) -> str:
    metin = (metin or "").strip()
    if not metin:
        return metin

    if metin.endswith(")") and "(" in metin:
        sol, sag = metin.rsplit("(", 1)
        sayi = sag[:-1]
        if sayi.isdigit():
            return f"{sol}({int(sayi) + 1})"

    i = len(metin) - 1
    while i >= 0 and metin[i].isdigit():
        i -= 1
    if i < len(metin) - 1:
        sayi = metin[i + 1:]
        prefix = metin[:i + 1]
        try:
            yeni = str(int(sayi) + 1).zfill(len(sayi))
            return prefix + yeni
        except Exception:
            return metin

    return metin


class PersistenceManager:
    """Uygulama ayarlarını ve kalemleri sqlite üzerinde saklar."""

    def __init__(self, db_path: str):
        self.db_path = db_path
        self._bekleyen_uyari = ""

    def _connect(self):
        conn = sqlite3.connect(self.db_path, timeout=10)
        conn.row_factory = sqlite3.Row
        return conn

    def _schema_olustur(self, conn):
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sira INTEGER NOT NULL,
                belge_no TEXT,
                fatura_no TEXT,
                fatura_tarihi TEXT,
                deger TEXT,
                miktar TEXT,
                tanim TEXT,
                urun_grubu TEXT,
                balya_sayisi TEXT,
                xml_group_no TEXT DEFAULT '',
                excel_group_no TEXT DEFAULT '',
                durum TEXT DEFAULT 'Hazır',
                brut_kg TEXT DEFAULT '',
                navlun TEXT DEFAULT '',
                sigorta TEXT DEFAULT '',
                ilave_yurtdisi TEXT DEFAULT '',
                excel_secili INTEGER DEFAULT 0,
                belge_secili INTEGER DEFAULT 0,
                belge_0101_ref TEXT DEFAULT '',
                belge_0101_tarih TEXT DEFAULT '',
                belge_0102_ref TEXT DEFAULT '',
                belge_0102_tarih TEXT DEFAULT '',
                belge_0979_ref TEXT DEFAULT '',
                belge_0979_tarih TEXT DEFAULT '',
                belge_0903_ref TEXT DEFAULT '',
                belge_0903_tarih TEXT DEFAULT '',
                belge_0877_ref TEXT DEFAULT '',
                belge_0877_tarih TEXT DEFAULT '',
                created_at TEXT,
                updated_at TEXT
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS seferler (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ad TEXT NOT NULL,
                durum TEXT DEFAULT 'aktif',
                created_at TEXT,
                updated_at TEXT
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS kalemler (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sefer_id INTEGER NOT NULL,
                sira INTEGER NOT NULL,
                belge_no TEXT,
                fatura_no TEXT,
                fatura_tarihi TEXT,
                deger TEXT,
                miktar TEXT,
                tanim TEXT,
                urun_grubu TEXT,
                balya_sayisi TEXT,
                brut_kg TEXT DEFAULT '',
                navlun TEXT DEFAULT '',
                sigorta TEXT DEFAULT '',
                ilave_yurtdisi TEXT DEFAULT '',
                excel_secili INTEGER DEFAULT 0,
                belge_secili INTEGER DEFAULT 0,
                belge_0101_ref TEXT DEFAULT '',
                belge_0101_tarih TEXT DEFAULT '',
                belge_0102_ref TEXT DEFAULT '',
                belge_0102_tarih TEXT DEFAULT '',
                belge_0979_ref TEXT DEFAULT '',
                belge_0979_tarih TEXT DEFAULT '',
                belge_0903_ref TEXT DEFAULT '',
                belge_0903_tarih TEXT DEFAULT '',
                belge_0877_ref TEXT DEFAULT '',
                belge_0877_tarih TEXT DEFAULT '',
                durum TEXT DEFAULT 'bekliyor',
                tamamlanma_grup_id INTEGER DEFAULT NULL,
                created_at TEXT,
                updated_at TEXT
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS tamamlanma_gruplari (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sefer_id INTEGER NOT NULL,
                grup_no TEXT NOT NULL,
                notlar TEXT DEFAULT '',
                created_at TEXT,
                updated_at TEXT
            )
            """
        )

        # Eski veritabanı dosyaları için kontrollü kolon göçü (records tablosu)
        mevcut_kolonlar = {row["name"] for row in conn.execute("PRAGMA table_info(records)").fetchall()}
        eksik_kolonlar = {
            "xml_group_no": "TEXT DEFAULT ''",
            "excel_group_no": "TEXT DEFAULT ''",
            "durum": "TEXT DEFAULT 'Hazır'",
            "brut_kg": "TEXT DEFAULT ''",
            "navlun": "TEXT DEFAULT ''",
            "sigorta": "TEXT DEFAULT ''",
            "ilave_yurtdisi": "TEXT DEFAULT ''",
            "excel_secili": "INTEGER DEFAULT 0",
            "belge_secili": "INTEGER DEFAULT 0",
            "belge_0101_ref": "TEXT DEFAULT ''",
            "belge_0101_tarih": "TEXT DEFAULT ''",
            "belge_0102_ref": "TEXT DEFAULT ''",
            "belge_0102_tarih": "TEXT DEFAULT ''",
            "belge_0979_ref": "TEXT DEFAULT ''",
            "belge_0979_tarih": "TEXT DEFAULT ''",
            "belge_0903_ref": "TEXT DEFAULT ''",
            "belge_0903_tarih": "TEXT DEFAULT ''",
            "belge_0877_ref": "TEXT DEFAULT ''",
            "belge_0877_tarih": "TEXT DEFAULT ''",
            "created_at": "TEXT",
            "updated_at": "TEXT",
        }
        for kolon, tanim in eksik_kolonlar.items():
            if kolon not in mevcut_kolonlar:
                conn.execute(f"ALTER TABLE records ADD COLUMN {kolon} {tanim}")

        # Migration: seferler tablosu boşsa records'tan geç
        sefer_sayisi = conn.execute("SELECT COUNT(*) FROM seferler").fetchone()[0]
        if sefer_sayisi == 0:
            simdi = datetime.now().isoformat(timespec="seconds")
            records_sayisi = conn.execute("SELECT COUNT(*) FROM records").fetchone()[0]
            if records_sayisi > 0:
                cur = conn.execute(
                    "INSERT INTO seferler (ad, durum, created_at, updated_at) VALUES (?, 'aktif', ?, ?)",
                    ("Eski Kayıtlar", simdi, simdi),
                )
                sefer_id = cur.lastrowid
                rows = conn.execute(
                    """
                    SELECT sira, belge_no, fatura_no, fatura_tarihi, deger, miktar, tanim,
                           urun_grubu, balya_sayisi, brut_kg, navlun, sigorta, ilave_yurtdisi,
                           excel_secili, belge_secili,
                           belge_0101_ref, belge_0101_tarih, belge_0102_ref, belge_0102_tarih,
                           belge_0979_ref, belge_0979_tarih, belge_0903_ref, belge_0903_tarih,
                           belge_0877_ref, belge_0877_tarih, created_at, updated_at
                    FROM records ORDER BY sira ASC, id ASC
                    """
                ).fetchall()
                for row in rows:
                    conn.execute(
                        """
                        INSERT INTO kalemler (
                            sefer_id, sira, belge_no, fatura_no, fatura_tarihi, deger, miktar, tanim,
                            urun_grubu, balya_sayisi, brut_kg, navlun, sigorta, ilave_yurtdisi,
                            excel_secili, belge_secili,
                            belge_0101_ref, belge_0101_tarih, belge_0102_ref, belge_0102_tarih,
                            belge_0979_ref, belge_0979_tarih, belge_0903_ref, belge_0903_tarih,
                            belge_0877_ref, belge_0877_tarih, durum, tamamlanma_grup_id,
                            created_at, updated_at
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            sefer_id,
                            row["sira"],
                            row["belge_no"] or "",
                            row["fatura_no"] or "",
                            row["fatura_tarihi"] or "",
                            row["deger"] or "",
                            row["miktar"] or "",
                            row["tanim"] or "",
                            row["urun_grubu"] or "",
                            row["balya_sayisi"] or "",
                            row["brut_kg"] or "",
                            row["navlun"] or "",
                            row["sigorta"] or "",
                            row["ilave_yurtdisi"] or "",
                            row["excel_secili"] or 0,
                            row["belge_secili"] or 0,
                            row["belge_0101_ref"] or "",
                            row["belge_0101_tarih"] or "",
                            row["belge_0102_ref"] or "",
                            row["belge_0102_tarih"] or "",
                            row["belge_0979_ref"] or "",
                            row["belge_0979_tarih"] or "",
                            row["belge_0903_ref"] or "",
                            row["belge_0903_tarih"] or "",
                            row["belge_0877_ref"] or "",
                            row["belge_0877_tarih"] or "",
                            "bekliyor",
                            None,
                            row["created_at"] or simdi,
                            simdi,
                        ),
                    )
            else:
                cur = conn.execute(
                    "INSERT INTO seferler (ad, durum, created_at, updated_at) VALUES (?, 'aktif', ?, ?)",
                    ("Varsayılan Sefer", simdi, simdi),
                )
                sefer_id = cur.lastrowid
            conn.execute(
                "INSERT OR REPLACE INTO settings (key, value) VALUES ('active_sefer_id', ?)",
                (str(sefer_id),),
            )

        conn.commit()

    def _bozuk_veritabanini_yedekle(self, exc: Exception):
        damga = datetime.now().strftime("%Y%m%d_%H%M%S")
        yedek_yol = self.db_path.replace(".db", f"_bozuk_{damga}.db")

        if os.path.exists(self.db_path):
            os.replace(self.db_path, yedek_yol)

        for ek in ("-wal", "-shm"):
            yan_dosya = self.db_path + ek
            if os.path.exists(yan_dosya):
                try:
                    os.remove(yan_dosya)
                except OSError:
                    pass

        self._bekleyen_uyari = (
            "Veritabanı dosyasında sorun algılandı ve temiz veritabanı oluşturuldu.\n"
            f"Bozuk kopya: {yedek_yol}\n\nDetay: {exc}"
        )

    def _db_isle(self, islem, allow_recovery=True):
        try:
            with self._connect() as conn:
                self._schema_olustur(conn)
                return islem(conn)
        except sqlite3.DatabaseError as exc:
            if not allow_recovery:
                raise
            self._bozuk_veritabanini_yedekle(exc)
            return self._db_isle(islem, allow_recovery=False)

    def init_db(self):
        def islem(conn):
            sonuc = conn.execute("PRAGMA integrity_check").fetchone()
            if sonuc and str(sonuc[0]).lower() != "ok":
                raise sqlite3.DatabaseError(f"integrity_check sonucu: {sonuc[0]}")
            return True

        return self._db_isle(islem)

    def save_settings(self, settings_dict):
        kayitlar = [(str(key), "1" if value is True else "0" if value is False else str(value)) for key, value in settings_dict.items()]

        def islem(conn):
            conn.executemany(
                "INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)",
                kayitlar,
            )
            conn.commit()
            return True

        return self._db_isle(islem)

    def load_settings(self):
        def islem(conn):
            rows = conn.execute("SELECT key, value FROM settings").fetchall()
            return {row["key"]: row["value"] for row in rows}

        return self._db_isle(islem)

    def upsert_record(self, kalem):
        simdi = datetime.now().isoformat(timespec="seconds")
        created_at = kalem.get("created_at") or simdi

        def islem(conn):
            if kalem.get("db_id"):
                conn.execute(
                    """
                    UPDATE records
                    SET sira = ?, belge_no = ?, fatura_no = ?, fatura_tarihi = ?,
                        deger = ?, miktar = ?, tanim = ?, urun_grubu = ?, balya_sayisi = ?,
                        xml_group_no = ?, excel_group_no = ?, durum = ?, brut_kg = ?,
                        navlun = ?, sigorta = ?, ilave_yurtdisi = ?, excel_secili = ?, belge_secili = ?,
                        belge_0101_ref = ?, belge_0101_tarih = ?, belge_0102_ref = ?, belge_0102_tarih = ?,
                        belge_0979_ref = ?, belge_0979_tarih = ?, belge_0903_ref = ?, belge_0903_tarih = ?,
                        belge_0877_ref = ?, belge_0877_tarih = ?, created_at = ?, updated_at = ?
                    WHERE id = ?
                    """,
                    (
                        kalem.get("sira", 0),
                        kalem.get("belgeNo", ""),
                        kalem.get("faturaNo", ""),
                        kalem.get("faturaTarihi", ""),
                        kalem.get("deger", ""),
                        kalem.get("miktar", ""),
                        kalem.get("tanim", ""),
                        kalem.get("urunGrubu", ""),
                        kalem.get("balyaSayisi", ""),
                        kalem.get("xmlGroupNo", ""),
                        kalem.get("excelGroupNo", ""),
                        kalem.get("durum", "Hazır"),
                        kalem.get("brutKg", ""),
                        kalem.get("navlun", ""),
                        kalem.get("sigorta", ""),
                        kalem.get("ilaveYurtdisi", ""),
                        1 if kalem.get("excelSecili") else 0,
                        1 if kalem.get("belgeSecili") else 0,
                        kalem.get("belge0101Ref", ""),
                        kalem.get("belge0101Tarih", ""),
                        kalem.get("belge0102Ref", ""),
                        kalem.get("belge0102Tarih", ""),
                        kalem.get("belge0979Ref", ""),
                        kalem.get("belge0979Tarih", ""),
                        kalem.get("belge0903Ref", ""),
                        kalem.get("belge0903Tarih", ""),
                        kalem.get("belge0877Ref", ""),
                        kalem.get("belge0877Tarih", ""),
                        created_at,
                        simdi,
                        kalem["db_id"],
                    ),
                )
                kayit_id = kalem["db_id"]
            else:
                cur = conn.execute(
                    """
                    INSERT INTO records (
                        sira, belge_no, fatura_no, fatura_tarihi, deger, miktar, tanim,
                        urun_grubu, balya_sayisi, xml_group_no, excel_group_no, durum,
                        brut_kg, navlun, sigorta, ilave_yurtdisi, excel_secili, belge_secili,
                        belge_0101_ref, belge_0101_tarih, belge_0102_ref, belge_0102_tarih,
                        belge_0979_ref, belge_0979_tarih, belge_0903_ref, belge_0903_tarih,
                        belge_0877_ref, belge_0877_tarih, created_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        kalem.get("sira", 0),
                        kalem.get("belgeNo", ""),
                        kalem.get("faturaNo", ""),
                        kalem.get("faturaTarihi", ""),
                        kalem.get("deger", ""),
                        kalem.get("miktar", ""),
                        kalem.get("tanim", ""),
                        kalem.get("urunGrubu", ""),
                        kalem.get("balyaSayisi", ""),
                        kalem.get("xmlGroupNo", ""),
                        kalem.get("excelGroupNo", ""),
                        kalem.get("durum", "Hazır"),
                        kalem.get("brutKg", ""),
                        kalem.get("navlun", ""),
                        kalem.get("sigorta", ""),
                        kalem.get("ilaveYurtdisi", ""),
                        1 if kalem.get("excelSecili") else 0,
                        1 if kalem.get("belgeSecili") else 0,
                        kalem.get("belge0101Ref", ""),
                        kalem.get("belge0101Tarih", ""),
                        kalem.get("belge0102Ref", ""),
                        kalem.get("belge0102Tarih", ""),
                        kalem.get("belge0979Ref", ""),
                        kalem.get("belge0979Tarih", ""),
                        kalem.get("belge0903Ref", ""),
                        kalem.get("belge0903Tarih", ""),
                        kalem.get("belge0877Ref", ""),
                        kalem.get("belge0877Tarih", ""),
                        created_at,
                        simdi,
                    ),
                )
                kayit_id = cur.lastrowid
            conn.commit()
            return kayit_id, created_at, simdi

        kayit_id, created_at, updated_at = self._db_isle(islem)
        kalem["db_id"] = kayit_id
        kalem["created_at"] = created_at
        kalem["updated_at"] = updated_at
        return kayit_id

    def delete_record(self, db_id):
        if not db_id:
            return False

        def islem(conn):
            conn.execute("DELETE FROM records WHERE id = ?", (db_id,))
            conn.commit()
            return True

        return self._db_isle(islem)

    def save_records(self, kalemler):
        aktif_idler = []
        for sira, kalem in enumerate(kalemler, start=1):
            kalem["sira"] = sira
            aktif_idler.append(self.upsert_record(kalem))

        def islem(conn):
            if aktif_idler:
                yerler = ",".join("?" for _ in aktif_idler)
                conn.execute(f"DELETE FROM records WHERE id NOT IN ({yerler})", tuple(aktif_idler))
            else:
                conn.execute("DELETE FROM records")
            conn.commit()
            return True

        return self._db_isle(islem)

    def load_records(self):
        def islem(conn):
            rows = conn.execute(
                """
                SELECT id, sira, belge_no, fatura_no, fatura_tarihi, deger, miktar, tanim,
                       urun_grubu, balya_sayisi, xml_group_no, excel_group_no, durum,
                       brut_kg, navlun, sigorta, ilave_yurtdisi, excel_secili, belge_secili,
                       belge_0101_ref, belge_0101_tarih, belge_0102_ref, belge_0102_tarih,
                       belge_0979_ref, belge_0979_tarih, belge_0903_ref, belge_0903_tarih,
                       belge_0877_ref, belge_0877_tarih, created_at, updated_at
                FROM records
                ORDER BY sira ASC, id ASC
                """
            ).fetchall()

            kalemler = []
            for row in rows:
                kalemler.append(
                    {
                        "db_id": row["id"],
                        "sira": row["sira"],
                        "belgeNo": row["belge_no"] or "",
                        "faturaNo": row["fatura_no"] or "",
                        "faturaTarihi": row["fatura_tarihi"] or "",
                        "deger": row["deger"] or "",
                        "miktar": row["miktar"] or "",
                        "tanim": row["tanim"] or "",
                        "urunGrubu": row["urun_grubu"] or "",
                        "balyaSayisi": row["balya_sayisi"] or "",
                        "xmlGroupNo": row["xml_group_no"] or "",
                        "excelGroupNo": row["excel_group_no"] or "",
                        "durum": row["durum"] or "Hazır",
                        "brutKg": row["brut_kg"] or "",
                        "navlun": row["navlun"] or "",
                        "sigorta": row["sigorta"] or "",
                        "ilaveYurtdisi": row["ilave_yurtdisi"] or "",
                        "excelSecili": 1 if row["excel_secili"] else 0,
                        "belgeSecili": 1 if row["belge_secili"] else 0,
                        "belge0101Ref": row["belge_0101_ref"] or "",
                        "belge0101Tarih": row["belge_0101_tarih"] or "",
                        "belge0102Ref": row["belge_0102_ref"] or "",
                        "belge0102Tarih": row["belge_0102_tarih"] or "",
                        "belge0979Ref": row["belge_0979_ref"] or "",
                        "belge0979Tarih": row["belge_0979_tarih"] or "",
                        "belge0903Ref": row["belge_0903_ref"] or "",
                        "belge0903Tarih": row["belge_0903_tarih"] or "",
                        "belge0877Ref": row["belge_0877_ref"] or "",
                        "belge0877Tarih": row["belge_0877_tarih"] or "",
                        "created_at": row["created_at"] or "",
                        "updated_at": row["updated_at"] or "",
                    }
                )
            return kalemler

        return self._db_isle(islem)

    def save_all_state(self, settings_dict, kalemler):
        self.save_settings(settings_dict)
        self.save_records(kalemler)
        return True

    def load_all_state(self):
        return self.load_settings(), self.load_records()

    def bekleyen_uyariyi_al(self):
        mesaj = self._bekleyen_uyari
        self._bekleyen_uyari = ""
        return mesaj

    # ------------------------------------------------------------------
    # Sefer yönetimi
    # ------------------------------------------------------------------
    def create_sefer(self, ad: str) -> int:
        simdi = datetime.now().isoformat(timespec="seconds")

        def islem(conn):
            cur = conn.execute(
                "INSERT INTO seferler (ad, durum, created_at, updated_at) VALUES (?, 'aktif', ?, ?)",
                (ad, simdi, simdi),
            )
            conn.commit()
            return cur.lastrowid

        return self._db_isle(islem)

    def list_seferler(self, tumu=False) -> list:
        def islem(conn):
            if tumu:
                rows = conn.execute("SELECT * FROM seferler ORDER BY id ASC").fetchall()
            else:
                rows = conn.execute(
                    "SELECT * FROM seferler WHERE durum='aktif' ORDER BY id ASC"
                ).fetchall()
            return [dict(row) for row in rows]

        return self._db_isle(islem)

    def archive_sefer(self, sefer_id: int):
        simdi = datetime.now().isoformat(timespec="seconds")

        def islem(conn):
            conn.execute(
                "UPDATE seferler SET durum='arsiv', updated_at=? WHERE id=?",
                (simdi, sefer_id),
            )
            conn.commit()
            return True

        return self._db_isle(islem)

    def get_active_sefer_id(self):
        def islem(conn):
            row = conn.execute(
                "SELECT value FROM settings WHERE key='active_sefer_id'"
            ).fetchone()
            if row and row["value"]:
                try:
                    return int(row["value"])
                except (ValueError, TypeError):
                    return None
            return None

        return self._db_isle(islem)

    def set_active_sefer_id(self, sefer_id: int):
        def islem(conn):
            conn.execute(
                "INSERT OR REPLACE INTO settings (key, value) VALUES ('active_sefer_id', ?)",
                (str(sefer_id),),
            )
            conn.commit()
            return True

        return self._db_isle(islem)

    # ------------------------------------------------------------------
    # Kalem yönetimi (yeni tablo)
    # ------------------------------------------------------------------
    def upsert_kalem(self, kalem: dict) -> int:
        simdi = datetime.now().isoformat(timespec="seconds")
        created_at = kalem.get("created_at") or simdi
        sefer_id = kalem.get("seferId") or kalem.get("sefer_id")

        def islem(conn):
            if kalem.get("db_id"):
                conn.execute(
                    """
                    UPDATE kalemler
                    SET sefer_id=?, sira=?, belge_no=?, fatura_no=?, fatura_tarihi=?,
                        deger=?, miktar=?, tanim=?, urun_grubu=?, balya_sayisi=?,
                        brut_kg=?, navlun=?, sigorta=?, ilave_yurtdisi=?,
                        excel_secili=?, belge_secili=?,
                        belge_0101_ref=?, belge_0101_tarih=?, belge_0102_ref=?, belge_0102_tarih=?,
                        belge_0979_ref=?, belge_0979_tarih=?, belge_0903_ref=?, belge_0903_tarih=?,
                        belge_0877_ref=?, belge_0877_tarih=?,
                        durum=?, tamamlanma_grup_id=?, created_at=?, updated_at=?
                    WHERE id=?
                    """,
                    (
                        sefer_id,
                        kalem.get("sira", 0),
                        kalem.get("belgeNo", ""),
                        kalem.get("faturaNo", ""),
                        kalem.get("faturaTarihi", ""),
                        kalem.get("deger", ""),
                        kalem.get("miktar", ""),
                        kalem.get("tanim", ""),
                        kalem.get("urunGrubu", ""),
                        kalem.get("balyaSayisi", ""),
                        kalem.get("brutKg", ""),
                        kalem.get("navlun", ""),
                        kalem.get("sigorta", ""),
                        kalem.get("ilaveYurtdisi", ""),
                        1 if kalem.get("excelSecili") else 0,
                        1 if kalem.get("belgeSecili") else 0,
                        kalem.get("belge0101Ref", ""),
                        kalem.get("belge0101Tarih", ""),
                        kalem.get("belge0102Ref", ""),
                        kalem.get("belge0102Tarih", ""),
                        kalem.get("belge0979Ref", ""),
                        kalem.get("belge0979Tarih", ""),
                        kalem.get("belge0903Ref", ""),
                        kalem.get("belge0903Tarih", ""),
                        kalem.get("belge0877Ref", ""),
                        kalem.get("belge0877Tarih", ""),
                        kalem.get("durum", "bekliyor"),
                        kalem.get("tamamlanmaGrupId"),
                        created_at,
                        simdi,
                        kalem["db_id"],
                    ),
                )
                kayit_id = kalem["db_id"]
            else:
                cur = conn.execute(
                    """
                    INSERT INTO kalemler (
                        sefer_id, sira, belge_no, fatura_no, fatura_tarihi, deger, miktar, tanim,
                        urun_grubu, balya_sayisi, brut_kg, navlun, sigorta, ilave_yurtdisi,
                        excel_secili, belge_secili,
                        belge_0101_ref, belge_0101_tarih, belge_0102_ref, belge_0102_tarih,
                        belge_0979_ref, belge_0979_tarih, belge_0903_ref, belge_0903_tarih,
                        belge_0877_ref, belge_0877_tarih, durum, tamamlanma_grup_id,
                        created_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        sefer_id,
                        kalem.get("sira", 0),
                        kalem.get("belgeNo", ""),
                        kalem.get("faturaNo", ""),
                        kalem.get("faturaTarihi", ""),
                        kalem.get("deger", ""),
                        kalem.get("miktar", ""),
                        kalem.get("tanim", ""),
                        kalem.get("urunGrubu", ""),
                        kalem.get("balyaSayisi", ""),
                        kalem.get("brutKg", ""),
                        kalem.get("navlun", ""),
                        kalem.get("sigorta", ""),
                        kalem.get("ilaveYurtdisi", ""),
                        1 if kalem.get("excelSecili") else 0,
                        1 if kalem.get("belgeSecili") else 0,
                        kalem.get("belge0101Ref", ""),
                        kalem.get("belge0101Tarih", ""),
                        kalem.get("belge0102Ref", ""),
                        kalem.get("belge0102Tarih", ""),
                        kalem.get("belge0979Ref", ""),
                        kalem.get("belge0979Tarih", ""),
                        kalem.get("belge0903Ref", ""),
                        kalem.get("belge0903Tarih", ""),
                        kalem.get("belge0877Ref", ""),
                        kalem.get("belge0877Tarih", ""),
                        kalem.get("durum", "bekliyor"),
                        kalem.get("tamamlanmaGrupId"),
                        created_at,
                        simdi,
                    ),
                )
                kayit_id = cur.lastrowid
            conn.commit()
            return kayit_id, created_at, simdi

        kayit_id, created_at, updated_at = self._db_isle(islem)
        kalem["db_id"] = kayit_id
        kalem["created_at"] = created_at
        kalem["updated_at"] = updated_at
        return kayit_id

    def delete_kalem(self, db_id: int) -> bool:
        if not db_id:
            return False

        def islem(conn):
            conn.execute("DELETE FROM kalemler WHERE id=?", (db_id,))
            conn.commit()
            return True

        return self._db_isle(islem)

    def save_kalemler(self, sefer_id: int, kalemler: list) -> bool:
        if sefer_id is None:
            return False
        aktif_idler = []
        for sira, kalem in enumerate(kalemler, start=1):
            kalem["sira"] = sira
            if not kalem.get("seferId"):
                kalem["seferId"] = sefer_id
            aktif_idler.append(self.upsert_kalem(kalem))

        def islem(conn):
            # Sadece bekliyor durumundaki kayıtları temizle; tamamlandi olanları koru
            if aktif_idler:
                guvenceli_idler = [int(x) for x in aktif_idler]
                placeholders = ",".join(["?"] * len(guvenceli_idler))
                sql = (
                    "DELETE FROM kalemler WHERE sefer_id=? AND durum='bekliyor'"
                    " AND id NOT IN (" + placeholders + ")"
                )
                conn.execute(sql, (sefer_id, *guvenceli_idler))
            else:
                conn.execute(
                    "DELETE FROM kalemler WHERE sefer_id=? AND durum='bekliyor'",
                    (sefer_id,),
                )
            conn.commit()
            return True

        return self._db_isle(islem)

    def load_kalemler(self, sefer_id: int, show_completed=False, tamamlanma_grup_no=None) -> list:
        def islem(conn):
            if tamamlanma_grup_no:
                rows = conn.execute(
                    """
                    SELECT k.*, tg.grup_no as tamamlanma_grup_no
                    FROM kalemler k
                    LEFT JOIN tamamlanma_gruplari tg ON k.tamamlanma_grup_id = tg.id
                    WHERE k.sefer_id=? AND tg.grup_no=?
                    ORDER BY k.sira ASC, k.id ASC
                    """,
                    (sefer_id, tamamlanma_grup_no),
                ).fetchall()
            elif show_completed:
                rows = conn.execute(
                    """
                    SELECT k.*, tg.grup_no as tamamlanma_grup_no
                    FROM kalemler k
                    LEFT JOIN tamamlanma_gruplari tg ON k.tamamlanma_grup_id = tg.id
                    WHERE k.sefer_id=?
                    ORDER BY k.sira ASC, k.id ASC
                    """,
                    (sefer_id,),
                ).fetchall()
            else:
                rows = conn.execute(
                    """
                    SELECT k.*, tg.grup_no as tamamlanma_grup_no
                    FROM kalemler k
                    LEFT JOIN tamamlanma_gruplari tg ON k.tamamlanma_grup_id = tg.id
                    WHERE k.sefer_id=? AND k.durum='bekliyor'
                    ORDER BY k.sira ASC, k.id ASC
                    """,
                    (sefer_id,),
                ).fetchall()

            kalemler = []
            for row in rows:
                kalemler.append(
                    {
                        "db_id": row["id"],
                        "seferId": row["sefer_id"],
                        "sira": row["sira"],
                        "belgeNo": row["belge_no"] or "",
                        "faturaNo": row["fatura_no"] or "",
                        "faturaTarihi": row["fatura_tarihi"] or "",
                        "deger": row["deger"] or "",
                        "miktar": row["miktar"] or "",
                        "tanim": row["tanim"] or "",
                        "urunGrubu": row["urun_grubu"] or "",
                        "balyaSayisi": row["balya_sayisi"] or "",
                        "xmlGroupNo": "",
                        "excelGroupNo": "",
                        "durum": row["durum"] or "bekliyor",
                        "brutKg": row["brut_kg"] or "",
                        "navlun": row["navlun"] or "",
                        "sigorta": row["sigorta"] or "",
                        "ilaveYurtdisi": row["ilave_yurtdisi"] or "",
                        "excelSecili": 1 if row["excel_secili"] else 0,
                        "belgeSecili": 1 if row["belge_secili"] else 0,
                        "belge0101Ref": row["belge_0101_ref"] or "",
                        "belge0101Tarih": row["belge_0101_tarih"] or "",
                        "belge0102Ref": row["belge_0102_ref"] or "",
                        "belge0102Tarih": row["belge_0102_tarih"] or "",
                        "belge0979Ref": row["belge_0979_ref"] or "",
                        "belge0979Tarih": row["belge_0979_tarih"] or "",
                        "belge0903Ref": row["belge_0903_ref"] or "",
                        "belge0903Tarih": row["belge_0903_tarih"] or "",
                        "belge0877Ref": row["belge_0877_ref"] or "",
                        "belge0877Tarih": row["belge_0877_tarih"] or "",
                        "tamamlanmaGrupId": row["tamamlanma_grup_id"],
                        "tamamlanmaGrupNo": row["tamamlanma_grup_no"] or "",
                        "created_at": row["created_at"] or "",
                        "updated_at": row["updated_at"] or "",
                    }
                )
            return kalemler

        return self._db_isle(islem)

    def create_or_get_tamamlanma_grubu(self, sefer_id: int, grup_no: str) -> int:
        simdi = datetime.now().isoformat(timespec="seconds")

        def islem(conn):
            row = conn.execute(
                "SELECT id FROM tamamlanma_gruplari WHERE sefer_id=? AND grup_no=?",
                (sefer_id, grup_no),
            ).fetchone()
            if row:
                return row["id"]
            cur = conn.execute(
                "INSERT INTO tamamlanma_gruplari (sefer_id, grup_no, notlar, created_at, updated_at) VALUES (?, ?, '', ?, ?)",
                (sefer_id, grup_no, simdi, simdi),
            )
            conn.commit()
            return cur.lastrowid

        return self._db_isle(islem)

    def complete_kalemler(self, kalem_ids: list, sefer_id: int, grup_no: str):
        grup_id = self.create_or_get_tamamlanma_grubu(sefer_id, grup_no)
        simdi = datetime.now().isoformat(timespec="seconds")

        def islem(conn):
            for kid in kalem_ids:
                conn.execute(
                    "UPDATE kalemler SET durum='tamamlandi', tamamlanma_grup_id=?, updated_at=? WHERE id=?",
                    (grup_id, simdi, kid),
                )
            conn.commit()
            return True

        return self._db_isle(islem)

    def undo_complete_kalemler(self, kalem_ids: list):
        simdi = datetime.now().isoformat(timespec="seconds")

        def islem(conn):
            for kid in kalem_ids:
                conn.execute(
                    "UPDATE kalemler SET durum='bekliyor', tamamlanma_grup_id=NULL, updated_at=? WHERE id=?",
                    (simdi, kid),
                )
            conn.commit()
            return True

        return self._db_isle(islem)

    def list_kalemler_by_tamamlanma(self, sefer_id: int, grup_no: str) -> list:
        return self.load_kalemler(sefer_id, show_completed=True, tamamlanma_grup_no=grup_no)


class BasvuruXMLVeExcel:
    def __init__(self, root):
        self.root = root
        self.root.title("Başvuru XML + Excel Aktar")
        self.root.geometry("1460x920")
        self.root.minsize(1240, 760)

        self.kalemler = []
        self.duzenlenen_index = None
        self._yukleme_modu = False
        self._kaydet_job = None
        self.aktif_sefer_id = None
        self.aktif_tamamlanma_filtre = None
        self.db_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), DB_DOSYA_ADI)
        self.persistence = PersistenceManager(self.db_path)
        self.persistence.init_db()

        self._stil_hazirla()
        self._arayuz_olustur()
        self._otomatik_kayitlari_bagla()
        self.load_all_state()
        self.root.protocol("WM_DELETE_WINDOW", self.pencereyi_kapat)
        self._veritabani_uyarisini_goster()

    # -----------------------------------------------------
    # Stil
    # -----------------------------------------------------
    def _stil_hazirla(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure("TLabel", font=("Segoe UI", 10))
        style.configure("TButton", font=("Segoe UI", 10))
        style.configure("TEntry", font=("Segoe UI", 10))
        style.configure("TCheckbutton", font=("Segoe UI", 10))
        style.configure("TLabelframe.Label", font=("Segoe UI", 10, "bold"))
        style.configure("Baslik.TLabel", font=("Segoe UI", 14, "bold"))
        style.configure("Aciklama.TLabel", font=("Segoe UI", 9))
        style.configure("Ozet.TLabel", font=("Segoe UI", 10, "bold"))
        style.configure("Not.TLabel", font=("Segoe UI", 9))

    # -----------------------------------------------------
    # Arayüz
    # -----------------------------------------------------
    def _arayuz_olustur(self):
        ana = ttk.Frame(self.root, padding=12)
        ana.pack(fill="both", expand=True)

        ttk.Label(ana, text="Başvuru XML + Excel Aktar", style="Baslik.TLabel").pack(anchor="w")
        ttk.Label(
            ana,
            text="Sekme 1 XML ekranını korur, Sekme 2 aynı kalemlerden Excel üretir.",
            style="Aciklama.TLabel"
        ).pack(anchor="w", pady=(4, 10))

        self._sefer_yonetim_bar(ana)

        self.notebook = ttk.Notebook(ana)
        self.notebook.pack(fill="both", expand=True)

        self.tab_xml = ttk.Frame(self.notebook)
        self.tab_excel = ttk.Frame(self.notebook)
        self.tab_belgeler = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_xml, text="XML")
        self.notebook.add(self.tab_excel, text="Excel Aktar")
        self.notebook.add(self.tab_belgeler, text="Belgeler")

        self.sabit_varlar = {
            "gtip": tk.StringVar(value=K_SABIT["gtip"]),
            "belgeTipi": tk.StringVar(value=K_SABIT["belgeTipi"]),
            "menseUlke": tk.StringVar(value=K_SABIT["menseUlke"]),
            "degerBirimi": tk.StringVar(value=K_SABIT["degerBirimi"]),
            "miktarBirimi": tk.StringVar(value=K_SABIT["miktarBirimi"]),
            "urunGrubu": tk.StringVar(value=K_SABIT["urunGrubu"]),
            "grubu": tk.StringVar(value=K_SABIT["grubu"]),
            "urunCinsi": tk.StringVar(value=K_SABIT["urunCinsi"]),
            "imalatciBaslik": tk.StringVar(value=K_SABIT["imalatciBaslik"]),
            "imalatciDeger": tk.StringVar(value=K_SABIT["imalatciDeger"]),
        }

        self._xml_sekmesini_olustur(self.tab_xml)
        self._excel_sekmesini_olustur(self.tab_excel)
        self._belgeler_sekmesini_olustur(self.tab_belgeler)

        self.genel_duzenleme_durumunu_uygula()
        self.sabitler_duzenleme_durumunu_uygula()
        self.belge_fatura_esitleme_durumunu_uygula()
        self.sabitler_gorunumunu_uygula()
        self._ozeti_guncelle()
        self.excel_onizleme_guncelle()

    # -----------------------------------------------------
    # XML SEKMESİ
    # -----------------------------------------------------
    def _sefer_yonetim_bar(self, parent):
        bar = ttk.LabelFrame(parent, text="Sefer Yönetimi", padding=6)
        bar.pack(fill="x", pady=(0, 8))

        sol = ttk.Frame(bar)
        sol.pack(side="left", fill="x", expand=True)

        ttk.Label(sol, text="Aktif Sefer:").pack(side="left")
        self.sefer_combobox_var = tk.StringVar()
        self.sefer_combobox = ttk.Combobox(sol, textvariable=self.sefer_combobox_var, state="readonly", width=30)
        self.sefer_combobox.pack(side="left", padx=(4, 8))
        self.sefer_combobox.bind("<<ComboboxSelected>>", self._sefer_secildi)

        ttk.Button(sol, text="Yeni Sefer Başlat", command=self.yeni_sefer_baslat).pack(side="left", padx=(0, 4))
        ttk.Button(sol, text="Seferi Arşivle", command=self.seferi_arsivle).pack(side="left", padx=(0, 16))

        self.tamamlananlar_goster_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            sol,
            text="Tamamlananları Göster",
            variable=self.tamamlananlar_goster_var,
            command=self._filtre_degisti,
        ).pack(side="left", padx=(0, 12))

        ttk.Label(sol, text="Tamamlanma No:").pack(side="left")
        self.tamamlanma_filtre_var = tk.StringVar()
        ttk.Entry(sol, textvariable=self.tamamlanma_filtre_var, width=8).pack(side="left", padx=(4, 4))
        ttk.Button(sol, text="Göster", command=self._tamamlanma_no_filtrele).pack(side="left", padx=(0, 4))
        ttk.Button(sol, text="Temizle", command=self._filtre_temizle).pack(side="left")

        self.aktif_sefer_bilgi_var = tk.StringVar(value="")
        ttk.Label(bar, textvariable=self.aktif_sefer_bilgi_var, style="Ozet.TLabel").pack(side="right", padx=(8, 0))

    def _xml_sekmesini_olustur(self, parent):
        self._ust_aksiyon_bar(parent)
        self._genel_bilgiler_bolumu(parent)
        self._sabitler_bolumu(parent)
        self._kalem_giris_bolumu(parent)
        self._liste_bolumu(parent)

    def _ust_aksiyon_bar(self, parent):
        bar = ttk.Frame(parent)
        bar.pack(fill="x", pady=(0, 8))

        ttk.Label(bar, text="Taslak Adı").pack(side="left")
        self.taslak_adi_var = tk.StringVar(value=VARSAYILAN_TASLAK_ADI)
        self.taslak_adi_ent = ttk.Entry(bar, textvariable=self.taslak_adi_var, width=24)
        self.taslak_adi_ent.pack(side="left", padx=(6, 12))

        self.genel_duzenle_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            bar,
            text="Genel bilgileri düzenlemeye aç",
            variable=self.genel_duzenle_var,
            command=self.genel_duzenleme_durumunu_uygula
        ).pack(side="left", padx=(0, 12))

        self.sabitleri_goster_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            bar,
            text="Sabitleri göster",
            variable=self.sabitleri_goster_var,
            command=self.sabitler_gorunumunu_uygula
        ).pack(side="left", padx=(0, 12))

        self.sabit_duzenle_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            bar,
            text="Sabitleri düzenlemeye aç",
            variable=self.sabit_duzenle_var,
            command=self.sabitler_duzenleme_durumunu_uygula
        ).pack(side="left")

    def _genel_bilgiler_bolumu(self, parent):
        lf = ttk.LabelFrame(parent, text="Genel Bilgiler", padding=10)
        lf.pack(fill="x", pady=(0, 8))

        self.genel_varlar = {
            "gondericiVergiNo": tk.StringVar(value=VARSAYILAN_GENEL["gondericiVergiNo"]),
            "gondericiUnvan": tk.StringVar(value=VARSAYILAN_GENEL["gondericiUnvan"]),
            "sevkUlkesi": tk.StringVar(value=VARSAYILAN_GENEL["sevkUlkesi"]),
            "ticaretYapilanUlke": tk.StringVar(value=VARSAYILAN_GENEL["ticaretYapilanUlke"]),
            "cikisUlkesi": tk.StringVar(value=VARSAYILAN_GENEL["cikisUlkesi"]),
            "sinirdakiTasimaSekli": tk.StringVar(value=VARSAYILAN_GENEL["sinirdakiTasimaSekli"]),
            "esyaninBulunduguYer": tk.StringVar(value=VARSAYILAN_GENEL["esyaninBulunduguYer"]),
            "girisGumruk": tk.StringVar(value=VARSAYILAN_GENEL["girisGumruk"]),
            "bosaltmaGumruk": tk.StringVar(value=VARSAYILAN_GENEL["bosaltmaGumruk"]),
        }

        self.genel_entryler = {}

        alanlar = [
            ("Gönderici Vergi No", "gondericiVergiNo"),
            ("Gönderici Ünvan", "gondericiUnvan"),
            ("Sevk Ülkesi", "sevkUlkesi"),
            ("Ticaret Yapılan Ülke", "ticaretYapilanUlke"),
            ("Çıkış Ülkesi", "cikisUlkesi"),
            ("Sınırdaki Taşıma Şekli", "sinirdakiTasimaSekli"),
            ("Eşyanın Bulunduğu Yer", "esyaninBulunduguYer"),
            ("Giriş Gümrük", "girisGumruk"),
            ("Boşaltma Gümrük", "bosaltmaGumruk"),
        ]

        for i, (etiket, key) in enumerate(alanlar):
            r = i // 3
            c = (i % 3) * 2
            ttk.Label(lf, text=etiket).grid(row=r, column=c, sticky="w", padx=(0, 6), pady=4)
            ent = ttk.Entry(lf, textvariable=self.genel_varlar[key], width=26)
            ent.grid(row=r, column=c + 1, sticky="we", padx=(0, 14), pady=4)
            self.genel_entryler[key] = ent

        for col in range(6):
            lf.columnconfigure(col, weight=1)

    def _sabitler_bolumu(self, parent):
        self.sabit_lf = ttk.LabelFrame(parent, text="Sabit Gelen Bilgiler", padding=10)
        self.sabit_entryler = {}

        alanlar = [
            ("GTİP", "gtip"),
            ("Belge Tipi", "belgeTipi"),
            ("Menşe Ülke", "menseUlke"),
            ("Değer Birimi", "degerBirimi"),
            ("Miktar Birimi", "miktarBirimi"),
            ("Ürün Grubu", "urunGrubu"),
            ("Grubu", "grubu"),
            ("Ürün Cinsi", "urunCinsi"),
            ("İmalatçı Başlığı", "imalatciBaslik"),
            ("İmalatçı Değeri", "imalatciDeger"),
        ]

        ttk.Label(self.sabit_lf, text="Belge Tarihi XML oluşturulduğu gün otomatik yazılır.").grid(
            row=0, column=0, columnspan=6, sticky="w", pady=(0, 8)
        )

        for i, (etiket, key) in enumerate(alanlar, start=1):
            r = ((i - 1) // 3) + 1
            c = ((i - 1) % 3) * 2
            ttk.Label(self.sabit_lf, text=etiket).grid(row=r, column=c, sticky="w", padx=(0, 6), pady=4)
            ent = ttk.Entry(self.sabit_lf, textvariable=self.sabit_varlar[key], width=28)
            ent.grid(row=r, column=c + 1, sticky="we", padx=(0, 14), pady=4)
            self.sabit_entryler[key] = ent

        for col in range(6):
            self.sabit_lf.columnconfigure(col, weight=1)

    def _kalem_giris_bolumu(self, parent):
        orta = ttk.Frame(parent)
        self.kalem_giris_container = orta
        orta.pack(fill="x", pady=(0, 8))

        form = ttk.LabelFrame(orta, text="Kalem Girişi", padding=10)
        form.pack(side="left", fill="both", expand=True, padx=(0, 8))

        self.belge_no_var = tk.StringVar()
        self.fatura_no_var = tk.StringVar()
        self.fatura_tarihi_var = tk.StringVar(value=bugunun_tarihi_gosterim())
        self.deger_var = tk.StringVar()
        self.miktar_var = tk.StringVar()
        self.balya_sayisi_var = tk.StringVar()
        self.tanim_var = tk.StringVar(value="BALYA PAMUK")

        self.belge_fatura_esit_var = tk.BooleanVar(value=True)
        self.no_otomatik_artsin_var = tk.BooleanVar(value=True)
        self.tarih_sabit_kalsin_var = tk.BooleanVar(value=True)

        self.belge_no_var.trace_add("write", self.belge_no_degisti)
        self.balya_sayisi_var.trace_add("write", self.balya_sayisi_degisti)

        alanlar = [
            ("Belge No", self.belge_no_var),
            ("Fatura No", self.fatura_no_var),
            ("Fatura Tarihi (GG.AA.YYYY)", self.fatura_tarihi_var),
            ("Değer", self.deger_var),
            ("Miktar", self.miktar_var),
            ("Balya Sayısı", self.balya_sayisi_var),
            ("Kaplar ve Eşyanın Tanımı", self.tanim_var),
        ]

        self.form_entryleri = {}
        for i, (etiket, var) in enumerate(alanlar):
            r = i // 2
            c = (i % 2) * 2
            ttk.Label(form, text=etiket).grid(row=r, column=c, sticky="w", padx=(0, 6), pady=5)
            width = 42 if etiket == "Kaplar ve Eşyanın Tanımı" else 28
            ent = ttk.Entry(form, textvariable=var, width=width)
            ent.grid(row=r, column=c + 1, sticky="we", padx=(0, 14), pady=5)
            ent.bind("<Return>", self.enter_ile_kaydet)
            self.form_entryleri[etiket] = ent

        ttk.Checkbutton(
            form,
            text="Belge No ile Fatura No aynı olsun",
            variable=self.belge_fatura_esit_var,
            command=self.belge_fatura_esitleme_durumunu_uygula
        ).grid(row=4, column=0, columnspan=2, sticky="w", pady=(4, 2))

        ttk.Checkbutton(
            form,
            text="Belge / Fatura No otomatik artsın",
            variable=self.no_otomatik_artsin_var
        ).grid(row=4, column=2, sticky="w", pady=(4, 2))

        ttk.Checkbutton(
            form,
            text="Fatura tarihi aynı kalsın",
            variable=self.tarih_sabit_kalsin_var
        ).grid(row=4, column=3, sticky="w", pady=(4, 2))

        ttk.Label(
            form,
            text="Balya Sayısı yazınca tanım otomatik '106 BALYA PAMUK' gibi oluşur.",
            style="Not.TLabel"
        ).grid(row=5, column=0, columnspan=4, sticky="w", pady=(4, 2))

        buton = ttk.Frame(form)
        buton.grid(row=6, column=0, columnspan=4, sticky="w", pady=(10, 0))

        self.btn_ekle = ttk.Button(buton, text="Kalem Ekle", command=self.kalem_ekle_veya_guncelle)
        self.btn_ekle.pack(side="left")

        ttk.Button(buton, text="Formu Temizle", command=self.kalem_alanlarini_temizle).pack(side="left", padx=(8, 0))
        ttk.Button(buton, text="Seçiliyi Forma Yükle", command=self.seciliyi_forma_yukle).pack(side="left", padx=(8, 0))

        for col in range(4):
            form.columnconfigure(col, weight=1)

        ozet = ttk.LabelFrame(orta, text="Özet", padding=10)
        ozet.pack(side="left", fill="y")

        self.ozet_kalem_var = tk.StringVar(value="Toplam Kalem: 0")
        self.ozet_deger_var = tk.StringVar(value="Toplam Değer: 0")
        self.ozet_miktar_var = tk.StringVar(value="Toplam Miktar: 0")
        self.duzenleme_modu_var = tk.StringVar(value="Mod: Yeni kayıt")

        ttk.Label(ozet, textvariable=self.ozet_kalem_var, style="Ozet.TLabel").pack(anchor="w", pady=3)
        ttk.Label(ozet, textvariable=self.ozet_deger_var, style="Ozet.TLabel").pack(anchor="w", pady=3)
        ttk.Label(ozet, textvariable=self.ozet_miktar_var, style="Ozet.TLabel").pack(anchor="w", pady=3)
        ttk.Separator(ozet, orient="horizontal").pack(fill="x", pady=8)
        ttk.Label(ozet, textvariable=self.duzenleme_modu_var).pack(anchor="w", pady=3)

    def _liste_bolumu(self, parent):
        lf = ttk.LabelFrame(parent, text="Eklenen Kalemler", padding=10)
        lf.pack(fill="both", expand=True)

        kolonlar = ("no", "belgeNo", "faturaTarihi", "faturaNo", "deger", "miktar", "urunGrubu", "balya", "tanim", "durum")
        self.tree = ttk.Treeview(lf, columns=kolonlar, show="headings", height=15)

        basliklar = {
            "no": "#",
            "belgeNo": "Belge No",
            "faturaTarihi": "Fatura Tarihi",
            "faturaNo": "Fatura No",
            "deger": "Değer",
            "miktar": "Miktar",
            "urunGrubu": "Ürün Grubu",
            "balya": "Balya Sayısı",
            "tanim": "Tanım",
            "durum": "Durum",
        }
        genislikler = {
            "no": 50,
            "belgeNo": 170,
            "faturaTarihi": 110,
            "faturaNo": 170,
            "deger": 110,
            "miktar": 110,
            "urunGrubu": 170,
            "balya": 90,
            "tanim": 220,
            "durum": 100,
        }

        for col in kolonlar:
            self.tree.heading(col, text=basliklar[col])
            anchor = "center" if col in ("no", "faturaTarihi", "balya", "durum") else "w"
            if col in ("deger", "miktar"):
                anchor = "e"
            self.tree.column(col, width=genislikler[col], anchor=anchor)

        self.tree.tag_configure("tamamlandi", foreground="#888888")
        self.tree.bind("<Double-1>", lambda event: self.seciliyi_forma_yukle())

        ysb = ttk.Scrollbar(lf, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=ysb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")

        buton = ttk.Frame(lf)
        buton.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(10, 0))

        ttk.Button(buton, text="Seçiliyi Forma Yükle", command=self.seciliyi_forma_yukle).pack(side="left")
        ttk.Button(buton, text="Seçiliyi Kopyala", command=self.secili_kalemi_kopyala).pack(side="left", padx=(8, 0))
        ttk.Button(buton, text="Seçiliyi Sil", command=self.secili_kalemi_sil).pack(side="left", padx=(8, 0))
        ttk.Button(buton, text="Tümünü Temizle", command=self.tum_kalemleri_temizle).pack(side="left", padx=(8, 0))

        ttk.Separator(buton, orient="vertical").pack(side="left", fill="y", padx=(8, 8))
        ttk.Button(buton, text="Seçilileri Tamamlandı Yap", command=self.secilileri_tamamlandi_yap).pack(side="left", padx=(0, 4))
        ttk.Button(buton, text="Tamamlandıyı Geri Al", command=self.tamamlandiyi_geri_al).pack(side="left")

        ttk.Button(buton, text="XML Önizleme", command=self.xml_onizleme).pack(side="right")
        ttk.Button(buton, text="XML Kaydet", command=self.xml_kaydet).pack(side="right", padx=(0, 8))

        self.durum_var = tk.StringVar(value="Hazır. Veriler otomatik kaydediliyor.")
        ttk.Label(lf, textvariable=self.durum_var).grid(row=2, column=0, columnspan=2, sticky="w", pady=(8, 0))

        lf.columnconfigure(0, weight=1)
        lf.rowconfigure(0, weight=1)

    # -----------------------------------------------------
    # EXCEL SEKMESİ
    # -----------------------------------------------------
    def _excel_sekmesini_olustur(self, parent):
        ana = ttk.Frame(parent, padding=10)
        ana.pack(fill="both", expand=True)

        # Eski Excel ayar alanları kaldırıldı; sadece satır bazlı elle giriş var.
        self.excel_varlar = {}
        self.excel_duzenlenen_index = None

        ttk.Label(
            ana,
            text="XML sekmesindeki kalemlerden Excel satırı üretilir.",
            style="Not.TLabel"
        ).pack(anchor="w", pady=(0, 2))
        ttk.Label(
            ana,
            text="Kullanıcı sadece Brüt KG, Navlun, Sigorta ve İlave Yurtdışı alanlarını girer.",
            style="Not.TLabel"
        ).pack(anchor="w", pady=(0, 6))

        orta = ttk.LabelFrame(ana, text="Excel'e Gidecek Satırlar", padding=10)
        orta.pack(fill="both", expand=True, pady=(0, 8))

        self.excel_tree = ttk.Treeview(
            orta,
            columns=("sec", "id", "belgeNo", "faturaNo", "tarih", "tutar", "miktar", "balya", "brutKg", "navlun", "sigorta", "ilaveYurtdisi"),
            show="headings",
            height=14,
            selectmode="browse"
        )

        basliklar = {
            "sec": "Seç",
            "id": "ID",
            "belgeNo": "Belge No",
            "faturaNo": "Fatura No",
            "tarih": "Fatura Tarihi",
            "tutar": "Fatura Tutar",
            "miktar": "Satış Miktarı",
            "balya": "Balya Sayısı",
            "brutKg": "Brüt KG",
            "navlun": "Navlun",
            "sigorta": "Sigorta",
            "ilaveYurtdisi": "İlave Yurtdışı",
        }
        genislikler = {
            "sec": 52,
            "id": 55,
            "belgeNo": 140,
            "faturaNo": 140,
            "tarih": 115,
            "tutar": 110,
            "miktar": 110,
            "balya": 100,
            "brutKg": 95,
            "navlun": 95,
            "sigorta": 95,
            "ilaveYurtdisi": 110,
        }

        for col in basliklar:
            self.excel_tree.heading(col, text=basliklar[col])
            anchor = "center" if col in ("sec", "id", "tarih") else "w"
            if col in ("tutar", "miktar", "balya", "brutKg", "navlun", "sigorta", "ilaveYurtdisi"):
                anchor = "e"
            self.excel_tree.column(col, width=genislikler[col], anchor=anchor)

        self.excel_tree.bind("<<TreeviewSelect>>", lambda event: self.excel_seciliyi_duzenleme_paneline_yukle())
        self.excel_tree.bind("<Button-1>", self.excel_tree_tiklandi)

        ysb = ttk.Scrollbar(orta, orient="vertical", command=self.excel_tree.yview)
        self.excel_tree.configure(yscroll=ysb.set)
        self.excel_tree.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        orta.columnconfigure(0, weight=1)
        orta.rowconfigure(0, weight=1)

        form = ttk.LabelFrame(ana, text="Seçili Satır İçin Elle Giriş", padding=10)
        form.pack(fill="x", pady=(0, 8))

        self.excel_secili_satir_var = tk.StringVar(value="Seçili satır: yok")
        ttk.Label(form, textvariable=self.excel_secili_satir_var, style="Ozet.TLabel").grid(
            row=0, column=0, columnspan=4, sticky="w", pady=(0, 8)
        )

        self.excel_duzenle_varlar = {
            "brutKg": tk.StringVar(),
            "navlun": tk.StringVar(),
            "sigorta": tk.StringVar(),
            "ilaveYurtdisi": tk.StringVar(),
        }

        alanlar = [
            ("Brüt KG", "brutKg"),
            ("Navlun", "navlun"),
            ("Sigorta", "sigorta"),
            ("İlave Yurtdışı", "ilaveYurtdisi"),
        ]

        for i, (etiket, key) in enumerate(alanlar, start=1):
            c = ((i - 1) % 4) * 2
            ttk.Label(form, text=etiket).grid(row=1, column=c, sticky="w", padx=(0, 6), pady=4)
            ent = ttk.Entry(form, textvariable=self.excel_duzenle_varlar[key], width=16)
            ent.grid(row=1, column=c + 1, sticky="we", padx=(0, 12), pady=4)
            ent.bind("<Return>", self.excel_secili_satira_uygula)

        ttk.Button(form, text="Seçili Satıra Uygula", command=self.excel_secili_satira_uygula).grid(
            row=2, column=0, columnspan=8, sticky="w", pady=(8, 0)
        )

        for col in range(8):
            form.columnconfigure(col, weight=1)

        alt = ttk.Frame(ana)
        alt.pack(fill="x")

        self.excel_ozet_var = tk.StringVar(value="Excel için hazır satır yok.")
        ttk.Label(alt, textvariable=self.excel_ozet_var, style="Ozet.TLabel").pack(side="left")

        ttk.Button(alt, text="İşaretleri Temizle", command=self.excel_isaretleri_temizle).pack(side="right")
        ttk.Button(alt, text="Tümünü İşaretle", command=self.excel_tumunu_isaretle).pack(side="right", padx=(0, 8))
        ttk.Button(alt, text="Seçili Satırı İşaretle/Kaldır", command=self.excel_secili_satiri_isaretle_kaldir).pack(side="right", padx=(0, 8))
        ttk.Button(alt, text="Listeyi Güncelle", command=self.excel_onizleme_guncelle).pack(side="right", padx=(0, 8))
        ttk.Button(
            alt,
            text="İşaretlileri Excel Oluştur",
            command=lambda: self.excel_olustur(yalnizca_secilen=True)
        ).pack(side="right", padx=(0, 8))
        ttk.Button(alt, text="Tümünü Excel Oluştur (.xlsx)", command=self.excel_olustur).pack(side="right", padx=(0, 8))

    def _belgeler_sekmesini_olustur(self, parent):
        ana = ttk.Frame(parent, padding=10)
        ana.pack(fill="both", expand=True)

        ttk.Label(
            ana,
            text="Her kalem için belge ref ve tarih bilgilerini burada tutabilir, ayrı BELGELER Excel'i üretebilirsin.",
            style="Not.TLabel"
        ).pack(anchor="w", pady=(0, 6))

        ust = ttk.LabelFrame(ana, text="Kalemler", padding=10)
        ust.pack(fill="both", expand=True, pady=(0, 8))

        kolonlar = ("sec", "id", "belgeNo", "faturaNo", "faturaTarihi", "deger", "miktar")
        self.belgeler_tree = ttk.Treeview(ust, columns=kolonlar, show="headings", height=10, selectmode="browse")
        basliklar = {
            "sec": "Seç",
            "id": "ID",
            "belgeNo": "Belge No",
            "faturaNo": "Fatura No",
            "faturaTarihi": "Fatura Tarihi",
            "deger": "Değer",
            "miktar": "Miktar",
        }
        genislikler = {
            "sec": 52,
            "id": 55,
            "belgeNo": 170,
            "faturaNo": 170,
            "faturaTarihi": 110,
            "deger": 120,
            "miktar": 120,
        }

        for col in kolonlar:
            self.belgeler_tree.heading(col, text=basliklar[col])
            anchor = "center" if col in ("sec", "id", "faturaTarihi") else "w"
            if col in ("deger", "miktar"):
                anchor = "e"
            self.belgeler_tree.column(col, width=genislikler[col], anchor=anchor)

        self.belgeler_tree.bind("<<TreeviewSelect>>", lambda event: self.belgeler_seciliyi_forma_yukle())
        self.belgeler_tree.bind("<Button-1>", self.belgeler_tree_tiklandi)

        ysb = ttk.Scrollbar(ust, orient="vertical", command=self.belgeler_tree.yview)
        self.belgeler_tree.configure(yscroll=ysb.set)
        self.belgeler_tree.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        ust.columnconfigure(0, weight=1)
        ust.rowconfigure(0, weight=1)

        alt = ttk.LabelFrame(ana, text="Belge Giriş Paneli", padding=10)
        alt.pack(fill="x")

        self.belge_secili_satir_var = tk.StringVar(value="Seçili kalem: yok")
        ttk.Label(alt, textvariable=self.belge_secili_satir_var, style="Ozet.TLabel").grid(
            row=0, column=0, columnspan=5, sticky="w", pady=(0, 8)
        )

        self.belge_alan_varlari = {
            "faturaNo": tk.StringVar(),
            "faturaTarihi": tk.StringVar(),
            "belge0101Ref": tk.StringVar(),
            "belge0101Tarih": tk.StringVar(),
            "belge0102Ref": tk.StringVar(),
            "belge0102Tarih": tk.StringVar(),
            "belge0979Ref": tk.StringVar(),
            "belge0979Tarih": tk.StringVar(),
            "belge0903Ref": tk.StringVar(),
            "belge0903Tarih": tk.StringVar(),
            "belge0877Ref": tk.StringVar(),
            "belge0877Tarih": tk.StringVar(),
        }

        # Yeni: tek belge tipi seçimi
        belge_options = [f"{kod} - {baslik}" for kod, baslik, *_ in BELGE_KOD_TANIMLARI]
        self.selected_belge_code_var = tk.StringVar(value=BELGE_KOD_TANIMLARI[0][0])
        ttk.Label(alt, text="Belge Tipi").grid(row=1, column=0, sticky="w", padx=(0, 6), pady=4)
        self.belge_tipi_cb = ttk.Combobox(alt, values=belge_options, state="readonly", width=28)
        # combobox display uses '0101 - Navlun Fatura' but store only kod in var
        self.belge_tipi_cb.current(0)
        self.belge_tipi_cb.grid(row=1, column=1, columnspan=2, sticky="we", padx=(0, 10), pady=4)
        # vars for the single ref/tarih inputs
        self.selected_belge_ref_var = tk.StringVar()
        self.selected_belge_tarih_var = tk.StringVar()

        ttk.Label(alt, text="Ref").grid(row=2, column=0, sticky="w")
        self.selected_belge_ref_entry = ttk.Entry(alt, textvariable=self.selected_belge_ref_var, width=32)
        self.selected_belge_ref_entry.grid(row=2, column=1, sticky="we", padx=(0, 10), pady=4)

        ttk.Label(alt, text="Tarih (GG.AA.YYYY)").grid(row=2, column=2, sticky="w")
        self.selected_belge_tarih_entry = ttk.Entry(alt, textvariable=self.selected_belge_tarih_var, width=18)
        self.selected_belge_tarih_entry.grid(row=2, column=3, sticky="we", padx=(0, 12), pady=4)

        # When combobox changes, update which fields are editable
        def _on_belge_tipi_change(event=None):
            sel = self.belge_tipi_cb.get().split(" - ")[0]
            readonly = BELGE_READONLY.get(sel, False)
            if readonly:
                self.selected_belge_ref_entry.configure(state="readonly")
                self.selected_belge_tarih_entry.configure(state="readonly")
            else:
                self.selected_belge_ref_entry.configure(state="normal")
                self.selected_belge_tarih_entry.configure(state="normal")
            # yüklemeyi güncelle (seçili kalemde gösterilecek alan değişti)
            try:
                self.belgeler_seciliyi_forma_yukle()
            except Exception:
                pass

        self.belge_tipi_cb.bind("<<ComboboxSelected>>", _on_belge_tipi_change)
        _on_belge_tipi_change()

        self.belge_boslari_koru_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            alt,
            text="Boş bırakılan alanlar mevcut veriyi silmesin",
            variable=self.belge_boslari_koru_var
        ).grid(row=len(BELGE_KOD_TANIMLARI) + 1, column=0, columnspan=5, sticky="w", pady=(8, 0))
        buton_bar = ttk.Frame(alt)
        buton_bar.grid(row=4, column=0, columnspan=5, sticky="w", pady=(10, 0))
        ttk.Button(buton_bar, text="Seçili Kalemi İşaretle/Kaldır", command=self.belge_secili_satiri_isaretle_kaldir).pack(side="left")
        ttk.Button(buton_bar, text="Tümünü İşaretle", command=self.belgeler_tumunu_isaretle).pack(side="left", padx=(8, 0))
        ttk.Button(buton_bar, text="İşaretleri Temizle", command=self.belgeler_isaretleri_temizle).pack(side="left", padx=(8, 0))
        ttk.Button(buton_bar, text="Seçili Kaleme Uygula", command=self.belge_bilgilerini_kaydet).pack(side="left", padx=(8, 0))
        ttk.Button(buton_bar, text="İşaretli Kalemlere Uygula", command=self.belge_isaretli_kalemlere_uygula).pack(side="left", padx=(8, 0))
        ttk.Button(buton_bar, text="Formu Temizle", command=self.belge_formu_temizle).pack(side="left", padx=(8, 0))
        ttk.Button(buton_bar, text="Belgeler Excel Oluştur", command=self.belgeler_excel_olustur).pack(side="left", padx=(8, 0))

        for col in range(5):
            alt.columnconfigure(col, weight=1)

        self.belgeler_ozet_var = tk.StringVar(value="İşaretli kalem: 0 / 0")
        ttk.Label(ana, textvariable=self.belgeler_ozet_var, style="Ozet.TLabel").pack(anchor="w", pady=(6, 0))

    # -----------------------------------------------------
    # Kalıcı hafıza / state yönetimi
    # -----------------------------------------------------
    def _ayar_verisini_topla(self):
        veri = {"taslakAdi": self.taslak_adi_var.get().strip()}

        for key, var in self.genel_varlar.items():
            veri[key] = var.get().strip()
        for key, var in self.sabit_varlar.items():
            veri[key] = var.get().strip()
        for key, var in self.excel_varlar.items():
            veri[key] = var.get().strip()

        veri.update(
            {
                "genel_duzenle_var": self.genel_duzenle_var.get(),
                "sabitleri_goster_var": self.sabitleri_goster_var.get(),
                "sabit_duzenle_var": self.sabit_duzenle_var.get(),
                "belge_fatura_esit_var": self.belge_fatura_esit_var.get(),
                "no_otomatik_artsin_var": self.no_otomatik_artsin_var.get(),
                "tarih_sabit_kalsin_var": self.tarih_sabit_kalsin_var.get(),
            }
        )
        return veri

    def _bool_ayari_coz(self, value, default=False):
        if value is None:
            return default
        return str(value).strip().lower() in {"1", "true", "evet", "yes", "on"}

    def _otomatik_kayitlari_bagla(self):
        izlenecek_varlar = [
            self.taslak_adi_var,
            *self.genel_varlar.values(),
            *self.sabit_varlar.values(),
            *self.excel_varlar.values(),
            self.genel_duzenle_var,
            self.sabitleri_goster_var,
            self.sabit_duzenle_var,
            self.belge_fatura_esit_var,
            self.no_otomatik_artsin_var,
            self.tarih_sabit_kalsin_var,
        ]

        for var in izlenecek_varlar:
            var.trace_add("write", self._ayar_kaydet_zamanla)

    def _ayar_kaydet_zamanla(self, *args):
        if self._yukleme_modu:
            return
        if self._kaydet_job is not None:
            self.root.after_cancel(self._kaydet_job)
        self._kaydet_job = self.root.after(350, self.save_settings)

    def _veritabani_uyarisini_goster(self):
        uyari = self.persistence.bekleyen_uyariyi_al()
        if uyari:
            messagebox.showwarning("Veritabanı Uyarısı", uyari)

    def _kayit_bilgisi_yaz(self, mesaj):
        saat = datetime.now().strftime("%H:%M:%S")
        self.durum_var.set(f"{mesaj} | Veriler otomatik kaydediliyor | Son kayıt: {saat}")

    def save_settings(self):
        self._kaydet_job = None
        try:
            self.persistence.save_settings(self._ayar_verisini_topla())
            self._veritabani_uyarisini_goster()
        except Exception as exc:
            messagebox.showerror("Kayıt Hatası", f"Ayarlar kaydedilemedi:\n{exc}")

    def load_settings(self):
        try:
            settings = self.persistence.load_settings()
        except Exception as exc:
            messagebox.showerror("Yükleme Hatası", f"Ayarlar yüklenemedi:\n{exc}")
            return {}

        self._yukleme_modu = True
        try:
            if "taslakAdi" in settings:
                self.taslak_adi_var.set(settings["taslakAdi"])

            for key, var in self.genel_varlar.items():
                if key in settings:
                    var.set(settings[key])
            for key, var in self.sabit_varlar.items():
                if key in settings:
                    var.set(settings[key])
            for key, var in self.excel_varlar.items():
                if key in settings:
                    var.set(settings[key])

            self.genel_duzenle_var.set(self._bool_ayari_coz(settings.get("genel_duzenle_var"), False))
            self.sabitleri_goster_var.set(self._bool_ayari_coz(settings.get("sabitleri_goster_var"), False))
            self.sabit_duzenle_var.set(self._bool_ayari_coz(settings.get("sabit_duzenle_var"), False))
            self.belge_fatura_esit_var.set(self._bool_ayari_coz(settings.get("belge_fatura_esit_var"), True))
            self.no_otomatik_artsin_var.set(self._bool_ayari_coz(settings.get("no_otomatik_artsin_var"), True))
            self.tarih_sabit_kalsin_var.set(self._bool_ayari_coz(settings.get("tarih_sabit_kalsin_var"), True))
        finally:
            self._yukleme_modu = False

        self.genel_duzenleme_durumunu_uygula()
        self.sabitler_duzenleme_durumunu_uygula()
        self.belge_fatura_esitleme_durumunu_uygula()
        self.sabitler_gorunumunu_uygula()
        return settings

    def save_records(self):
        try:
            self.persistence.save_kalemler(self.aktif_sefer_id, self.kalemler)
            self._veritabani_uyarisini_goster()
        except Exception as exc:
            messagebox.showerror("Kayıt Hatası", f"Kalemler kaydedilemedi:\n{exc}")
            raise

    def load_records(self):
        self._kalemler_yukle()
        return self.kalemler

    def upsert_record(self, kalem):
        try:
            kayit_id = self.persistence.upsert_kalem(kalem)
            self._veritabani_uyarisini_goster()
            return kayit_id
        except Exception as exc:
            messagebox.showerror("Kayıt Hatası", f"Kalem veritabanına yazılamadı:\n{exc}")
            raise

    def delete_record(self, db_id):
        try:
            sonuc = self.persistence.delete_kalem(db_id)
            self._veritabani_uyarisini_goster()
            return sonuc
        except Exception as exc:
            messagebox.showerror("Silme Hatası", f"Kalem silinemedi:\n{exc}")
            raise

    def save_all_state(self):
        if self._kaydet_job is not None:
            self.root.after_cancel(self._kaydet_job)
            self._kaydet_job = None
        try:
            if hasattr(self, "excel_duzenle_varlar"):
                self.excel_secili_satira_uygula(sessiz=True)
            if hasattr(self, "belge_alan_varlari"):
                self.belge_bilgilerini_kaydet(sessiz=True)
            self.persistence.save_all_state(self._ayar_verisini_topla(), self.kalemler)
            self._veritabani_uyarisini_goster()
        except Exception as exc:
            messagebox.showerror("Kayıt Hatası", f"Uygulama verileri kaydedilemedi:\n{exc}")
            raise

    def load_all_state(self):
        self.load_settings()

        # Sefer durumunu başlat
        try:
            aktif_id = self.persistence.get_active_sefer_id()
            if aktif_id:
                self.aktif_sefer_id = aktif_id
            else:
                seferler = self.persistence.list_seferler(tumu=True)
                if seferler:
                    self.aktif_sefer_id = seferler[0]["id"]
                    self.persistence.set_active_sefer_id(self.aktif_sefer_id)
                else:
                    sefer_id = self.persistence.create_sefer("Varsayılan Sefer")
                    self.aktif_sefer_id = sefer_id
                    self.persistence.set_active_sefer_id(sefer_id)
        except Exception as exc:
            messagebox.showerror("Sefer Hatası", f"Sefer durumu yüklenemedi:\n{exc}")

        self._sefer_combobox_guncelle()
        self._aktif_sefer_bilgisini_guncelle()
        self._kalemler_yukle()

        if self.kalemler:
            self._kayit_bilgisi_yaz(f"{len(self.kalemler)} kayıt yüklendi")
        else:
            self.durum_var.set("Hazır. Veriler otomatik kaydediliyor.")

    def pencereyi_kapat(self):
        try:
            self.save_all_state()
        except Exception:
            if not messagebox.askyesno("Kapatılsın mı?", "Son kayıt sırasında hata oluştu. Yine de uygulama kapatılsın mı?"):
                return
        self.root.destroy()

    # -----------------------------------------------------
    # Belgeler sekmesi yardımcıları
    # -----------------------------------------------------
    def _belge_isaret_simgesi(self, kalem):
        return "☑" if kalem.get("belgeSecili") else "☐"

    def belgeler_isaretli_kayitlari_getir(self):
        return [kalem for kalem in self.kalemler if kalem.get("belgeSecili")]

    def belgeler_tree_tiklandi(self, event):
        bolge = self.belgeler_tree.identify("region", event.x, event.y)
        kolon = self.belgeler_tree.identify_column(event.x)
        item = self.belgeler_tree.identify_row(event.y)

        if bolge == "cell" and kolon == "#1" and item:
            index = self.belgeler_tree.index(item)
            if 0 <= index < len(self.kalemler):
                self.belgeler_tree.selection_set(item)
                self.belge_satir_isaretini_degistir(index)
            return "break"
        return None

    def belge_satir_isaretini_degistir(self, index):
        if not (0 <= index < len(self.kalemler)):
            return
        kalem = self.kalemler[index]
        kalem["belgeSecili"] = 0 if kalem.get("belgeSecili") else 1
        self.upsert_record(kalem)
        self.belgeler_listesini_guncelle(secili_index=index)

    def belge_secili_satiri_isaretle_kaldir(self):
        secim = self.belgeler_tree.selection()
        if not secim:
            messagebox.showwarning("Seçim Yok", "Önce Belgeler tablosundan bir kalem seç.")
            return
        index = self.belgeler_tree.index(secim[0])
        self.belge_satir_isaretini_degistir(index)

    def belgeler_tumunu_isaretle(self):
        if not self.kalemler:
            return
        for kalem in self.kalemler:
            kalem["belgeSecili"] = 1
        self.save_records()
        secili_index = None
        secim = self.belgeler_tree.selection()
        if secim:
            secili_index = self.belgeler_tree.index(secim[0])
        self.belgeler_listesini_guncelle(secili_index=secili_index)
        self._kayit_bilgisi_yaz("Tüm kalemler belge uygulaması için işaretlendi")

    def belgeler_isaretleri_temizle(self):
        if not self.kalemler:
            return
        for kalem in self.kalemler:
            kalem["belgeSecili"] = 0
        self.save_records()
        secili_index = None
        secim = self.belgeler_tree.selection()
        if secim:
            secili_index = self.belgeler_tree.index(secim[0])
        self.belgeler_listesini_guncelle(secili_index=secili_index)
        self._kayit_bilgisi_yaz("Belge işaretleri temizlendi")

    def belgeler_listesini_guncelle(self, secili_index=None):
        if not hasattr(self, "belgeler_tree"):
            return

        if secili_index is None:
            secim = self.belgeler_tree.selection()
            if secim:
                secili_index = self.belgeler_tree.index(secim[0])

        for item in self.belgeler_tree.get_children():
            self.belgeler_tree.delete(item)

        for kalem in self.kalemler:
            self.belgeler_tree.insert(
                "",
                "end",
                values=(
                    self._belge_isaret_simgesi(kalem),
                    kalem.get("sira", ""),
                    kalem.get("belgeNo", ""),
                    kalem.get("faturaNo", ""),
                    kalem.get("faturaTarihi", ""),
                    kalem.get("deger", ""),
                    kalem.get("miktar", ""),
                )
            )

        if hasattr(self, "belgeler_ozet_var"):
            isaretli_sayi = len(self.belgeler_isaretli_kayitlari_getir())
            self.belgeler_ozet_var.set(f"İşaretli kalem: {isaretli_sayi} / {len(self.kalemler)}")

        cocuklar = self.belgeler_tree.get_children()
        if secili_index is not None and 0 <= secili_index < len(cocuklar):
            item = cocuklar[secili_index]
            self.belgeler_tree.selection_set(item)
            self.belgeler_tree.focus(item)
        self.belgeler_seciliyi_forma_yukle()

    def belgeler_seciliyi_forma_yukle(self):
        if not hasattr(self, "belge_alan_varlari"):
            return

        secim = self.belgeler_tree.selection()
        if not secim:
            self.belge_secili_satir_var.set("Seçili kalem: yok")
            # temizle sadece formdaki tek alanları
            self.selected_belge_ref_var.set("")
            self.selected_belge_tarih_var.set("")
            return

        index = self.belgeler_tree.index(secim[0])
        if not (0 <= index < len(self.kalemler)):
            return

        kalem = self.kalemler[index]
        self.belge_secili_satir_var.set(f"Seçili kalem: {kalem['sira']} | Belge No: {kalem['belgeNo']}")
        # fatura alanlarını de koru (DB alanı olarak mevcut)
        if "faturaNo" in self.belge_alan_varlari:
            self.belge_alan_varlari["faturaNo"].set(kalem.get("faturaNo", ""))
            self.belge_alan_varlari["faturaTarihi"].set(kalem.get("faturaTarihi", ""))

        # combobox'tan seçili belge koduna göre ilgili alanları formda göster
        sel = self.belge_tipi_cb.get().split(" - ")[0]
        if sel in BELGE_ALAN_HARITASI:
            ref_key, tarih_key = BELGE_ALAN_HARITASI[sel]
            self.selected_belge_ref_var.set(kalem.get(ref_key, ""))
            self.selected_belge_tarih_var.set(kalem.get(tarih_key, ""))
        else:
            self.selected_belge_ref_var.set("")
            self.selected_belge_tarih_var.set("")

    def _belge_form_degerlerini_hazirla(self):
        # Yeni mantık: combobox'tan seçilen belge koduna göre sadece o alanları hazırla
        sel = self.belge_tipi_cb.get().split(" - ")[0]
        if sel not in BELGE_ALAN_HARITASI:
            messagebox.showerror("Seçim Hatası", "Geçersiz belge tipi seçildi.")
            return None
        ref_key, tarih_key = BELGE_ALAN_HARITASI[sel]
        ref = self.selected_belge_ref_var.get().strip()
        tarih = self.selected_belge_tarih_var.get().strip()
        if tarih and not self.tarih_gecerli_mi(tarih):
            messagebox.showerror("Tarih Hatası", f"{sel} tarihi GG.AA.YYYY formatında olmalı.")
            return None
        return {ref_key: ref, tarih_key: tarih}

    def _belge_degerlerini_hedeflere_uygula(self, hedef_kalemler, secili_index=None, sessiz=False, toplu=False):
        form_degerleri = self._belge_form_degerlerini_hazirla()
        if form_degerleri is None:
            return False

        boslari_koru = self.belge_boslari_koru_var.get() if hasattr(self, "belge_boslari_koru_var") else True
        if boslari_koru and not any(form_degerleri.values()):
            if not sessiz:
                messagebox.showwarning("Bilgi Yok", "Uygulanacak en az bir belge alanı doldur.")
            return False

        degisen_sayisi = 0
        for kalem in hedef_kalemler:
            degisti = False
            for alan, yeni_deger in form_degerleri.items():
                # sadece seçilen belge alanlarına dokun
                if boslari_koru and yeni_deger == "":
                    continue
                if kalem.get(alan, "") != yeni_deger:
                    kalem[alan] = yeni_deger
                    degisti = True
            if degisti:
                self.upsert_record(kalem)
                degisen_sayisi += 1

        self.belgeler_listesini_guncelle(secili_index=secili_index)

        if not sessiz:
            if toplu:
                if degisen_sayisi:
                    self._kayit_bilgisi_yaz(f"{degisen_sayisi} işaretli kaleme belge bilgileri uygulandı")
                else:
                    self._kayit_bilgisi_yaz("İşaretli kalemlerde değişiklik olmadı")
            else:
                kalem = hedef_kalemler[0]
                if degisen_sayisi:
                    self._kayit_bilgisi_yaz(f"{kalem['sira']}. kalem belge bilgileri kaydedildi")
                else:
                    self._kayit_bilgisi_yaz(f"{kalem['sira']}. kalemde değişiklik olmadı")
        return True

    def belge_bilgilerini_kaydet(self, sessiz=False):
        if not hasattr(self, "belgeler_tree"):
            return False

        secim = self.belgeler_tree.selection()
        if not secim:
            if not sessiz:
                messagebox.showwarning("Seçim Yok", "Önce Belgeler sekmesinden bir kalem seç.")
            return False

        index = self.belgeler_tree.index(secim[0])
        if not (0 <= index < len(self.kalemler)):
            return False

        return self._belge_degerlerini_hedeflere_uygula(
            [self.kalemler[index]],
            secili_index=index,
            sessiz=sessiz,
            toplu=False,
        )

    def belge_isaretli_kalemlere_uygula(self):
        if not hasattr(self, "belgeler_tree"):
            return False

        hedef_kalemler = self.belgeler_isaretli_kayitlari_getir()
        if not hedef_kalemler:
            messagebox.showwarning("İşaret Yok", "Önce en az bir kalemi işaretle.")
            return False

        secili_index = None
        secim = self.belgeler_tree.selection()
        if secim:
            secili_index = self.belgeler_tree.index(secim[0])

        return self._belge_degerlerini_hedeflere_uygula(
            hedef_kalemler,
            secili_index=secili_index,
            toplu=True,
        )

    def belge_formu_temizle(self, reset_belge_tipi=False):
        # Ref ve Tarih alanlarını temizle; belge tipi isteğe bağlı korunur
        self.selected_belge_ref_var.set("")
        self.selected_belge_tarih_var.set("")
        if reset_belge_tipi:
            # combobox'u varsayılan ilk öğeye döndür
            try:
                self.belge_tipi_cb.current(0)
            except Exception:
                pass
        self._kayit_bilgisi_yaz("Belge formu temizlendi")

    def belgeler_satirlarini_uret(self, kayitlar=None):
        satirlar = []
        kaynak_kayitlar = self.kalemler if kayitlar is None else kayitlar

        for kalem in kaynak_kayitlar:
            for kod, _, ref_key, tarih_key, readonly in BELGE_KOD_TANIMLARI:
                if readonly:
                    ref = kalem.get("faturaNo", "")
                    tarih = kalem.get("faturaTarihi", "")
                    var_yok = "V"
                else:
                    ref = kalem.get(ref_key, "")
                    tarih = kalem.get(tarih_key, "")
                    var_yok = "V" if ref else ""

                satirlar.append(
                    {
                        "Kalem No": kalem.get("sira", ""),
                        "Kod": kod,
                        "Açıklama": "",
                        "Var/Yok": var_yok,
                        "Belge Ref": ref,
                        "Belge Tarih": tarihi_belge_excel_formatina_cevir(tarih),
                        "Vize Tarih": "",
                    }
                )
        return satirlar

    def belgeler_excel_olustur(self):
        if Workbook is None:
            messagebox.showerror("Kütüphane Yok", "Excel üretmek için openpyxl gerekli.")
            return

        self.belge_bilgilerini_kaydet(sessiz=True)
        kayitlar = self.belgeler_isaretli_kayitlari_getir()
        if not kayitlar:
            messagebox.showwarning("İşaret Yok", "Önce en az bir kalemi işaretle.")
            return

        satirlar = self.belgeler_satirlarini_uret(kayitlar=kayitlar)
        if not satirlar:
            messagebox.showwarning("Kalem Yok", "Önce XML sekmesinde en az 1 kalem ekle.")
            return

        wb = Workbook()
        ws = wb.active
        ws.title = "BELGELER"
        basliklar = ["Kalem No", "Kod", "Açıklama", "Var/Yok", "Belge Ref", "Belge Tarih", "Vize Tarih"]

        for col_idx, header in enumerate(basliklar, start=1):
            ws.cell(row=1, column=col_idx, value=header)

        for row_idx, row_data in enumerate(satirlar, start=2):
            for col_idx, header in enumerate(basliklar, start=1):
                ws.cell(row=row_idx, column=col_idx, value=row_data.get(header, ""))

        from openpyxl.utils import get_column_letter
        for col_idx, header in enumerate(basliklar, start=1):
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max(len(header) + 2, 14), 26)

        dosya_adi = (self.taslak_adi_var.get().strip() or "belgeler") + "_belgeler.xlsx"
        yol = filedialog.asksaveasfilename(
            title="Belgeler Excel Kaydet",
            defaultextension=".xlsx",
            initialfile=dosya_adi,
            filetypes=[("Excel Dosyası", "*.xlsx"), ("Tüm Dosyalar", "*.*")]
        )
        if not yol:
            return

        try:
            wb.save(yol)
            messagebox.showinfo("Başarılı", f"Belgeler Excel oluşturuldu.\n\n{yol}")
        except Exception as e:
            messagebox.showerror("Hata", f"Belgeler Excel kaydedilirken hata oluştu:\n{e}")

    # -----------------------------------------------------
    # Ortak yardımcılar
    # -----------------------------------------------------
    def sabitler_gorunumunu_uygula(self):
        if self.sabitleri_goster_var.get():
            if not self.sabit_lf.winfo_ismapped():
                self.sabit_lf.pack(fill="x", pady=(0, 8), before=self.kalem_giris_container)
        else:
            if self.sabit_lf.winfo_ismapped():
                self.sabit_lf.pack_forget()

    def sabitler_duzenleme_durumunu_uygula(self):
        durum = "normal" if self.sabit_duzenle_var.get() else "readonly"
        for ent in self.sabit_entryler.values():
            ent.configure(state=durum)

    def genel_duzenleme_durumunu_uygula(self):
        durum = "normal" if self.genel_duzenle_var.get() else "readonly"
        for ent in self.genel_entryler.values():
            ent.configure(state=durum)
        self.taslak_adi_ent.configure(state=durum)

    def belge_fatura_esitleme_durumunu_uygula(self):
        if self.belge_fatura_esit_var.get():
            self.form_entryleri["Fatura No"].configure(state="readonly")
            self.fatura_no_var.set(self.belge_no_var.get())
        else:
            self.form_entryleri["Fatura No"].configure(state="normal")

    def belge_no_degisti(self, *args):
        if self.belge_fatura_esit_var.get():
            self.fatura_no_var.set(self.belge_no_var.get())

    def balya_sayisi_degisti(self, *args):
        deger = self.balya_sayisi_var.get().strip()
        if deger:
            self.tanim_var.set(f"{deger} BALYA PAMUK")
        else:
            self.tanim_var.set("BALYA PAMUK")

    def enter_ile_kaydet(self, event=None):
        self.kalem_ekle_veya_guncelle()
        return "break"

    def sayi_duzelt(self, metin: str) -> str:
        metin = (metin or "").strip().replace(" ", "")
        if not metin:
            return ""

        if "," in metin and "." in metin:
            metin = metin.replace(".", "").replace(",", ".")
        else:
            metin = metin.replace(",", ".")

        return metin

    def sayi_gecerli_mi(self, metin: str) -> bool:
        try:
            float(self.sayi_duzelt(metin))
            return True
        except Exception:
            return False

    def tarih_gecerli_mi(self, metin: str) -> bool:
        parca = (metin or "").strip().split(".")
        if len(parca) != 3:
            return False
        try:
            gun, ay, yil = int(parca[0]), int(parca[1]), int(parca[2])
            _ = date(yil, ay, gun)
            return True
        except Exception:
            return False

    def _sayiya_cevir(self, metin: str) -> float:
        try:
            return float(self.sayi_duzelt(metin))
        except Exception:
            return 0.0

    def _ozeti_guncelle(self):
        toplam_deger = sum(self._sayiya_cevir(k["deger"]) for k in self.kalemler)
        toplam_miktar = sum(self._sayiya_cevir(k["miktar"]) for k in self.kalemler)
        self.ozet_kalem_var.set(f"Toplam Kalem: {len(self.kalemler)}")
        self.ozet_deger_var.set(f"Toplam Değer: {toplam_deger:,.2f}")
        self.ozet_miktar_var.set(f"Toplam Miktar: {toplam_miktar:,.2f}")

    def _kalem_treeye_ekle(self, kalem):
        durum_goster = ""
        if kalem.get("durum") == "tamamlandi":
            grup_no = kalem.get("tamamlanmaGrupNo", "")
            durum_goster = f"✓ {grup_no}" if grup_no else "✓"
        tag = "tamamlandi" if kalem.get("durum") == "tamamlandi" else ""
        self.tree.insert(
            "",
            "end",
            values=(
                kalem["sira"],
                kalem["belgeNo"],
                kalem["faturaTarihi"],
                kalem["faturaNo"],
                kalem["deger"],
                kalem["miktar"],
                kalem["urunGrubu"],
                kalem["balyaSayisi"],
                kalem["tanim"],
                durum_goster,
            ),
            tags=(tag,) if tag else (),
        )

    def _kalem_dict_olustur(
        self,
        belge_no,
        fatura_tarihi,
        fatura_no,
        deger,
        miktar,
        tanim,
        urun_grubu,
        balya_sayisi,
        sira,
        db_id=None,
        xml_group_no="",
        excel_group_no="",
        durum="bekliyor",
        brut_kg="",
        navlun="",
        sigorta="",
        ilave_yurtdisi="",
        excel_secili=0,
        belge_secili=0,
        belge_0101_ref="",
        belge_0101_tarih="",
        belge_0102_ref="",
        belge_0102_tarih="",
        belge_0979_ref="",
        belge_0979_tarih="",
        belge_0903_ref="",
        belge_0903_tarih="",
        belge_0877_ref="",
        belge_0877_tarih="",
        created_at="",
        updated_at="",
        sefer_id=None,
        tamamlanma_grup_id=None,
        tamamlanma_grup_no="",
    ):
        return {
            "db_id": db_id,
            "seferId": sefer_id,
            "sira": sira,
            "belgeNo": belge_no.strip(),
            "faturaTarihi": fatura_tarihi.strip(),
            "faturaNo": fatura_no.strip(),
            "deger": self.sayi_duzelt(deger),
            "miktar": self.sayi_duzelt(miktar),
            "tanim": tanim.strip(),
            "urunGrubu": urun_grubu.strip(),
            "balyaSayisi": self.sayi_duzelt(balya_sayisi),
            "xmlGroupNo": xml_group_no or "",
            "excelGroupNo": excel_group_no or "",
            "durum": durum or "bekliyor",
            "brutKg": self.sayi_duzelt(brut_kg) if brut_kg else "",
            "navlun": self.sayi_duzelt(navlun) if navlun else "",
            "sigorta": self.sayi_duzelt(sigorta) if sigorta else "",
            "ilaveYurtdisi": self.sayi_duzelt(ilave_yurtdisi) if ilave_yurtdisi else "",
            "excelSecili": 1 if excel_secili else 0,
            "belgeSecili": 1 if belge_secili else 0,
            "belge0101Ref": belge_0101_ref.strip(),
            "belge0101Tarih": belge_0101_tarih.strip(),
            "belge0102Ref": belge_0102_ref.strip(),
            "belge0102Tarih": belge_0102_tarih.strip(),
            "belge0979Ref": belge_0979_ref.strip(),
            "belge0979Tarih": belge_0979_tarih.strip(),
            "belge0903Ref": belge_0903_ref.strip(),
            "belge0903Tarih": belge_0903_tarih.strip(),
            "belge0877Ref": belge_0877_ref.strip(),
            "belge0877Tarih": belge_0877_tarih.strip(),
            "tamamlanmaGrupId": tamamlanma_grup_id,
            "tamamlanmaGrupNo": tamamlanma_grup_no or "",
            "created_at": created_at or "",
            "updated_at": updated_at or "",
        }

    def kalem_alanlarini_temizle(self, sonraki_belge_no="", sonraki_fatura_no="", sonraki_tarih=None, odak_alan="Belge No"):
        self.duzenlenen_index = None
        self.duzenleme_modu_var.set("Mod: Yeni kayıt")
        self.btn_ekle.configure(text="Kalem Ekle")

        self.belge_no_var.set(sonraki_belge_no)
        if sonraki_tarih is None:
            self.fatura_tarihi_var.set(bugunun_tarihi_gosterim())
        else:
            self.fatura_tarihi_var.set(sonraki_tarih)

        if self.belge_fatura_esit_var.get():
            self.fatura_no_var.set(sonraki_belge_no if sonraki_belge_no else "")
        else:
            self.fatura_no_var.set(sonraki_fatura_no)

        self.deger_var.set("")
        self.miktar_var.set("")
        self.balya_sayisi_var.set("")
        self.tanim_var.set("BALYA PAMUK")

        hedef = self.form_entryleri.get(odak_alan, self.form_entryleri["Belge No"])
        hedef.focus_set()
        try:
            hedef.selection_range(0, "end")
        except Exception:
            pass

    def _girisleri_al_ve_kontrol_et(self):
        belge_no = self.belge_no_var.get().strip()
        fatura_tarihi = self.fatura_tarihi_var.get().strip()
        fatura_no = self.fatura_no_var.get().strip()
        deger = self.deger_var.get().strip()
        miktar = self.miktar_var.get().strip()
        urun_grubu = self.sabit_varlar["urunGrubu"].get().strip()
        balya_sayisi = self.balya_sayisi_var.get().strip()
        tanim = self.tanim_var.get().strip()

        boslar = []
        if not belge_no:
            boslar.append("Belge No")
        if not fatura_tarihi:
            boslar.append("Fatura Tarihi")
        if not fatura_no:
            boslar.append("Fatura No")
        if not deger:
            boslar.append("Değer")
        if not miktar:
            boslar.append("Miktar")
        if not urun_grubu:
            boslar.append("Ürün Grubu")
        if not balya_sayisi:
            boslar.append("Balya Sayısı")
        if not tanim:
            boslar.append("Kaplar ve Eşyanın Tanımı")

        if boslar:
            messagebox.showwarning("Eksik Bilgi", "Şu alanlar boş bırakılamaz:\n- " + "\n- ".join(boslar))
            return None

        if not self.tarih_gecerli_mi(fatura_tarihi):
            messagebox.showerror("Tarih Hatası", "Fatura Tarihi GG.AA.YYYY formatında olmalı.\nÖrnek: 03.04.2026")
            return None

        if not self.sayi_gecerli_mi(deger) or not self.sayi_gecerli_mi(miktar):
            messagebox.showerror("Sayı Hatası", "Değer ve Miktar geçerli sayı olmalı.")
            return None

        if not self.sayi_gecerli_mi(balya_sayisi):
            messagebox.showerror("Sayı Hatası", "Balya Sayısı geçerli sayı olmalı.")
            return None

        return belge_no, fatura_tarihi, fatura_no, deger, miktar, tanim, urun_grubu, balya_sayisi

    def kalem_ekle_veya_guncelle(self):
        veri = self._girisleri_al_ve_kontrol_et()
        if veri is None:
            return

        belge_no, fatura_tarihi, fatura_no, deger, miktar, tanim, urun_grubu, balya_sayisi = veri

        duzenleme_modu = self.duzenlenen_index is not None

        if not duzenleme_modu:
            kalem = self._kalem_dict_olustur(
                belge_no, fatura_tarihi, fatura_no, deger, miktar, tanim, urun_grubu, balya_sayisi,
                len(self.kalemler) + 1,
                sefer_id=self.aktif_sefer_id,
            )
            self.upsert_record(kalem)
            self.kalemler.append(kalem)
            self._kalem_treeye_ekle(kalem)
            self._kayit_bilgisi_yaz(f"{len(self.kalemler)} adet kalem eklendi")
        else:
            eski_kalem = self.kalemler[self.duzenlenen_index]
            eski_sira = eski_kalem["sira"]
            kalem = self._kalem_dict_olustur(
                belge_no,
                fatura_tarihi,
                fatura_no,
                deger,
                miktar,
                tanim,
                urun_grubu,
                balya_sayisi,
                eski_sira,
                db_id=eski_kalem.get("db_id"),
                xml_group_no=eski_kalem.get("xmlGroupNo", ""),
                excel_group_no=eski_kalem.get("excelGroupNo", ""),
                durum=eski_kalem.get("durum", "bekliyor"),
                brut_kg=eski_kalem.get("brutKg", ""),
                navlun=eski_kalem.get("navlun", ""),
                sigorta=eski_kalem.get("sigorta", ""),
                ilave_yurtdisi=eski_kalem.get("ilaveYurtdisi", ""),
                excel_secili=eski_kalem.get("excelSecili", 0),
                belge_secili=eski_kalem.get("belgeSecili", 0),
                belge_0101_ref=eski_kalem.get("belge0101Ref", ""),
                belge_0101_tarih=eski_kalem.get("belge0101Tarih", ""),
                belge_0102_ref=eski_kalem.get("belge0102Ref", ""),
                belge_0102_tarih=eski_kalem.get("belge0102Tarih", ""),
                belge_0979_ref=eski_kalem.get("belge0979Ref", ""),
                belge_0979_tarih=eski_kalem.get("belge0979Tarih", ""),
                belge_0903_ref=eski_kalem.get("belge0903Ref", ""),
                belge_0903_tarih=eski_kalem.get("belge0903Tarih", ""),
                belge_0877_ref=eski_kalem.get("belge0877Ref", ""),
                belge_0877_tarih=eski_kalem.get("belge0877Tarih", ""),
                created_at=eski_kalem.get("created_at", ""),
                updated_at=eski_kalem.get("updated_at", ""),
                sefer_id=eski_kalem.get("seferId") or self.aktif_sefer_id,
                tamamlanma_grup_id=eski_kalem.get("tamamlanmaGrupId"),
                tamamlanma_grup_no=eski_kalem.get("tamamlanmaGrupNo", ""),
            )
            self.upsert_record(kalem)
            self.kalemler[self.duzenlenen_index] = kalem
            self.kalem_numaralarini_yenile()
            self._kayit_bilgisi_yaz(f"{eski_sira}. kalem güncellendi")

        self._ozeti_guncelle()
        self.excel_onizleme_guncelle()
        self.belgeler_listesini_guncelle()

        if not duzenleme_modu:
            sonraki_belge_no = belge_numarasi_artir(belge_no) if self.no_otomatik_artsin_var.get() else ""
            if self.belge_fatura_esit_var.get():
                sonraki_fatura_no = sonraki_belge_no
            else:
                sonraki_fatura_no = belge_numarasi_artir(fatura_no) if self.no_otomatik_artsin_var.get() else ""
            sonraki_tarih = fatura_tarihi if self.tarih_sabit_kalsin_var.get() else bugunun_tarihi_gosterim()
            odak = "Değer" if self.no_otomatik_artsin_var.get() else "Belge No"
            self.kalem_alanlarini_temizle(sonraki_belge_no, sonraki_fatura_no, sonraki_tarih, odak)
        else:
            self.kalem_alanlarini_temizle()

    def seciliyi_forma_yukle(self):
        secim = self.tree.selection()
        if not secim:
            messagebox.showwarning("Seçim Yok", "Önce listeden bir kalem seç.")
            return

        item = secim[0]
        index = self.tree.index(item)
        if not (0 <= index < len(self.kalemler)):
            return

        kalem = self.kalemler[index]
        self.duzenlenen_index = index
        self.duzenleme_modu_var.set(f"Mod: Düzenleme ({kalem['sira']}. kalem)")
        self.btn_ekle.configure(text="Seçili Kalemi Güncelle")

        self.belge_no_var.set(kalem["belgeNo"])
        self.fatura_tarihi_var.set(kalem["faturaTarihi"])
        if self.belge_fatura_esit_var.get():
            self.fatura_no_var.set(kalem["belgeNo"])
        else:
            self.fatura_no_var.set(kalem["faturaNo"])
        self.deger_var.set(kalem["deger"])
        self.miktar_var.set(kalem["miktar"])
        self.balya_sayisi_var.set(kalem["balyaSayisi"])
        self.tanim_var.set(kalem["tanim"])
        self.form_entryleri["Belge No"].focus_set()
        try:
            self.form_entryleri["Belge No"].selection_range(0, "end")
        except Exception:
            pass

    def secili_kalemi_kopyala(self):
        secim = self.tree.selection()
        if not secim:
            messagebox.showwarning("Seçim Yok", "Önce listeden bir kalem seç.")
            return

        item = secim[0]
        index = self.tree.index(item)
        if not (0 <= index < len(self.kalemler)):
            return

        yeni = copy.deepcopy(self.kalemler[index])
        yeni["db_id"] = None
        yeni["created_at"] = ""
        yeni["updated_at"] = ""
        yeni["excelSecili"] = 0
        yeni["belgeSecili"] = 0
        yeni["sira"] = len(self.kalemler) + 1
        yeni["seferId"] = self.aktif_sefer_id
        yeni["durum"] = "bekliyor"
        yeni["tamamlanmaGrupId"] = None
        yeni["tamamlanmaGrupNo"] = ""
        self.upsert_record(yeni)
        self.kalemler.append(yeni)
        self._kalem_treeye_ekle(yeni)
        self._ozeti_guncelle()
        self.excel_onizleme_guncelle()
        self.belgeler_listesini_guncelle()
        self._kayit_bilgisi_yaz(f"{index + 1}. kalem kopyalandı. Yeni sıra: {yeni['sira']}")

    def secili_kalemi_sil(self):
        secim = self.tree.selection()
        if not secim:
            messagebox.showwarning("Seçim Yok", "Önce listeden bir kalem seç.")
            return

        indeksler = sorted({self.tree.index(item) for item in secim}, reverse=True)
        for index in indeksler:
            if 0 <= index < len(self.kalemler):
                kalem = self.kalemler[index]
                self.delete_record(kalem.get("db_id"))
                self.kalemler.pop(index)

        self.kalem_numaralarini_yenile()
        self.save_records()
        self._ozeti_guncelle()
        self.excel_onizleme_guncelle()
        self.belgeler_listesini_guncelle()
        self.kalem_alanlarini_temizle()
        self._kayit_bilgisi_yaz(f"Kalem silindi. Toplam kalem: {len(self.kalemler)}")

    def tum_kalemleri_temizle(self):
        if not self.kalemler:
            return
        if not messagebox.askyesno("Onay", "Tüm kalemler silinsin mi?"):
            return

        self.kalemler.clear()
        self.save_records()
        for item in self.tree.get_children():
            self.tree.delete(item)
        self._ozeti_guncelle()
        self.excel_onizleme_guncelle()
        self.belgeler_listesini_guncelle()
        self.kalem_alanlarini_temizle()
        self._kayit_bilgisi_yaz("Tüm kalemler temizlendi")

    def kalem_numaralarini_yenile(self):
        for i, kalem in enumerate(self.kalemler, start=1):
            kalem["sira"] = i

        for item in self.tree.get_children():
            self.tree.delete(item)

        for kalem in self.kalemler:
            self._kalem_treeye_ekle(kalem)

    # -----------------------------------------------------
    # XML
    # -----------------------------------------------------
    def _alt_etiket(self, parent, etiket, deger):
        sub = ET.SubElement(parent, etiket)
        sub.text = deger if deger is not None else ""

    def xml_agaci_olustur(self):
        if not self.kalemler:
            messagebox.showwarning("Kalem Yok", "XML oluşturmak için en az 1 kalem eklemelisin.")
            return None

        root = ET.Element("basvuru")
        self._alt_etiket(root, "taslakAdi", self.taslak_adi_var.get().strip())

        genel = ET.SubElement(root, "genelBilgiler")
        self._alt_etiket(genel, "gondericiVergiNo", self.genel_varlar["gondericiVergiNo"].get().strip())
        self._alt_etiket(genel, "gondericiUnvan", self.genel_varlar["gondericiUnvan"].get().strip())
        self._alt_etiket(genel, "sevkUlkesi", self.genel_varlar["sevkUlkesi"].get().strip())
        self._alt_etiket(genel, "ticaretYapilanUlke", self.genel_varlar["ticaretYapilanUlke"].get().strip())
        self._alt_etiket(genel, "cikisUlkesi", self.genel_varlar["cikisUlkesi"].get().strip())
        self._alt_etiket(genel, "sinirdakiTasimaSekli", self.genel_varlar["sinirdakiTasimaSekli"].get().strip())
        self._alt_etiket(genel, "esyaninBulunduguYer", self.genel_varlar["esyaninBulunduguYer"].get().strip())
        self._alt_etiket(genel, "girisGumruk", self.genel_varlar["girisGumruk"].get().strip())
        self._alt_etiket(genel, "bosaltmaGumruk", self.genel_varlar["bosaltmaGumruk"].get().strip())

        for kalem in self.kalemler:
            node = ET.SubElement(root, "kalemler")
            self._alt_etiket(node, "gtip", self.sabit_varlar["gtip"].get().strip())
            self._alt_etiket(node, "belgeTipi", self.sabit_varlar["belgeTipi"].get().strip())
            self._alt_etiket(node, "belgeTarihi", bugunun_tarihi_xml())
            self._alt_etiket(node, "belgeNo", kalem["belgeNo"])
            self._alt_etiket(node, "satirNo", "")
            self._alt_etiket(node, "tasimaSenediNo", "")
            self._alt_etiket(node, "konteynerNo", "")
            self._alt_etiket(node, "kaplarVeEsyaninTanimi", kalem["tanim"])
            self._alt_etiket(node, "menseUlke", self.sabit_varlar["menseUlke"].get().strip())
            self._alt_etiket(node, "faturaTarihi", tarihi_xml_formatina_cevir(kalem["faturaTarihi"]))
            self._alt_etiket(node, "faturaNo", kalem["faturaNo"])
            self._alt_etiket(node, "deger", kalem["deger"])
            self._alt_etiket(node, "degerBirimi", self.sabit_varlar["degerBirimi"].get().strip())
            self._alt_etiket(node, "miktar", kalem["miktar"])
            self._alt_etiket(node, "miktarBirimi", self.sabit_varlar["miktarBirimi"].get().strip())
            self._alt_etiket(node, "kapsamDisi", "false")
            ET.SubElement(node, "muafiyet")
            ET.SubElement(node, "geriGelenEsyaKayitNo")
            ET.SubElement(node, "cevreUyumIzniBelgeNo")
            ET.SubElement(node, "tseMuafiyetBelgeNo")
            self._alt_etiket(node, "urunGrubu", self.sabit_varlar["urunGrubu"].get().strip())

            dyn = ET.SubElement(node, "dinamikler")
            self._alt_etiket(dyn, "baslik", "BALYA SAYISI")
            self._alt_etiket(dyn, "dinamikDeger", kalem["balyaSayisi"])

            dyn = ET.SubElement(node, "dinamikler")
            self._alt_etiket(dyn, "baslik", "GRUBU")
            self._alt_etiket(dyn, "dinamikDeger", self.sabit_varlar["grubu"].get().strip())

            dyn = ET.SubElement(node, "dinamikler")
            self._alt_etiket(dyn, "baslik", "ÜRÜN CİNSİ")
            self._alt_etiket(dyn, "dinamikDeger", self.sabit_varlar["urunCinsi"].get().strip())

            dyn = ET.SubElement(node, "dinamikler")
            self._alt_etiket(dyn, "baslik", self.sabit_varlar["imalatciBaslik"].get().strip())
            self._alt_etiket(dyn, "dinamikDeger", self.sabit_varlar["imalatciDeger"].get().strip())

        return root

    def xml_string_uret(self):
        root = self.xml_agaci_olustur()
        if root is None:
            return None

        ham = ET.tostring(root, encoding="utf-8")
        pretty = minidom.parseString(ham).toprettyxml(indent="    ", encoding="UTF-8")
        return pretty.decode("utf-8")

    def xml_onizleme(self):
        xml_text = self.xml_string_uret()
        if not xml_text:
            return

        pencere = tk.Toplevel(self.root)
        pencere.title("XML Önizleme")
        pencere.geometry("980x720")

        txt = tk.Text(pencere, wrap="none", font=("Consolas", 10))
        txt.pack(fill="both", expand=True)
        txt.insert("1.0", xml_text)
        txt.configure(state="disabled")

    def xml_kaydet(self):
        xml_text = self.xml_string_uret()
        if not xml_text:
            return

        dosya_adi = (self.taslak_adi_var.get().strip() or "basvuru") + ".xml"
        yol = filedialog.asksaveasfilename(
            title="XML Kaydet",
            defaultextension=".xml",
            initialfile=dosya_adi,
            filetypes=[("XML Dosyası", "*.xml"), ("Tüm Dosyalar", "*.*")]
        )
        if not yol:
            return

        try:
            with open(yol, "w", encoding="utf-8", newline="") as f:
                f.write(xml_text)
            self.durum_var.set(f"XML kaydedildi: {yol}")
            messagebox.showinfo("Başarılı", f"XML başarıyla kaydedildi.\n\n{yol}")
        except Exception as e:
            messagebox.showerror("Hata", f"Kaydetme sırasında hata oluştu:\n{e}")

    # -----------------------------------------------------
    # EXCEL
    # -----------------------------------------------------
    def _excel_sabit_degeri(self, kolon):
        deger = EXCEL_SABIT_DOLU.get(kolon, "")
        if isinstance(deger, tuple) and len(deger) == 2 and deger[0] == "sabit_var":
            return self.sabit_varlar[deger[1]].get().strip()
        return deger

    def _excel_sayi_veya_bos(self, metin: str):
        metin = (metin or "").strip()
        if not metin:
            return ""
        return self._sayiya_cevir(metin)

    def excel_metni_sayi_formatla(self, metin: str) -> str:
        """Excel hücresine sayı yerine noktalı metin yazar."""
        metin = self.sayi_duzelt("" if metin is None else str(metin))
        return metin

    def _excel_isaret_simgesi(self, kalem):
        return "☑" if kalem.get("excelSecili") else "☐"

    def excel_isaretli_kayitlari_getir(self):
        return [kalem for kalem in self.kalemler if kalem.get("excelSecili")]

    def excel_tree_tiklandi(self, event):
        bolge = self.excel_tree.identify("region", event.x, event.y)
        kolon = self.excel_tree.identify_column(event.x)
        item = self.excel_tree.identify_row(event.y)

        if bolge == "cell" and kolon == "#1" and item:
            index = self.excel_tree.index(item)
            if 0 <= index < len(self.kalemler):
                self.excel_tree.selection_set(item)
                self.excel_satir_isaretini_degistir(index)
            return "break"
        return None

    def excel_satir_isaretini_degistir(self, index):
        if not (0 <= index < len(self.kalemler)):
            return
        kalem = self.kalemler[index]
        kalem["excelSecili"] = 0 if kalem.get("excelSecili") else 1
        self.upsert_record(kalem)
        self.excel_onizleme_guncelle(secili_index=index)

    def excel_secili_satiri_isaretle_kaldir(self):
        secim = self.excel_tree.selection()
        if not secim:
            messagebox.showwarning("Seçim Yok", "Önce Excel tablosundan bir satır seç.")
            return
        index = self.excel_tree.index(secim[0])
        self.excel_satir_isaretini_degistir(index)

    def excel_tumunu_isaretle(self):
        if not self.kalemler:
            return
        for kalem in self.kalemler:
            kalem["excelSecili"] = 1
        self.save_records()
        self.excel_onizleme_guncelle(secili_index=self.excel_duzenlenen_index)
        self._kayit_bilgisi_yaz("Tüm kayıtlar Excel için işaretlendi")

    def excel_isaretleri_temizle(self):
        if not self.kalemler:
            return
        for kalem in self.kalemler:
            kalem["excelSecili"] = 0
        self.save_records()
        self.excel_onizleme_guncelle(secili_index=self.excel_duzenlenen_index)
        self._kayit_bilgisi_yaz("Excel işaretleri temizlendi")

    def excel_seciliyi_duzenleme_paneline_yukle(self):
        secim = self.excel_tree.selection()
        if not secim:
            self.excel_duzenlenen_index = None
            self.excel_secili_satir_var.set("Seçili satır: yok")
            for var in self.excel_duzenle_varlar.values():
                var.set("")
            return

        index = self.excel_tree.index(secim[0])
        if not (0 <= index < len(self.kalemler)):
            return

        kalem = self.kalemler[index]
        self.excel_duzenlenen_index = index
        self.excel_secili_satir_var.set(f"Seçili satır: {kalem['sira']} | Belge No: {kalem['belgeNo']}")
        self.excel_duzenle_varlar["brutKg"].set(kalem.get("brutKg", ""))
        self.excel_duzenle_varlar["navlun"].set(kalem.get("navlun", ""))
        self.excel_duzenle_varlar["sigorta"].set(kalem.get("sigorta", ""))
        self.excel_duzenle_varlar["ilaveYurtdisi"].set(kalem.get("ilaveYurtdisi", ""))

    def excel_secili_satira_uygula(self, event=None, sessiz=False):
        if self.excel_duzenlenen_index is None:
            self.excel_seciliyi_duzenleme_paneline_yukle()

        if self.excel_duzenlenen_index is None:
            if not sessiz:
                messagebox.showwarning("Seçim Yok", "Önce Excel tablosundan bir satır seç.")
            return "break"

        yeni_degerler = {}
        alan_adlari = {
            "brutKg": "Brüt KG",
            "navlun": "Navlun",
            "sigorta": "Sigorta",
            "ilaveYurtdisi": "İlave Yurtdışı",
        }

        for key in EXCEL_MANUEL_ALANLAR:
            ham = self.excel_duzenle_varlar[key].get().strip()
            if ham and not self.sayi_gecerli_mi(ham):
                messagebox.showerror("Sayı Hatası", f"{alan_adlari[key]} geçerli sayı olmalı.")
                return "break"
            yeni_degerler[key] = self.sayi_duzelt(ham) if ham else ""

        kalem = self.kalemler[self.excel_duzenlenen_index]
        kalem.update(yeni_degerler)
        self.upsert_record(kalem)
        self.excel_onizleme_guncelle(secili_index=self.excel_duzenlenen_index)
        self._kayit_bilgisi_yaz(f"Excel alanları {kalem['sira']}. satır için kaydedildi")
        return "break"

    def excel_satirlarini_uret(self, kayitlar=None):
        satirlar = []
        kaynak_kayitlar = self.kalemler if kayitlar is None else kayitlar

        for kalem in kaynak_kayitlar:
            row = {h: "" for h in EXCEL_HEADERS}

            for kolon in EXCEL_SABIT_BOS:
                row[kolon] = ""
            for kolon in EXCEL_SABIT_DOLU:
                row[kolon] = self._excel_sabit_degeri(kolon)

            row["Fatura Tutar"] = self.excel_metni_sayi_formatla(kalem.get("deger", ""))
            row["Satış Miktarı"] = self.excel_metni_sayi_formatla(kalem.get("miktar", ""))
            row["Brüt KG"] = self.excel_metni_sayi_formatla(kalem.get("brutKg", ""))
            row["Net KG"] = self.excel_metni_sayi_formatla(kalem.get("miktar", ""))
            row["Tamamlayıcı Ölçü Miktar"] = self.excel_metni_sayi_formatla(kalem.get("miktar", ""))
            row["Kap Adet"] = self.excel_metni_sayi_formatla(kalem.get("balyaSayisi", ""))
            row["Navlun"] = self.excel_metni_sayi_formatla(kalem.get("navlun", ""))
            row["Sigorta"] = self.excel_metni_sayi_formatla(kalem.get("sigorta", ""))
            row["İlave Yurtdışı"] = self.excel_metni_sayi_formatla(kalem.get("ilaveYurtdisi", ""))
            row["Fatura No"] = ""
            row["Fatura Tarihi"] = ""

            satirlar.append(row)

        return satirlar

    def excel_onizleme_guncelle(self, secili_index=None):
        if secili_index is None:
            secim = self.excel_tree.selection()
            if secim:
                secili_index = self.excel_tree.index(secim[0])

        for item in self.excel_tree.get_children():
            self.excel_tree.delete(item)

        for kalem in self.kalemler:
            self.excel_tree.insert(
                "",
                "end",
                values=(
                    self._excel_isaret_simgesi(kalem),
                    kalem.get("sira", ""),
                    kalem.get("belgeNo", ""),
                    kalem.get("faturaNo", ""),
                    kalem.get("faturaTarihi", ""),
                    kalem.get("deger", ""),
                    kalem.get("miktar", ""),
                    kalem.get("balyaSayisi", ""),
                    kalem.get("brutKg", ""),
                    kalem.get("navlun", ""),
                    kalem.get("sigorta", ""),
                    kalem.get("ilaveYurtdisi", ""),
                )
            )

        self.excel_ozet_var.set(f"Excel için hazır satır sayısı: {len(self.kalemler)}")

        if secili_index is not None and 0 <= secili_index < len(self.excel_tree.get_children()):
            item = self.excel_tree.get_children()[secili_index]
            self.excel_tree.selection_set(item)
            self.excel_tree.focus(item)
            self.excel_seciliyi_duzenleme_paneline_yukle()
        else:
            self.excel_seciliyi_duzenleme_paneline_yukle()

    def excel_olustur(self, yalnizca_secilen=False):
        if Workbook is None:
            messagebox.showerror("Kütüphane Yok", "Excel üretmek için openpyxl gerekli.")
            return

        if self.excel_duzenlenen_index is not None:
            self.excel_secili_satira_uygula(sessiz=True)

        kayitlar = self.excel_isaretli_kayitlari_getir() if yalnizca_secilen else self.kalemler
        if yalnizca_secilen and not kayitlar:
            messagebox.showwarning("İşaret Yok", "Önce en az bir satırı işaretle.")
            return

        satirlar = self.excel_satirlarini_uret(kayitlar=kayitlar)
        if not satirlar:
            messagebox.showwarning("Satır Yok", "Önce XML sekmesinde en az 1 kalem ekle.")
            return

        wb = Workbook()
        ws = wb.active
        ws.title = "KALEMLER"

        for col_idx, header in enumerate(EXCEL_HEADERS, start=1):
            ws.cell(row=1, column=col_idx, value=header)

        for row_idx, row_data in enumerate(satirlar, start=2):
            for col_idx, header in enumerate(EXCEL_HEADERS, start=1):
                ws.cell(row=row_idx, column=col_idx, value=row_data.get(header, ""))

        from openpyxl.utils import get_column_letter
        for col_idx, header in enumerate(EXCEL_HEADERS, start=1):
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max(len(header) + 2, 14), 28)

        dosya_adi = (self.taslak_adi_var.get().strip() or "kalemler") + "_excel.xlsx"
        yol = filedialog.asksaveasfilename(
            title="Excel Kaydet",
            defaultextension=".xlsx",
            initialfile=dosya_adi,
            filetypes=[("Excel Dosyası", "*.xlsx"), ("Tüm Dosyalar", "*.*")]
        )
        if not yol:
            return

        try:
            wb.save(yol)
            messagebox.showinfo("Başarılı", f"Excel oluşturuldu.\n\n{yol}")
        except Exception as e:
            messagebox.showerror("Hata", f"Excel kaydedilirken hata oluştu:\n{e}")

    # -----------------------------------------------------
    # Sefer yönetimi metodları
    # -----------------------------------------------------
    def _sefer_combobox_guncelle(self):
        seferler = self.persistence.list_seferler(tumu=True)
        self._sefer_listesi = seferler
        etiketler = [
            f"{s['ad']} (#{s['id']})" + (" [arşiv]" if s["durum"] == "arsiv" else "")
            for s in seferler
        ]
        self.sefer_combobox["values"] = etiketler
        if self.aktif_sefer_id:
            for i, s in enumerate(seferler):
                if s["id"] == self.aktif_sefer_id:
                    self.sefer_combobox.current(i)
                    break

    def _sefer_secildi(self, event=None):
        secim_index = self.sefer_combobox.current()
        if secim_index < 0 or not hasattr(self, "_sefer_listesi"):
            return
        sefer = self._sefer_listesi[secim_index]
        self.aktif_sefer_id = sefer["id"]
        self.persistence.set_active_sefer_id(self.aktif_sefer_id)
        self._kalemler_yukle()
        self._aktif_sefer_bilgisini_guncelle()

    def _aktif_sefer_bilgisini_guncelle(self):
        if not self.aktif_sefer_id or not hasattr(self, "_sefer_listesi"):
            self.aktif_sefer_bilgi_var.set("")
            return
        for s in self._sefer_listesi:
            if s["id"] == self.aktif_sefer_id:
                self.aktif_sefer_bilgi_var.set(f"Aktif: {s['ad']} (#{s['id']})")
                return

    def yeni_sefer_baslat(self):
        ad = simpledialog.askstring("Yeni Sefer", "Sefer adı:", parent=self.root)
        if not ad or not ad.strip():
            return
        sefer_id = self.persistence.create_sefer(ad.strip())
        self.aktif_sefer_id = sefer_id
        self.persistence.set_active_sefer_id(sefer_id)
        self.kalemler.clear()
        for item in self.tree.get_children():
            self.tree.delete(item)
        self._ozeti_guncelle()
        self.excel_onizleme_guncelle()
        self.belgeler_listesini_guncelle()
        self.kalem_alanlarini_temizle()
        self._sefer_combobox_guncelle()
        self._aktif_sefer_bilgisini_guncelle()
        self._kayit_bilgisi_yaz(f"Yeni sefer başlatıldı: {ad.strip()}")

    def seferi_arsivle(self):
        if not self.aktif_sefer_id:
            messagebox.showwarning("Sefer Yok", "Aktif sefer bulunamadı.")
            return
        if not messagebox.askyesno("Arşivle", "Aktif sefer arşivlensin mi? Veriler kaybolmaz."):
            return
        self.persistence.archive_sefer(self.aktif_sefer_id)
        self._sefer_combobox_guncelle()
        self._kayit_bilgisi_yaz("Sefer arşivlendi.")

    def _filtre_degisti(self):
        self.aktif_tamamlanma_filtre = None
        self.tamamlanma_filtre_var.set("")
        self._kalemler_yukle()

    def _tamamlanma_no_filtrele(self):
        grup_no = self.tamamlanma_filtre_var.get().strip()
        if not grup_no:
            messagebox.showwarning("Eksik Bilgi", "Tamamlanma No giriniz.")
            return
        self.aktif_tamamlanma_filtre = grup_no
        self.tamamlananlar_goster_var.set(True)
        self._kalemler_yukle()

    def _filtre_temizle(self):
        self.aktif_tamamlanma_filtre = None
        self.tamamlanma_filtre_var.set("")
        self.tamamlananlar_goster_var.set(False)
        self._kalemler_yukle()

    def _kalemler_yukle(self):
        if not self.aktif_sefer_id:
            return
        show_completed = self.tamamlananlar_goster_var.get()
        try:
            kalemler = self.persistence.load_kalemler(
                self.aktif_sefer_id,
                show_completed=show_completed,
                tamamlanma_grup_no=self.aktif_tamamlanma_filtre,
            )
        except Exception as exc:
            messagebox.showerror("Yükleme Hatası", f"Kalemler yüklenemedi:\n{exc}")
            return
        self.kalemler = kalemler
        for item in self.tree.get_children():
            self.tree.delete(item)
        for kalem in self.kalemler:
            self._kalem_treeye_ekle(kalem)
        self._ozeti_guncelle()
        self.excel_onizleme_guncelle()
        self.belgeler_listesini_guncelle()

    def secilileri_tamamlandi_yap(self):
        secim = self.tree.selection()
        if not secim:
            messagebox.showwarning("Seçim Yok", "Önce listeden en az bir kalem seç.")
            return
        grup_no = simpledialog.askstring(
            "Tamamlanma No", "Tamamlanma numarası girin (örn: 55):", parent=self.root
        )
        if not grup_no or not grup_no.strip():
            return
        grup_no = grup_no.strip()
        indeksler = [self.tree.index(item) for item in secim]
        kalem_ids = [
            self.kalemler[i]["db_id"]
            for i in indeksler
            if 0 <= i < len(self.kalemler) and self.kalemler[i].get("db_id")
        ]
        if not kalem_ids:
            return
        try:
            self.persistence.complete_kalemler(kalem_ids, self.aktif_sefer_id, grup_no)
        except Exception as exc:
            messagebox.showerror("Hata", f"Tamamlandı yapılamadı:\n{exc}")
            return
        self._kalemler_yukle()
        self._kayit_bilgisi_yaz(f"{len(kalem_ids)} kalem tamamlandı (Grup: {grup_no})")

    def tamamlandiyi_geri_al(self):
        secim = self.tree.selection()
        if not secim:
            messagebox.showwarning("Seçim Yok", "Önce listeden en az bir kalem seç.")
            return
        indeksler = [self.tree.index(item) for item in secim]
        kalem_ids = []
        for i in indeksler:
            if 0 <= i < len(self.kalemler):
                k = self.kalemler[i]
                if k.get("db_id") and k.get("durum") == "tamamlandi":
                    kalem_ids.append(k["db_id"])
        if not kalem_ids:
            messagebox.showinfo("Bilgi", "Seçili kalemler arasında tamamlanmış kalem bulunamadı.")
            return
        try:
            self.persistence.undo_complete_kalemler(kalem_ids)
        except Exception as exc:
            messagebox.showerror("Hata", f"Geri alma başarısız:\n{exc}")
            return
        self._kalemler_yukle()
        self._kayit_bilgisi_yaz(f"{len(kalem_ids)} kalem tekrar bekliyor durumuna alındı")


def main():
    root = tk.Tk()
    BasvuruXMLVeExcel(root)
    root.mainloop()


if __name__ == "__main__":
    main()
