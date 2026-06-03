from __future__ import annotations

import re
from io import BytesIO
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st


DATA_FILE = Path("data_omma.xlsx")
SHEET_NAME = "DATA CLEARING FINAL"
IDENTITY_COL_COUNT = 17

PALETTE = [
    "#0E7490",
    "#F97316",
    "#7C3AED",
    "#16A34A",
    "#DC2626",
    "#2563EB",
    "#DB2777",
    "#65A30D",
    "#0891B2",
    "#9333EA",
    "#CA8A04",
    "#475569",
]

PLOT_CONFIG = {
    "displaylogo": False,
    "modeBarButtonsToRemove": ["lasso2d", "select2d"],
}

LANGUAGE_OPTIONS = ("Indonesia", "English")
DATA_SOURCE_EXISTING = "existing"
DATA_SOURCE_UPLOAD = "upload"
UPLOAD_FILE_TYPES = ["xlsx", "csv", "parquet"]
MAX_MONTHLY_HOUSEHOLD_MONEY = 250_000_000
EXPENSE_EXTRA_ZERO_THRESHOLD = 20_000_000
EXPENSE_INCOME_RATIO_FOR_EXTRA_ZERO = 2
MAX_LAND_SIZE_HECTARE = 50
MAX_TREE_COUNT = 5_000
MAX_FARMING_YEARS = 80


def language_code(choice: str) -> str:
    return "en" if choice == "English" else "id"


APP_LANG = language_code(st.session_state.get("language_choice", "Indonesia"))

TEXT = {
    "id": {
        "page_title": "Analisis Deskriptif Sawit",
        "main_title": "Profil Deskriptif Responden Petani Sawit",
        "intro": "Ringkasan ini menampilkan profil responden secara deskriptif agar karakter data mudah dibaca dari komposisi kategori, sebaran numerik, dan statistik inti.",
        "sidebar_title": "Pengaturan",
        "language": "Bahasa Tampilan",
        "data_source": "Sumber Data",
        "use_available_data": "Pakai Data Tersedia",
        "upload_data": "Unggah Data",
        "upload_file": "Unggah File Data",
        "upload_help": "Gunakan file Excel, CSV, atau Parquet dengan struktur kolom data sawit yang sama.",
        "upload_missing": "Silakan unggah file data untuk memakai mode unggah.",
        "source_available_caption": "Menggunakan data tersedia: {file_name}",
        "source_upload_caption": "Menggunakan data unggahan: {file_name}",
        "file_not_found": "File data tidak ditemukan: {file_name}",
        "data_load_error": "Data gagal dimuat: {error}",
        "unsupported_file_type": "Format file tidak didukung: {extension}",
        "missing_required_columns": "Kolom wajib tidak ditemukan: {columns}",
        "no_readable_excel": "Tidak ada sheet Excel yang dapat dibaca.",
        "invalid_age_filter": "Kolom umur tidak memiliki data valid untuk filter.",
        "filters": "Filter Profil",
        "reload_data": "Muat Ulang Data",
        "gender": "Jenis Kelamin",
        "education": "Pendidikan",
        "land_status": "Status Lahan",
        "training_status": "Status Pelatihan",
        "age_range": "Rentang Umur",
        "respondents": "Responden",
        "avg_age": "Rata-rata Umur",
        "majority_gender": "Gender Dominan",
        "median_family": "Median Anggota Keluarga",
        "median_land": "Median Lahan",
        "median_production": "Median Produksi",
        "no_data": "Tidak ada data pada kombinasi filter saat ini.",
        "overview": "Ikhtisar",
        "categorical": "Variabel Kategori",
        "numeric": "Variabel Numerik",
        "processed_data": "Data Terolah",
        "profile_snapshot": "Gambaran Cepat Profil",
        "gender_pie": "Perbandingan Laki-laki dan Perempuan",
        "top_regions": "Wilayah Responden Terbanyak",
        "education_profile": "Profil Pendidikan Responden",
        "numeric_summary": "Statistik Deskriptif Numerik",
        "data_completeness": "Kelengkapan Data",
        "missing_count": "Jumlah Kosong",
        "missing_percent": "Persentase Kosong",
        "category": "Kategori",
        "count": "Jumlah",
        "percentage": "Persentase",
        "count_axis": "Jumlah Responden",
        "frequency": "Frekuensi",
        "not_filled": "Tidak Diisi",
        "other": "Lainnya",
        "median": "Median",
        "mean": "Rata-rata",
        "std": "Simpangan Baku",
        "minimum": "Minimum",
        "maximum": "Maksimum",
        "variable": "Variabel",
        "distribution_of": "Distribusi {variable}",
        "boxplot_of": "Boxplot {variable}",
        "composition_of": "Komposisi {variable}",
        "descriptive_stats_for": "Statistik deskriptif {variable}",
        "valid_data": "Data valid",
        "numeric_insight": "Median {variable} adalah {median}; rentang data berada pada {minimum} sampai {maximum}.",
        "categorical_insight": "Kategori terbesar: {category} ({count} responden, {percent}).",
        "sanitized_dataset": "Dataset Analisis Tanpa Identitas Pribadi",
        "download_csv": "Unduh CSV Data Terolah",
        "download_file_name": "profil_deskriptif_sawit_terolah.csv",
    },
    "en": {
        "page_title": "Oil Palm Descriptive Analysis",
        "main_title": "Descriptive Profile of Oil Palm Farmer Respondents",
        "intro": "This summary describes respondent profiles through category composition, numeric distributions, and core descriptive statistics.",
        "sidebar_title": "Settings",
        "language": "Display Language",
        "data_source": "Data Source",
        "use_available_data": "Use Available Data",
        "upload_data": "Upload Data",
        "upload_file": "Upload Data File",
        "upload_help": "Use an Excel, CSV, or Parquet file with the same oil palm data column structure.",
        "upload_missing": "Please upload a data file to use upload mode.",
        "source_available_caption": "Using available data: {file_name}",
        "source_upload_caption": "Using uploaded data: {file_name}",
        "file_not_found": "Data file was not found: {file_name}",
        "data_load_error": "Data could not be loaded: {error}",
        "unsupported_file_type": "Unsupported file format: {extension}",
        "missing_required_columns": "Required columns are missing: {columns}",
        "no_readable_excel": "No readable Excel sheet was found.",
        "invalid_age_filter": "The age column has no valid data for filtering.",
        "filters": "Profile Filters",
        "reload_data": "Reload Data",
        "gender": "Gender",
        "education": "Education",
        "land_status": "Land Status",
        "training_status": "Training Status",
        "age_range": "Age Range",
        "respondents": "Respondents",
        "avg_age": "Average Age",
        "majority_gender": "Dominant Gender",
        "median_family": "Median Household Size",
        "median_land": "Median Land Size",
        "median_production": "Median Production",
        "no_data": "No data matches the current filter combination.",
        "overview": "Overview",
        "categorical": "Categorical Variables",
        "numeric": "Numeric Variables",
        "processed_data": "Processed Data",
        "profile_snapshot": "Quick Profile Snapshot",
        "gender_pie": "Male and Female Comparison",
        "top_regions": "Largest Respondent Regions",
        "education_profile": "Respondent Education Profile",
        "numeric_summary": "Numeric Descriptive Statistics",
        "data_completeness": "Data Completeness",
        "missing_count": "Missing Count",
        "missing_percent": "Missing Percentage",
        "category": "Category",
        "count": "Count",
        "percentage": "Percentage",
        "count_axis": "Number of Respondents",
        "frequency": "Frequency",
        "not_filled": "Missing",
        "other": "Other",
        "median": "Median",
        "mean": "Mean",
        "std": "Standard Deviation",
        "minimum": "Minimum",
        "maximum": "Maximum",
        "variable": "Variable",
        "distribution_of": "Distribution of {variable}",
        "boxplot_of": "Boxplot of {variable}",
        "composition_of": "Composition of {variable}",
        "descriptive_stats_for": "Descriptive statistics for {variable}",
        "valid_data": "Valid data",
        "numeric_insight": "The median {variable} is {median}; values range from {minimum} to {maximum}.",
        "categorical_insight": "Largest category: {category} ({count} respondents, {percent}).",
        "sanitized_dataset": "Analysis Dataset Without Personal Identity",
        "download_csv": "Download Processed CSV",
        "download_file_name": "oil_palm_descriptive_profile_processed.csv",
    },
}

CATEGORY_TRANSLATIONS = {
    "en": {
        "Laki-laki": "Male",
        "Perempuan": "Female",
        "Tidak Diisi": "Missing",
        "Pernah Mengikuti": "Attended",
        "Belum Pernah": "Never Attended",
        "Petani": "Farmer",
        "Buruh": "Laborer",
        "Pedagang": "Trader",
        "Guru": "Teacher",
        "Wiraswasta": "Entrepreneur",
        "Tidak Sekolah": "No Schooling",
        "SD": "Elementary School",
        "SMP": "Junior High School",
        "SMA": "Senior High School",
        "Diploma": "Diploma",
        "Sarjana": "Bachelor",
        "Diploma/Sarjana": "Diploma/Bachelor",
        "Lainnya": "Other",
        "Milik Sendiri": "Owned",
        "Milik Mertua": "Owned by In-laws",
        "Sewa": "Rented",
        "Bagi Hasil": "Profit Sharing",
        "Pekerja": "Worker",
        "Belum Pernah": "Never Attended",
        "PTP": "PTP",
        "Pemupukan": "Fertilization",
        "Panen Sawit": "Oil Palm Harvesting",
        "Penanaman Sawit": "Oil Palm Planting",
        "Penanaman Sawit yang Baik": "Good Oil Palm Planting",
        "Pelatihan Pertanian Sawit": "Oil Palm Agriculture Training",
        "Pelatihan Sawit": "Oil Palm Training",
    }
}

VARIABLE_LABELS = {
    "No": {"id": "No", "en": "No"},
    "Jenis Kelamin": {"id": "Jenis Kelamin", "en": "Gender"},
    "Wilayah": {"id": "Wilayah", "en": "Region"},
    "Umur (tahun)": {"id": "Umur (tahun)", "en": "Age (years)"},
    "Jumlah Anggota Keluarga": {
        "id": "Jumlah Anggota Keluarga",
        "en": "Household Members",
    },
    "Pekerjaan Utama Pencari Nafkah": {
        "id": "Pekerjaan Utama Pencari Nafkah",
        "en": "Main Breadwinner Occupation",
    },
    "Pendidikan Terakhir": {"id": "Pendidikan Terakhir", "en": "Highest Education"},
    "Penghasilan Keluarga per Bulan (Rupiah)": {
        "id": "Penghasilan Keluarga per Bulan (Rupiah)",
        "en": "Monthly Household Income (Rupiah)",
    },
    "Pengeluaran Keluarga per Bulan (Rupiah)": {
        "id": "Pengeluaran Keluarga per Bulan (Rupiah)",
        "en": "Monthly Household Expense (Rupiah)",
    },
    "Luas Lahan Sawit (Hektar)": {
        "id": "Luas Lahan Sawit (Hektar)",
        "en": "Oil Palm Land Size (hectares)",
    },
    "Status Lahan": {"id": "Status Lahan", "en": "Land Status"},
    "Jumlah Pohon Sawit (pohon)": {
        "id": "Jumlah Pohon Sawit (pohon)",
        "en": "Number of Oil Palm Trees",
    },
    "Produksi Rata-rata (ton/bulan)": {
        "id": "Produksi Rata-rata (ton/bulan)",
        "en": "Average Production (tons/month)",
    },
    "Lama menjadi Petani Sawit (tahun)": {
        "id": "Lama menjadi Petani Sawit (tahun)",
        "en": "Years as Oil Palm Farmer",
    },
    "Jenis Pelatihan Sawit": {
        "id": "Jenis Pelatihan Sawit",
        "en": "Oil Palm Training Type",
    },
    "Status Pelatihan Sawit": {
        "id": "Status Pelatihan Sawit",
        "en": "Oil Palm Training Status",
    },
    "Surplus Pendapatan (Rupiah/bulan)": {
        "id": "Surplus Pendapatan (Rupiah/bulan)",
        "en": "Income Surplus (Rupiah/month)",
    },
    "Penghasilan (juta Rupiah/bulan)": {
        "id": "Penghasilan (juta Rupiah/bulan)",
        "en": "Income (million Rupiah/month)",
    },
    "Pengeluaran (juta Rupiah/bulan)": {
        "id": "Pengeluaran (juta Rupiah/bulan)",
        "en": "Expense (million Rupiah/month)",
    },
}


st.set_page_config(
    page_title=TEXT[APP_LANG]["page_title"],
    page_icon=None,
    layout="wide",
    initial_sidebar_state="expanded",
)


st.markdown(
    """
    <style>
    :root {
        --ink: #111827;
        --muted: #4b5563;
        --line: #d1d5db;
        --panel: #ffffff;
        --bg: #f4f7fb;
        --accent: #0e7490;
        --accent-soft: #ecfeff;
    }

    .stApp {
        background: var(--bg);
        color: var(--ink);
    }

    h1, h2, h3 {
        color: var(--ink);
        letter-spacing: 0;
    }

    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
        max-width: 1440px;
    }

    [data-testid="stMetric"] {
        background: var(--panel);
        border: 1px solid #e5e7eb;
        border-radius: 8px;
        padding: 0.85rem 1rem;
        box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04);
    }

    [data-testid="stMetricLabel"] {
        color: var(--muted);
        font-size: 0.82rem;
    }

    [data-testid="stMetricValue"] {
        color: var(--ink);
        font-size: 1.55rem;
        font-weight: 700;
    }

    div[data-testid="stDataFrame"] {
        border: 1px solid #e5e7eb;
        border-radius: 8px;
        overflow: hidden;
    }

    section[data-testid="stSidebar"] {
        background: #ffffff;
        border-right: 1px solid #e5e7eb;
    }

    .hero-panel {
        background: linear-gradient(135deg, #ffffff 0%, #ecfeff 42%, #fff7ed 100%);
        border: 1px solid #dbeafe;
        border-radius: 8px;
        padding: 1.1rem 1.2rem;
        margin-bottom: 1rem;
        box-shadow: 0 10px 30px rgba(15, 23, 42, 0.06);
    }

    .hero-panel h1 {
        margin: 0 0 0.35rem 0;
        font-size: clamp(1.65rem, 2.5vw, 2.45rem);
        line-height: 1.12;
    }

    .small-note {
        color: var(--muted);
        font-size: 0.9rem;
        margin: 0;
    }

    .insight-card {
        background: #ffffff;
        border: 1px solid #e5e7eb;
        border-left: 5px solid var(--accent);
        border-radius: 8px;
        padding: 0.8rem 0.95rem;
        margin: 0.35rem 0 0.8rem 0;
        color: var(--ink);
        box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04);
    }

    .insight-card strong {
        display: block;
        font-size: 0.82rem;
        color: var(--muted);
        margin-bottom: 0.25rem;
        text-transform: uppercase;
        letter-spacing: 0.03em;
    }

    .insight-card span {
        font-size: 0.98rem;
        line-height: 1.45;
    }

    div[data-testid="stExpander"] {
        background: #ffffff;
        border-radius: 8px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def normalize_spaces(value: object) -> str:
    return re.sub(r"\s+", " ", str(value).strip())


def clean_column_name(value: object) -> str:
    text = normalize_spaces(value)
    text = re.sub(r"^[A-Za-z]\d+\.\s*", "", text)
    text = re.sub(r"^[A-Za-z]\d+\s+", "", text)
    text = text.replace("No Hp", "No HP").replace("No hp", "No HP")
    return text


def parse_number_token(token: str) -> float:
    token = token.strip()
    if not token:
        return np.nan

    sign = ""
    if token.startswith("-"):
        sign = "-"
        token = token[1:]

    if "." in token and "," in token:
        last_dot = token.rfind(".")
        last_comma = token.rfind(",")
        decimal_separator = "," if last_comma > last_dot else "."
        fractional = token.split(decimal_separator)[-1]
        if len(fractional) <= 2:
            thousands_separator = "." if decimal_separator == "," else ","
            token = token.replace(thousands_separator, "").replace(decimal_separator, ".")
        else:
            token = token.replace(".", "").replace(",", "")
    elif token.count(",") > 1 or token.count(".") > 1:
        separator = "," if "," in token else "."
        parts = token.split(separator)
        if len(parts[-1]) <= 2 and all(len(part) == 3 for part in parts[1:-1]):
            token = "".join(parts[:-1]) + "." + parts[-1]
        else:
            token = "".join(parts)
    elif "," in token or "." in token:
        separator = "," if "," in token else "."
        whole, fractional = token.split(separator, 1)
        whole_without_sign = whole.lstrip("-")
        if len(fractional) <= 2 or whole_without_sign == "0":
            token = f"{whole}.{fractional}"
        elif len(fractional) == 3:
            token = whole + fractional
        else:
            token = f"{whole}.{fractional}"

    try:
        return float(sign + token)
    except ValueError:
        return np.nan


def extract_number(value: object) -> float:
    if pd.isna(value):
        return np.nan
    if isinstance(value, (int, float, np.integer, np.floating)):
        return float(value)

    text = str(value).strip().lower()
    range_match = re.search(
        r"(\d+(?:[.,]\d+)*)\s*(?:-|–|—|s/d|sd|hingga|sampai|to)\s*(\d+(?:[.,]\d+)*)",
        text,
    )
    if range_match:
        low = parse_number_token(range_match.group(1))
        high = parse_number_token(range_match.group(2))
        if not pd.isna(low) and not pd.isna(high):
            return (low + high) / 2

    match = re.search(r"-?\d+(?:[.,]\d+)*", text)
    if not match:
        return np.nan
    return parse_number_token(match.group(0))


def parse_money_rupiah(value: object) -> float:
    number = extract_number(value)
    if pd.isna(number):
        return np.nan

    text = str(value).lower()
    if re.search(r"\b(juta|jt)\b", text):
        return number * 1_000_000
    if re.search(r"\b(ribu|rb)\b", text):
        return number * 1_000

    if 1_000 <= number < 100_000:
        number *= 1_000
    while number > MAX_MONTHLY_HOUSEHOLD_MONEY:
        number /= 1_000
    return number


def parse_land_hectare(value: object) -> float:
    number = extract_number(value)
    if pd.isna(number):
        return np.nan

    text = str(value).lower()
    if re.search(r"\b(m|m2|meter)\b|m²", text) and "ha" not in text:
        return number / 10_000
    while number > MAX_LAND_SIZE_HECTARE:
        number /= 100
    return number


def parse_tree_count(value: object) -> float:
    number = extract_number(value)
    if pd.isna(number):
        return np.nan
    while number > MAX_TREE_COUNT:
        number /= 1_000
    return number


def parse_farming_years(value: object) -> float:
    number = extract_number(value)
    if pd.isna(number) or number > MAX_FARMING_YEARS:
        return np.nan
    return number


def parse_production_ton(value: object) -> float:
    number = extract_number(value)
    if pd.isna(number):
        return np.nan
    text = str(value).lower()
    if "kg" in text and "ton" not in text:
        return number / 1000
    return number


def clean_category(value: object, missing: str = "Tidak Diisi") -> str:
    if pd.isna(value):
        return missing
    text = normalize_spaces(value)
    if not text or text.lower() in {"nan", "none", "-"}:
        return missing
    return text.title()


def normalize_gender(value: object) -> str:
    text = clean_category(value).lower().replace("-", " ").replace("_", " ")
    text = normalize_spaces(text)
    if "laki" in text:
        return "Laki-laki"
    if "perempuan" in text or "prempuan" in text:
        return "Perempuan"
    return clean_category(value)


def normalize_job(value: object) -> str:
    text = clean_category(value).lower()
    if text in {"ptani", "petani sawit", "tani"}:
        return "Petani"
    if "petani" in text:
        return "Petani"
    return clean_category(value)


def normalize_education(value: object) -> str:
    text = clean_category(value).lower().replace(".", "")
    if text in {"sd", "sekolah dasar"}:
        return "SD"
    if text in {"smp", "sltp"}:
        return "SMP"
    if text in {"sma", "smk", "slta"}:
        return "SMA"
    if "tidak" in text and "sekolah" in text:
        return "Tidak Sekolah"
    if "diploma" in text and "sarjana" in text:
        return "Diploma/Sarjana"
    if "diploma" in text:
        return "Diploma"
    if "sarjana" in text or text in {"s1", "strata 1"}:
        return "Sarjana"
    return "Lainnya"


def normalize_land_status(value: object) -> str:
    text = clean_category(value).lower()
    mapping = {
        "milik sendiri": "Milik Sendiri",
        "milik mertua": "Milik Mertua",
        "sewa": "Sewa",
        "bagi hasil": "Bagi Hasil",
        "pekerja": "Pekerja",
        "lainnya": "Lainnya",
    }
    return mapping.get(text, clean_category(value))


def normalize_training_type(value: object) -> str:
    if pd.isna(value):
        return "Tidak Diisi"

    text = normalize_spaces(value).lower()
    if not text or text in {"nan", "none", "-"}:
        return "Tidak Diisi"
    if "belum" in text or "tidak pernah" in text:
        return "Belum Pernah"
    if "ptp" in text:
        return "PTP"
    if "pupuk" in text:
        return "Pemupukan"
    if "panen" in text:
        return "Panen Sawit"
    if "penanaman" in text or "menanam" in text:
        if "baik" in text:
            return "Penanaman Sawit yang Baik"
        return "Penanaman Sawit"
    if "pertanian" in text:
        return "Pelatihan Pertanian Sawit"
    if "sawit" in text:
        return "Pelatihan Sawit"
    return clean_category(value)


def normalize_training_status(value: object) -> str:
    training_type = normalize_training_type(value)
    if training_type == "Tidak Diisi":
        return "Tidak Diisi"
    if training_type == "Belum Pernah":
        return "Belum Pernah"
    return "Pernah Mengikuti"


def extract_region(value: object) -> str:
    if pd.isna(value):
        return "Tidak Diisi"

    original = normalize_spaces(value)
    text = original.lower()
    text = text.replace(".", " ").replace(",", " ")
    text = normalize_spaces(text)

    rules = [
        (("sukadamai", "suka damai", "s damai"), "Sukadamai"),
        (("bangsal aceh", "b aceh", "bangsal"), "Bangsal Aceh"),
        (("bagan keladi", "b keladi", "keladi"), "Bagan Keladi"),
        (("pantai raja",), "Pantai Raja"),
        (("tanjung penyembal",), "Tanjung Penyembal"),
        (("basilan baru",), "Basilan Baru"),
        (("lubuk gaung",), "Lubuk Gaung"),
        (("mataram",), "Mataram"),
        (("purnama",), "Purnama"),
        (("bukit kapur",), "Bukit Kapur"),
        (("merasan",), "Merasan"),
        (("sei sembilan", "sungai sembilan", "sungai 9", "sembilan", "seisimbilan"), "Sungai Sembilan"),
        (("dumai barat",), "Dumai Barat"),
    ]

    for keywords, label in rules:
        if any(keyword in text for keyword in keywords):
            return label

    first_part = re.split(r",|\bkec\b|\bkecamatan\b|\bkab\b|\bkota\b", original, flags=re.IGNORECASE)[0]
    first_part = normalize_spaces(re.sub(r"^(desa|kelurahan)\s+", "", first_part, flags=re.IGNORECASE))
    return first_part.title() if first_part else "Wilayah Lainnya"


REQUIRED_RAW_COLUMNS = [
    "No",
    "Nama",
    "No HP",
    "Jenis Kelamin",
    "Alamat",
    "Umur (tahun)",
    "Jumlah Anggota Keluarga",
    "Pekerjaan Utama Pencari Nafkah",
    "Pendidikan Terakhir",
    "Penghasilan Keluarga per Bulan (Rupiah)",
    "Pengeluaran Keluarga per Bulan (Rupiah)",
    "Luas Lahan Sawit (Hektar)",
    "Status Lahan",
    "Jumlah Pohon Sawit (pohon)",
    "Produksi Rata-rata (ton/bulan)",
    "Lama menjadi Petani Sawit (tahun)",
    "Pelatihan yang pernah diikuti terkait Sawit",
]


def standardize_raw_columns(raw: pd.DataFrame) -> pd.DataFrame:
    raw = raw.copy()
    raw.columns = [clean_column_name(col) for col in raw.columns]
    return raw.dropna(how="all").copy()


def missing_raw_columns(raw: pd.DataFrame) -> list[str]:
    return [column for column in REQUIRED_RAW_COLUMNS if column not in raw.columns]


def read_excel_data(source: str | BytesIO) -> pd.DataFrame:
    candidates: list[pd.DataFrame] = []
    errors: list[Exception] = []

    for sheet_name in (SHEET_NAME, 0):
        for header in (1, 0):
            if hasattr(source, "seek"):
                source.seek(0)
            try:
                raw = pd.read_excel(
                    source,
                    sheet_name=sheet_name,
                    header=header,
                    usecols=range(IDENTITY_COL_COUNT),
                )
            except Exception as error:
                errors.append(error)
                continue

            raw = standardize_raw_columns(raw)
            candidates.append(raw)
            if not missing_raw_columns(raw):
                return raw

    if candidates:
        return candidates[0]
    if errors:
        raise errors[0]
    raise ValueError(t("no_readable_excel"))


def read_raw_data(source: str | BytesIO, extension: str) -> pd.DataFrame:
    extension = extension.lower()
    if extension in {".xlsx", ".xls"}:
        return read_excel_data(source)
    if extension == ".csv":
        return standardize_raw_columns(pd.read_csv(source))
    if extension == ".parquet":
        return standardize_raw_columns(pd.read_parquet(source))
    raise ValueError(t("unsupported_file_type").format(extension=extension or "-"))


def process_raw_data(raw: pd.DataFrame) -> pd.DataFrame:
    raw = standardize_raw_columns(raw)
    missing_columns = missing_raw_columns(raw)
    if missing_columns:
        raise ValueError(t("missing_required_columns").format(columns=", ".join(missing_columns)))

    df = pd.DataFrame()
    df["No"] = raw["No"]
    df["Nama"] = raw["Nama"].astype(str)
    df["No HP"] = raw["No HP"]
    df["Jenis Kelamin"] = raw["Jenis Kelamin"].map(normalize_gender)
    df["Alamat"] = raw["Alamat"]
    df["Wilayah"] = raw["Alamat"].map(extract_region)
    df["Umur (tahun)"] = raw["Umur (tahun)"].map(extract_number)
    df["Jumlah Anggota Keluarga"] = raw["Jumlah Anggota Keluarga"].map(extract_number)
    df["Pekerjaan Utama Pencari Nafkah"] = raw["Pekerjaan Utama Pencari Nafkah"].map(normalize_job)
    df["Pendidikan Terakhir"] = raw["Pendidikan Terakhir"].map(normalize_education)
    df["Penghasilan Keluarga per Bulan (Rupiah)"] = raw[
        "Penghasilan Keluarga per Bulan (Rupiah)"
    ].map(parse_money_rupiah)
    df["Pengeluaran Keluarga per Bulan (Rupiah)"] = raw[
        "Pengeluaran Keluarga per Bulan (Rupiah)"
    ].map(parse_money_rupiah)
    extra_zero_expense = (
        (df["Pengeluaran Keluarga per Bulan (Rupiah)"] > EXPENSE_EXTRA_ZERO_THRESHOLD)
        & df["Penghasilan Keluarga per Bulan (Rupiah)"].notna()
        & (
            df["Pengeluaran Keluarga per Bulan (Rupiah)"]
            > df["Penghasilan Keluarga per Bulan (Rupiah)"] * EXPENSE_INCOME_RATIO_FOR_EXTRA_ZERO
        )
    )
    df.loc[extra_zero_expense, "Pengeluaran Keluarga per Bulan (Rupiah)"] /= 10
    df["Luas Lahan Sawit (Hektar)"] = raw["Luas Lahan Sawit (Hektar)"].map(parse_land_hectare)
    df["Status Lahan"] = raw["Status Lahan"].map(normalize_land_status)
    df["Jumlah Pohon Sawit (pohon)"] = raw["Jumlah Pohon Sawit (pohon)"].map(parse_tree_count)
    df["Produksi Rata-rata (ton/bulan)"] = raw["Produksi Rata-rata (ton/bulan)"].map(
        parse_production_ton
    )
    df["Lama menjadi Petani Sawit (tahun)"] = raw[
        "Lama menjadi Petani Sawit (tahun)"
    ].map(parse_farming_years)
    df["Jenis Pelatihan Sawit"] = raw[
        "Pelatihan yang pernah diikuti terkait Sawit"
    ].map(normalize_training_type)
    df["Status Pelatihan Sawit"] = raw[
        "Pelatihan yang pernah diikuti terkait Sawit"
    ].map(normalize_training_status)

    df["Surplus Pendapatan (Rupiah/bulan)"] = (
        df["Penghasilan Keluarga per Bulan (Rupiah)"]
        - df["Pengeluaran Keluarga per Bulan (Rupiah)"]
    )
    df["Penghasilan (juta Rupiah/bulan)"] = df[
        "Penghasilan Keluarga per Bulan (Rupiah)"
    ] / 1_000_000
    df["Pengeluaran (juta Rupiah/bulan)"] = df[
        "Pengeluaran Keluarga per Bulan (Rupiah)"
    ] / 1_000_000
    return df


@st.cache_data(show_spinner=False)
def load_existing_data(file_path: str, file_signature: int) -> pd.DataFrame:
    raw = read_raw_data(file_path, Path(file_path).suffix)
    return process_raw_data(raw)


@st.cache_data(show_spinner=False)
def load_uploaded_data(file_name: str, file_content: bytes) -> pd.DataFrame:
    raw = read_raw_data(BytesIO(file_content), Path(file_name).suffix)
    return process_raw_data(raw)


def format_integer(value: float) -> str:
    if pd.isna(value):
        return "-"
    text = f"{value:,.0f}"
    if APP_LANG == "id":
        return text.replace(",", ".")
    return text


def format_decimal(value: float, digits: int = 2) -> str:
    if pd.isna(value):
        return "-"
    text = f"{value:,.{digits}f}"
    if APP_LANG == "en":
        return text
    return text.replace(",", "X").replace(".", ",").replace("X", ".")


def format_rupiah(value: float) -> str:
    if pd.isna(value):
        return "-"
    if abs(value) >= 1_000_000:
        suffix = " M" if APP_LANG == "en" else " jt"
        return f"Rp{format_decimal(value / 1_000_000, 2)}{suffix}"
    return f"Rp{format_integer(value)}"


def t(key: str) -> str:
    return TEXT.get(APP_LANG, TEXT["id"]).get(key, TEXT["id"].get(key, key))


def variable_label(column: str) -> str:
    labels = VARIABLE_LABELS.get(column)
    if not labels:
        return column
    return labels.get(APP_LANG, labels["id"])


def category_label(value: object) -> str:
    if pd.isna(value):
        return t("not_filled")
    text = str(value)
    if APP_LANG == "en":
        return CATEGORY_TRANSLATIONS.get("en", {}).get(text, text)
    return text


def format_percent(value: float, digits: int = 1) -> str:
    return f"{format_decimal(value, digits)}%"


def make_insight(text: str, label: str = "") -> None:
    heading = f"<strong>{label}</strong>" if label else ""
    st.markdown(
        f"<div class='insight-card'>{heading}<span>{text}</span></div>",
        unsafe_allow_html=True,
    )


def apply_journal_layout(
    fig: go.Figure,
    title: str,
    x_title: str = "",
    y_title: str = "",
    height: int = 430,
    showlegend: bool = False,
) -> go.Figure:
    fig.update_layout(
        title={"text": title, "x": 0.015, "xanchor": "left", "font": {"size": 18}},
        template="plotly_white",
        height=height,
        font={"family": "Arial, sans-serif", "size": 13, "color": "#111827"},
        paper_bgcolor="white",
        plot_bgcolor="white",
        margin={"l": 68, "r": 28, "t": 72, "b": 64},
        showlegend=showlegend,
        legend={
            "orientation": "h",
            "yanchor": "bottom",
            "y": 1.02,
            "xanchor": "right",
            "x": 1,
        },
    )
    fig.update_xaxes(
        title=x_title,
        showline=True,
        linewidth=1,
        linecolor="#111827",
        gridcolor="#e5e7eb",
        zeroline=False,
        ticks="outside",
    )
    fig.update_yaxes(
        title=y_title,
        showline=True,
        linewidth=1,
        linecolor="#111827",
        gridcolor="#e5e7eb",
        zeroline=False,
        ticks="outside",
    )
    return fig


def prepare_category_counts(
    data: pd.DataFrame,
    column: str,
    top_n: int | None = None,
    order: list[str] | None = None,
) -> pd.DataFrame:
    series = data[column].copy()
    if pd.api.types.is_numeric_dtype(series):
        series = series.map(lambda value: "Tidak Diisi" if pd.isna(value) else format_integer(value))
    else:
        series = series.fillna("Tidak Diisi")
    counts = series.value_counts()

    if order:
        ordered = [item for item in order if item in counts.index]
        remaining = [item for item in counts.index if item not in ordered]
        counts = counts.reindex(ordered + remaining)

    if top_n and len(counts) > top_n:
        top_counts = counts.head(top_n)
        counts = pd.concat([top_counts, pd.Series({"Lainnya": counts.iloc[top_n:].sum()})])

    chart = counts.rename_axis("Kategori Asli").reset_index(name="Jumlah")
    chart["Kategori"] = chart["Kategori Asli"].map(category_label)
    chart["Persentase"] = chart["Jumlah"] / chart["Jumlah"].sum() * 100
    chart["Label"] = chart.apply(
        lambda row: f"{format_integer(row['Jumlah'])} ({format_percent(row['Persentase'])})",
        axis=1,
    )
    return chart


def category_summary(
    data: pd.DataFrame,
    column: str,
    top_n: int | None = None,
    order: list[str] | None = None,
) -> pd.DataFrame:
    chart = prepare_category_counts(data, column, top_n=top_n, order=order)
    return pd.DataFrame(
        {
            t("category"): chart["Kategori"],
            t("count"): chart["Jumlah"],
            t("percentage"): chart["Persentase"].map(format_percent),
        }
    )


def count_bar(
    data: pd.DataFrame,
    column: str,
    title: str,
    top_n: int | None = None,
    orientation: str = "v",
    order: list[str] | None = None,
    height: int = 430,
    label: str | None = None,
) -> go.Figure:
    chart = prepare_category_counts(data, column, top_n=top_n, order=order)
    variable = label or variable_label(column)

    colors = [PALETTE[i % len(PALETTE)] for i in range(len(chart))]

    if orientation == "h":
        chart = chart.sort_values("Jumlah", ascending=True)
        colors = [PALETTE[i % len(PALETTE)] for i in range(len(chart))]
        fig = go.Figure(
            go.Bar(
                x=chart["Jumlah"],
                y=chart["Kategori"],
                orientation="h",
                marker={"color": colors[: len(chart)], "line": {"color": "#111827", "width": 0.4}},
                text=chart["Label"],
                textposition="outside",
                cliponaxis=False,
                hovertemplate=(
                    f"{variable}: %{{y}}<br>{t('count')}: %{{x}}"
                    f"<br>{t('percentage')}: %{{customdata}}<extra></extra>"
                ),
                customdata=chart["Persentase"].map(format_percent),
            )
        )
        return apply_journal_layout(fig, title, t("count_axis"), variable, height)

    fig = go.Figure(
        go.Bar(
            x=chart["Kategori"],
            y=chart["Jumlah"],
            marker={"color": colors[: len(chart)], "line": {"color": "#111827", "width": 0.4}},
            text=chart["Label"],
            textposition="outside",
            cliponaxis=False,
            hovertemplate=(
                f"{variable}: %{{x}}<br>{t('count')}: %{{y}}"
                f"<br>{t('percentage')}: %{{customdata}}<extra></extra>"
            ),
            customdata=chart["Persentase"].map(format_percent),
        )
    )
    fig = apply_journal_layout(fig, title, variable, t("count_axis"), height)
    if len(chart) > 6:
        fig.update_xaxes(tickangle=-25)
    return fig


def donut_chart(
    data: pd.DataFrame,
    column: str,
    title: str,
    order: list[str] | None = None,
    height: int = 430,
    label: str | None = None,
) -> go.Figure:
    chart = prepare_category_counts(data, column, order=order)
    variable = label or variable_label(column)
    fig = go.Figure(
        go.Pie(
            labels=chart["Kategori"],
            values=chart["Jumlah"],
            hole=0.55,
            sort=False,
            marker={
                "colors": [PALETTE[i % len(PALETTE)] for i in range(len(chart))],
                "line": {"color": "#ffffff", "width": 2},
            },
            texttemplate="%{label}<br>%{value} (%{percent})",
            textposition="inside",
            hovertemplate=(
                f"{variable}: %{{label}}<br>{t('count')}: %{{value}}"
                f"<br>{t('percentage')}: %{{percent}}<extra></extra>"
            ),
        )
    )
    fig.update_layout(
        title={"text": title, "x": 0.015, "xanchor": "left", "font": {"size": 18}},
        template="plotly_white",
        height=height,
        font={"family": "Arial, sans-serif", "size": 13, "color": "#111827"},
        paper_bgcolor="white",
        margin={"l": 24, "r": 24, "t": 72, "b": 32},
        legend={
            "orientation": "h",
            "yanchor": "bottom",
            "y": -0.05,
            "xanchor": "center",
            "x": 0.5,
        },
        uniformtext={"minsize": 11, "mode": "hide"},
    )
    return fig


def histogram(
    data: pd.DataFrame,
    column: str,
    title: str,
    x_title: str,
    bins: int = 22,
    color: str = PALETTE[0],
    height: int = 430,
) -> go.Figure:
    series = data[column].dropna()
    fig = go.Figure()
    fig.add_trace(
        go.Histogram(
            x=series,
            nbinsx=bins,
            marker={"color": color, "line": {"color": "#ffffff", "width": 0.8}},
            opacity=0.9,
            hovertemplate=f"{x_title}: %{{x}}<br>{t('frequency')}: %{{y}}<extra></extra>",
        )
    )
    if len(series) > 0:
        fig.add_vline(
            x=series.median(),
            line_width=2,
            line_dash="dash",
            line_color="#111827",
            annotation_text=t("median"),
            annotation_position="top right",
        )
    return apply_journal_layout(fig, title, x_title, t("frequency"), height)


def boxplot(
    data: pd.DataFrame,
    y_col: str,
    title: str,
    y_title: str,
    x_col: str | None = None,
    height: int = 430,
) -> go.Figure:
    fig = go.Figure()
    if x_col:
        categories = [item for item in data[x_col].dropna().unique()]
        for i, category in enumerate(categories):
            values = data.loc[data[x_col] == category, y_col].dropna()
            display_category = category_label(category)
            fig.add_trace(
                go.Box(
                    y=values,
                    name=display_category,
                    marker_color=PALETTE[i % len(PALETTE)],
                    boxmean=True,
                    jitter=0.25,
                    pointpos=-1.6,
                    boxpoints="outliers",
                    hovertemplate=f"{display_category}<br>{y_title}: %{{y}}<extra></extra>",
                )
            )
        return apply_journal_layout(fig, title, variable_label(x_col), y_title, height, showlegend=False)

    fig.add_trace(
        go.Box(
            y=data[y_col].dropna(),
            name=y_title,
            marker_color=PALETTE[0],
            boxmean=True,
            jitter=0.25,
            pointpos=-1.6,
            boxpoints="outliers",
            hovertemplate=f"{y_title}: %{{y}}<extra></extra>",
        )
    )
    return apply_journal_layout(fig, title, "", y_title, height)


def multi_boxplot(
    data: pd.DataFrame,
    columns: list[str],
    labels: list[str],
    title: str,
    y_title: str,
    height: int = 430,
) -> go.Figure:
    fig = go.Figure()
    for i, (column, label) in enumerate(zip(columns, labels)):
        fig.add_trace(
            go.Box(
                y=data[column].dropna(),
                name=label,
                marker_color=PALETTE[i % len(PALETTE)],
                boxmean=True,
                boxpoints="outliers",
                hovertemplate=f"{label}<br>{y_title}: %{{y}}<extra></extra>",
            )
        )
    return apply_journal_layout(fig, title, "", y_title, height)


def mean_bar(
    data: pd.DataFrame,
    columns: list[str],
    labels: list[str],
    title: str,
    y_title: str,
    height: int = 430,
) -> go.Figure:
    values = [data[column].mean() for column in columns]
    fig = go.Figure(
        go.Bar(
            x=labels,
            y=values,
            marker={
                "color": [PALETTE[i % len(PALETTE)] for i in range(len(labels))],
                "line": {"color": "#111827", "width": 0.4},
            },
            text=[format_decimal(value, 2) for value in values],
            textposition="outside",
            cliponaxis=False,
            hovertemplate=f"%{{x}}<br>{t('mean')}: %{{y:.2f}}<extra></extra>",
        )
    )
    return apply_journal_layout(fig, title, "", y_title, height)


def numeric_summary(data: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    summary = (
        data[columns]
        .agg(["count", "mean", "std", "median", "min", "max"])
        .T.rename(
            columns={
                "count": "N",
                "mean": t("mean"),
                "std": t("std"),
                "median": t("median"),
                "min": t("minimum"),
                "max": t("maximum"),
            }
        )
    )
    summary["N"] = summary["N"].astype(int)
    summary.index = [variable_label(column) for column in columns]
    summary.index.name = t("variable")
    return summary


def show_plot(fig: go.Figure) -> None:
    st.plotly_chart(fig, width="stretch", config=PLOT_CONFIG)


def category_insight_text(
    data: pd.DataFrame,
    column: str,
    top_n: int | None = None,
    order: list[str] | None = None,
) -> str:
    chart = prepare_category_counts(data, column, top_n=top_n, order=order)
    if chart.empty:
        return ""
    row = chart.sort_values("Jumlah", ascending=False).iloc[0]
    return t("categorical_insight").format(
        category=row["Kategori"],
        count=format_integer(row["Jumlah"]),
        percent=format_percent(row["Persentase"]),
    )


def render_category_profile(
    data: pd.DataFrame,
    column: str,
    chart_type: str = "bar",
    top_n: int | None = None,
    orientation: str = "v",
    order: list[str] | None = None,
    height: int = 420,
) -> None:
    label = variable_label(column)
    title = t("composition_of").format(variable=label)
    if chart_type == "donut":
        show_plot(donut_chart(data, column, title, order=order, height=height, label=label))
    else:
        show_plot(
            count_bar(
                data,
                column,
                title,
                top_n=top_n,
                orientation=orientation,
                order=order,
                height=height,
                label=label,
            )
        )
    insight = category_insight_text(data, column, top_n=top_n, order=order)
    if insight:
        make_insight(insight)
    st.dataframe(
        category_summary(data, column, top_n=top_n, order=order),
        width="stretch",
        hide_index=True,
    )


def format_measure(value: float, digits: int = 2, unit_id: str = "", unit_en: str = "") -> str:
    if pd.isna(value):
        return "-"
    base = format_integer(value) if digits == 0 else format_decimal(value, digits)
    unit = unit_en if APP_LANG == "en" else unit_id
    return f"{base} {unit}".strip()


def render_numeric_profile(
    data: pd.DataFrame,
    column: str,
    color: str,
    digits: int = 2,
    unit_id: str = "",
    unit_en: str = "",
    bins: int = 22,
) -> None:
    label = variable_label(column)
    series = data[column].dropna()

    st.subheader(label)
    if series.empty:
        st.warning(t("no_data"))
        return

    stat_cols = st.columns(6)
    stat_cols[0].metric("N", format_integer(series.count()))
    stat_cols[1].metric(t("mean"), format_measure(series.mean(), digits, unit_id, unit_en))
    stat_cols[2].metric(t("median"), format_measure(series.median(), digits, unit_id, unit_en))
    stat_cols[3].metric(t("std"), format_measure(series.std(), digits, unit_id, unit_en))
    stat_cols[4].metric(t("minimum"), format_measure(series.min(), digits, unit_id, unit_en))
    stat_cols[5].metric(t("maximum"), format_measure(series.max(), digits, unit_id, unit_en))

    make_insight(
        t("numeric_insight").format(
            variable=label,
            median=format_measure(series.median(), digits, unit_id, unit_en),
            minimum=format_measure(series.min(), digits, unit_id, unit_en),
            maximum=format_measure(series.max(), digits, unit_id, unit_en),
        ),
        label=t("descriptive_stats_for").format(variable=label),
    )

    left, right = st.columns(2)
    with left:
        show_plot(
            histogram(
                data,
                column,
                t("distribution_of").format(variable=label),
                label,
                bins=bins,
                color=color,
                height=390,
            )
        )
    with right:
        show_plot(
            boxplot(
                data,
                column,
                t("boxplot_of").format(variable=label),
                label,
                height=390,
            )
        )


with st.sidebar:
    language_choice = st.radio(
        t("language"),
        list(LANGUAGE_OPTIONS),
        horizontal=True,
        index=list(LANGUAGE_OPTIONS).index(st.session_state.get("language_choice", "Indonesia")),
        key="language_choice",
    )
    APP_LANG = language_code(language_choice)

    st.header(t("sidebar_title"))

    st.subheader(t("data_source"))
    data_source_choice = st.radio(
        t("data_source"),
        [DATA_SOURCE_EXISTING, DATA_SOURCE_UPLOAD],
        format_func=lambda option: (
            t("use_available_data") if option == DATA_SOURCE_EXISTING else t("upload_data")
        ),
        index=0,
        label_visibility="collapsed",
        key="data_source_choice",
    )
    uploaded_file = None
    if data_source_choice == DATA_SOURCE_UPLOAD:
        uploaded_file = st.file_uploader(
            t("upload_file"),
            type=UPLOAD_FILE_TYPES,
            help=t("upload_help"),
        )

    if st.button(t("reload_data"), use_container_width=True):
        load_existing_data.clear()
        load_uploaded_data.clear()
        st.rerun()

    try:
        if data_source_choice == DATA_SOURCE_UPLOAD:
            if uploaded_file is None:
                st.info(t("upload_missing"))
                st.stop()
            data = load_uploaded_data(uploaded_file.name, uploaded_file.getvalue())
        else:
            if not DATA_FILE.exists():
                st.error(t("file_not_found").format(file_name=DATA_FILE))
                st.stop()
            data_signature = DATA_FILE.stat().st_mtime_ns
            data = load_existing_data(str(DATA_FILE), data_signature)
    except Exception as error:
        st.error(t("data_load_error").format(error=error))
        st.stop()

    gender_options = sorted(data["Jenis Kelamin"].dropna().unique(), key=category_label)
    education_options = sorted(data["Pendidikan Terakhir"].dropna().unique(), key=category_label)
    land_options = sorted(data["Status Lahan"].dropna().unique(), key=category_label)
    training_options = sorted(data["Status Pelatihan Sawit"].dropna().unique(), key=category_label)

    with st.expander(t("filters"), expanded=True):
        gender_filter = st.multiselect(
            t("gender"),
            gender_options,
            default=gender_options,
            format_func=category_label,
        )
        education_filter = st.multiselect(
            t("education"),
            education_options,
            default=education_options,
            format_func=category_label,
        )
        land_filter = st.multiselect(
            t("land_status"),
            land_options,
            default=land_options,
            format_func=category_label,
        )
        training_filter = st.multiselect(
            t("training_status"),
            training_options,
            default=training_options,
            format_func=category_label,
        )

        ages = data["Umur (tahun)"].dropna()
        if ages.empty:
            st.warning(t("invalid_age_filter"))
            st.stop()
        age_min = int(ages.min())
        age_max = int(ages.max())
        age_filter = st.slider(t("age_range"), age_min, age_max, (age_min, age_max))


filtered = data[
    data["Jenis Kelamin"].isin(gender_filter)
    & data["Pendidikan Terakhir"].isin(education_filter)
    & data["Status Lahan"].isin(land_filter)
    & data["Status Pelatihan Sawit"].isin(training_filter)
    & data["Umur (tahun)"].between(age_filter[0], age_filter[1], inclusive="both")
].copy()

if filtered.empty:
    st.warning(t("no_data"))
    st.stop()


numeric_profiles = [
    {
        "column": "Umur (tahun)",
        "color": PALETTE[0],
        "digits": 1,
        "unit_id": "tahun",
        "unit_en": "years",
        "bins": 20,
    },
    {
        "column": "Jumlah Anggota Keluarga",
        "color": PALETTE[1],
        "digits": 0,
        "unit_id": "orang",
        "unit_en": "people",
        "bins": 12,
    },
    {
        "column": "Penghasilan (juta Rupiah/bulan)",
        "color": PALETTE[2],
        "digits": 2,
        "unit_id": "juta Rp/bln",
        "unit_en": "million Rp/month",
        "bins": 22,
    },
    {
        "column": "Pengeluaran (juta Rupiah/bulan)",
        "color": PALETTE[3],
        "digits": 2,
        "unit_id": "juta Rp/bln",
        "unit_en": "million Rp/month",
        "bins": 22,
    },
    {
        "column": "Luas Lahan Sawit (Hektar)",
        "color": PALETTE[5],
        "digits": 2,
        "unit_id": "ha",
        "unit_en": "ha",
        "bins": 18,
    },
    {
        "column": "Jumlah Pohon Sawit (pohon)",
        "color": PALETTE[6],
        "digits": 0,
        "unit_id": "pohon",
        "unit_en": "trees",
        "bins": 20,
    },
    {
        "column": "Produksi Rata-rata (ton/bulan)",
        "color": PALETTE[7],
        "digits": 2,
        "unit_id": "ton/bln",
        "unit_en": "tons/month",
        "bins": 20,
    },
    {
        "column": "Lama menjadi Petani Sawit (tahun)",
        "color": PALETTE[8],
        "digits": 1,
        "unit_id": "tahun",
        "unit_en": "years",
        "bins": 18,
    },
]

numeric_columns = [profile["column"] for profile in numeric_profiles]

public_columns = [
    "No",
    "Jenis Kelamin",
    "Wilayah",
    "Umur (tahun)",
    "Jumlah Anggota Keluarga",
    "Pekerjaan Utama Pencari Nafkah",
    "Pendidikan Terakhir",
    "Penghasilan Keluarga per Bulan (Rupiah)",
    "Pengeluaran Keluarga per Bulan (Rupiah)",
    "Luas Lahan Sawit (Hektar)",
    "Status Lahan",
    "Jumlah Pohon Sawit (pohon)",
    "Produksi Rata-rata (ton/bulan)",
    "Lama menjadi Petani Sawit (tahun)",
    "Jenis Pelatihan Sawit",
    "Status Pelatihan Sawit",
    "Surplus Pendapatan (Rupiah/bulan)",
]

categorical_columns = [
    "Jenis Kelamin",
    "Wilayah",
    "Pekerjaan Utama Pencari Nafkah",
    "Pendidikan Terakhir",
    "Status Lahan",
    "Jenis Pelatihan Sawit",
    "Status Pelatihan Sawit",
]

st.markdown(
    f"""
    <div class="hero-panel">
        <h1>{t("main_title")}</h1>
        <p class="small-note">{t("intro")}</p>
    </div>
    """,
    unsafe_allow_html=True,
)

gender_counts = prepare_category_counts(
    filtered,
    "Jenis Kelamin",
    order=["Laki-laki", "Perempuan"],
)
dominant_gender = gender_counts.sort_values("Jumlah", ascending=False).iloc[0]

metric_cols = st.columns(6)
metric_cols[0].metric(t("respondents"), format_integer(len(filtered)))
metric_cols[1].metric(
    t("avg_age"),
    format_measure(filtered["Umur (tahun)"].mean(), 1, "tahun", "years"),
)
metric_cols[2].metric(
    t("majority_gender"),
    dominant_gender["Kategori"],
    format_percent(dominant_gender["Persentase"]),
)
metric_cols[3].metric(
    t("median_family"),
    format_measure(filtered["Jumlah Anggota Keluarga"].median(), 0, "orang", "people"),
)
metric_cols[4].metric(
    t("median_land"),
    format_measure(filtered["Luas Lahan Sawit (Hektar)"].median(), 2, "ha", "ha"),
)
metric_cols[5].metric(
    t("median_production"),
    format_measure(filtered["Produksi Rata-rata (ton/bulan)"].median(), 2, "ton/bln", "tons/month"),
)

tab_overview, tab_category, tab_numeric, tab_data = st.tabs(
    [
        t("overview"),
        t("categorical"),
        t("numeric"),
        t("processed_data"),
    ]
)


with tab_overview:
    st.subheader(t("profile_snapshot"))
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        show_plot(
            donut_chart(
                filtered,
                "Jenis Kelamin",
                t("gender_pie"),
                order=["Laki-laki", "Perempuan"],
                height=390,
                label=variable_label("Jenis Kelamin"),
            )
        )
    with col_b:
        show_plot(
            count_bar(
                filtered,
                "Wilayah",
                t("top_regions"),
                top_n=8,
                orientation="h",
                height=390,
                label=variable_label("Wilayah"),
            )
        )
    with col_c:
        show_plot(
            count_bar(
                filtered,
                "Pendidikan Terakhir",
                t("education_profile"),
                order=[
                    "Tidak Sekolah",
                    "SD",
                    "SMP",
                    "SMA",
                    "Diploma",
                    "Sarjana",
                    "Diploma/Sarjana",
                    "Lainnya",
                ],
                height=390,
                label=variable_label("Pendidikan Terakhir"),
            )
        )

    st.subheader(t("numeric_summary"))
    summary_table = numeric_summary(filtered, numeric_columns)
    st.dataframe(
        summary_table.style.format(
            {
                "N": lambda value: format_integer(value),
                t("mean"): lambda value: format_decimal(value, 2),
                t("std"): lambda value: format_decimal(value, 2),
                t("median"): lambda value: format_decimal(value, 2),
                t("minimum"): lambda value: format_decimal(value, 2),
                t("maximum"): lambda value: format_decimal(value, 2),
            }
        ),
        width="stretch",
    )

    st.subheader(t("data_completeness"))
    missing = filtered[public_columns[1:]].isna().sum().rename(t("missing_count")).to_frame()
    missing[t("missing_percent")] = missing[t("missing_count")] / len(filtered) * 100
    missing.index = [variable_label(column) for column in public_columns[1:]]
    missing.index.name = t("variable")
    st.dataframe(
        missing.style.format(
            {
                t("missing_count"): lambda value: format_integer(value),
                t("missing_percent"): lambda value: format_percent(value, 2),
            }
        ),
        width="stretch",
    )


with tab_category:
    left, right = st.columns(2)
    with left:
        render_category_profile(
            filtered,
            "Jenis Kelamin",
            chart_type="donut",
            order=["Laki-laki", "Perempuan"],
        )
        render_category_profile(
            filtered,
            "Pekerjaan Utama Pencari Nafkah",
            top_n=8,
            orientation="h",
        )
        render_category_profile(
            filtered,
            "Status Lahan",
            top_n=8,
            orientation="h",
        )
        render_category_profile(
            filtered,
            "Status Pelatihan Sawit",
            order=["Pernah Mengikuti", "Belum Pernah", "Tidak Diisi"],
        )

    with right:
        render_category_profile(
            filtered,
            "Pendidikan Terakhir",
            order=[
                "Tidak Sekolah",
                "SD",
                "SMP",
                "SMA",
                "Diploma",
                "Sarjana",
                "Diploma/Sarjana",
                "Lainnya",
            ],
        )
        render_category_profile(
            filtered,
            "Wilayah",
            top_n=12,
            orientation="h",
            height=460,
        )
        render_category_profile(
            filtered,
            "Jenis Pelatihan Sawit",
            top_n=10,
            orientation="h",
            height=460,
        )


with tab_numeric:
    st.subheader(t("numeric_summary"))
    summary_table = numeric_summary(filtered, numeric_columns)
    st.dataframe(
        summary_table.style.format(
            {
                "N": lambda value: format_integer(value),
                t("mean"): lambda value: format_decimal(value, 2),
                t("std"): lambda value: format_decimal(value, 2),
                t("median"): lambda value: format_decimal(value, 2),
                t("minimum"): lambda value: format_decimal(value, 2),
                t("maximum"): lambda value: format_decimal(value, 2),
            }
        ),
        width="stretch",
    )

    for profile in numeric_profiles:
        render_numeric_profile(
            filtered,
            profile["column"],
            profile["color"],
            digits=profile["digits"],
            unit_id=profile["unit_id"],
            unit_en=profile["unit_en"],
            bins=profile["bins"],
        )


with tab_data:
    st.subheader(t("sanitized_dataset"))
    display_data = filtered[public_columns].copy()
    for column in categorical_columns:
        display_data[column] = display_data[column].map(category_label)
    display_data = display_data.rename(columns={column: variable_label(column) for column in public_columns})

    st.dataframe(display_data, width="stretch", height=560)
    csv_data = display_data.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        t("download_csv"),
        data=csv_data,
        file_name=t("download_file_name"),
        mime="text/csv",
    )
