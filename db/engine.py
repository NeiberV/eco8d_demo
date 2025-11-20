from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
import os
from pathlib import Path
import tomli

# Carpeta raíz del proyecto
BASE_DIR = Path(__file__).resolve().parents[1]

# Cargar settings TOML
def _load_settings():
    cfg_path = BASE_DIR / "config" / "settings.toml"
    with open(cfg_path, "rb") as f:
        return tomli.load(f)

_settings = _load_settings()

# Si el TOML trae ruta completa → úsala
# Si trae solo "sqlite:///eco8d.sqlite3" → forzar a almacenarlo en BASE_DIR
db_url = _settings.get("database", {}).get("url", "").strip()

if db_url.startswith("sqlite:///") and db_url.count("/") <= 3:
    # Extrae nombre del archivo
    filename = db_url.replace("sqlite:///", "")
    # Forzar ubicación en la raíz del proyecto
    db_url = f"sqlite:///{BASE_DIR / filename}"

DB_URL = db_url

engine = create_engine(DB_URL, echo=False, future=True)
SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False, future=True)

# Exportar rutas clave
EXCEL_MASTER_PATH = BASE_DIR / "BASE DE DATOS GENERAL.xlsx"
CARPETA_INFORMES_8D = BASE_DIR / "informes_8d"
CARPETA_INFORMES_8D.mkdir(exist_ok=True)