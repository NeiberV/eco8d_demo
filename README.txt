# 1) Abrir PowerShell en:
# C:\Users\Administrador\Desktop\Sistema de Gestion Operativa de calidad (control y seguimiento de incidencias)

python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt

# 2) Crear BD (SQLite)
python -m db.create_db

# 3) Validar Excel (sin insertar)
python -m ingest.run_validate

# 4) Ingerir a BD
python -m ingest.excel_ingest

# 5) Ver archivo eco8d.sqlite3 creado en la misma carpeta