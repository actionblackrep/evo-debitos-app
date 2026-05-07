#!/usr/bin/env bash
# Reporte de Debitos - macOS / Linux
set -e
cd "$(dirname "$0")"

if ! command -v python3 >/dev/null 2>&1; then
  echo "Python 3 no esta instalado. Descarga desde https://www.python.org/downloads/"
  exit 1
fi

if [ ! -d ".venv" ]; then
  echo "[setup] Creando entorno virtual..."
  python3 -m venv .venv
fi

# shellcheck disable=SC1091
source .venv/bin/activate

echo "[setup] Instalando dependencias..."
pip install --upgrade pip >/dev/null
pip install -r requirements.txt

echo "[run] Ejecutando agente..."
python generate_report.py "$@"

echo
echo "Listo. PDF disponible en ./output/"
