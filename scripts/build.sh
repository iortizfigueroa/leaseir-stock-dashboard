#!/bin/bash
# build.sh — combina src/ + data/master + data/ejercicios en un workdir plano,
# corre el script Python, y mueve el HTML a docs/
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
WORK="$ROOT/.build"

echo "[build] Preparando workspace en $WORK"
rm -rf "$WORK"
mkdir -p "$WORK"

# Copiar fuentes
cp "$ROOT/src/"*.py "$WORK/"
cp "$ROOT/src/"*.html "$WORK/"

# Copiar maestros (estáticos, en repo)
cp "$ROOT/data/master/"*.xlsx "$WORK/"

# Copiar ejercicios diarios (los más recientes descargados de Drive)
if compgen -G "$ROOT/data/ejercicios/*.xlsx" > /dev/null; then
  cp "$ROOT/data/ejercicios/"*.xlsx "$WORK/"
else
  echo "[build] WARNING: no hay ejercicios en data/ejercicios/"
fi

echo "[build] Ejecutando vision_html_full.py"
cd "$WORK"
python3 vision_html_full.py

echo "[build] Moviendo output a docs/"
mkdir -p "$ROOT/docs"
cp "$WORK/Vision_Stock_Mes.html" "$ROOT/docs/index.html"
cp "$WORK/Vision_Stock_Mes.html" "$ROOT/docs/Vision_Stock_Mes.html"

# Limpieza
rm -rf "$WORK"

echo "[build] OK"
