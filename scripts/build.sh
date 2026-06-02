#!/bin/bash
# build.sh — combina src/ + data/master + data/ejercicios en un workdir plano,
# corre el script Python, y mueve el HTML a docs/
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
WORK="$ROOT/.build"

echo "[build] Preparando workspace en $WORK"
rm -rf "$WORK"
mkdir -p "$WORK/SPECS"

# Copiar fuentes
cp "$ROOT/src/"*.py "$WORK/"
cp "$ROOT/src/"*.html "$WORK/"

# Copiar maestros root (Inventario_MASTER, escandallos consolidados, etc.)
cp "$ROOT/data/master/"*.xlsx "$WORK/" 2>/dev/null || true

# Copiar SPECS/ (escandallos individuales, formato por SPEC)
if compgen -G "$ROOT/data/master/SPECS/*.xlsx" > /dev/null; then
  cp "$ROOT/data/master/SPECS/"*.xlsx "$WORK/SPECS/"
    echo "[build] Copiados $(ls $WORK/SPECS/ | wc -l) escandallos a SPECS/"
    else
      echo "[build] WARNING: no hay xlsx en data/master/SPECS/ — escandallos no se cargarán"
      fi

      # Copiar ejercicios diarios
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

          rm -rf "$WORK"
          echo "[build] OK"
          
