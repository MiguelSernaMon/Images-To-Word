#!/bin/zsh
cd "$(dirname "$0")"
echo "🚀 Iniciando servidor..."
echo "📍 Abre tu navegador en: http://localhost:5001"
echo ""
echo "Para cerrar, presiona Ctrl+C o cierra esta ventana"
echo "=================================================="
open "http://localhost:5001"
".venv/bin/python" app.py
