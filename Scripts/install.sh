#!/usr/bin/bash
# Script de instalación del microservicio ExcelTemplates (v12)
set -euo pipefail
cd "$(dirname "$0")/.."
echo 'Instalando dependencias del microservicio ExcelTemplates'
if [[ -f package-lock.json ]]; then
    npm ci --omit=dev
else
    npm install --omit=dev
fi
