#!/usr/bin/bash
# Script de ejecución del microservicio ExcelTemplates (v12)
# Uso: scripts/run.sh [puerto]   ·   Variables: PORT, HOST, ETB_DATA_DIR, CORS_ORIGINS, TRUST_PROXY
set -euo pipefail
cd "$(dirname "$0")/.."
node_path="$(command -v node 2>/dev/null || true)"
if [[ -z "$node_path" ]]; then
    node_path="$(find "$HOME/.nvm" -maxdepth 5 -type f -path "*/bin/node" -executable 2>/dev/null | head -n 1 || true)"
fi
if [[ -z "$node_path" ]]; then
    node_path="$(find /usr/local/bin /usr/bin /bin /opt -maxdepth 3 -type f -name node -executable 2>/dev/null | head -n 1 || true)"
fi
[[ -z "$node_path" ]] && { echo "Node.js no encontrado (se requiere Node 18 o superior)"; exit 1; }
exec "$node_path" --max-http-header-size=1073741824 main.js "${1:-}"
