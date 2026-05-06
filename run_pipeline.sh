#!/usr/bin/env bash
# ─────────────────────────────────────────────────────────────────────────────
# Pipeline: raw → output_final/painel_fomento_audiovisual.html
# ─────────────────────────────────────────────────────────────────────────────
set -e
SCRIPTS="$(cd "$(dirname "$0")/scripts" && pwd)"

echo "=========================================="
echo "Fomento Audiovisual — Pipeline"
echo "=========================================="

echo ""
echo "[1/7] Gerando tabela consolidada..."
python "$SCRIPTS/01_gerar_tabela_consolidada.py"

echo ""
echo "[2/7] Gerando datasets analíticos..."
python "$SCRIPTS/02_gerar_datasets.py"

echo ""
echo "[3/7] Painel — Critério de Seleção..."
python "$SCRIPTS/03_gerar_painel_criterio_selecao.py"

echo ""
echo "[4/7] Painel — Produtoras..."
python "$SCRIPTS/04_gerar_painel_produtoras.py"

echo ""
echo "[5/7] Painel — Concentração..."
python "$SCRIPTS/05_gerar_painel_concentracao.py"

echo ""
echo "[6/7] Painel — Comparativo de Mecanismos..."
python "$SCRIPTS/06_gerar_painel_comparativo.py"

echo ""
echo "[7/7] Painel final unificado..."
python "$SCRIPTS/07_gerar_painel_final.py"

echo ""
echo "=========================================="
echo "Pronto! Painel disponível em:"
echo "  output_final/painel_fomento_audiovisual.html"
echo "=========================================="
