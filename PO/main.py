"""
main.py
-------
Orquestrador do pipeline de processamento de Pedidos de Compra.

Pipeline:
    1. extractor_msg        → extrai PDFs dos emails .msg
    2. extractor_ids        → mapeia ID → Nome de produto
    3. extractor_quantidades → consolida quantidades e gera Excel
"""

from __future__ import annotations

import logging
import sys
import time
from pathlib import Path

# ─── Configuração de logging ─────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(Path("output") / "pipeline.log", encoding="utf-8"),
    ],
)
logger = logging.getLogger(__name__)


def _banner(title: str) -> None:
    print(f"\n{'─' * 60}")
    print(f"  {title}")
    print(f"{'─' * 60}")


def _elapsed(start: float) -> str:
    return f"{time.time() - start:.1f}s"


# ─── Pipeline ────────────────────────────────────────────────────────────────

def run_pipeline() -> None:
    total_start = time.time()

    # ── Etapa 1: Extração de anexos PDF dos .msg ─────────────────────────────
    _banner("🚀  ETAPA 1 — Extração de PDFs dos emails (.msg)")
    t0 = time.time()
    try:
        from extractor_msg import extract_pdf_attachments

        pdfs = extract_pdf_attachments()
        print(f"\n✅  Etapa 1 concluída em {_elapsed(t0)} — {len(pdfs)} PDF(s) extraído(s).")
    except Exception as exc:
        logger.exception("Falha na Etapa 1: %s", exc)
        print(f"\n❌  Etapa 1 falhou: {exc}")
        sys.exit(1)

    # ── Etapa 2: Mapeamento de IDs e Nomes de produto ────────────────────────
    _banner("🔍  ETAPA 2 — Mapeamento de produtos (ID → Nome)")
    t0 = time.time()
    try:
        from extractor_ids import build_product_mapping

        mapping = build_product_mapping()
        print(
            f"\n✅  Etapa 2 concluída em {_elapsed(t0)} — "
            f"{len(mapping)} produto(s) mapeado(s)."
        )
    except Exception as exc:
        logger.exception("Falha na Etapa 2: %s", exc)
        print(f"\n❌  Etapa 2 falhou: {exc}")
        sys.exit(1)

    # ── Etapa 3: Consolidação de quantidades e geração do Excel ──────────────
    _banner("📊  ETAPA 3 — Consolidação de quantidades e geração do Excel")
    t0 = time.time()
    try:
        from extractor_quantidades import build_consolidation_report

        build_consolidation_report()
        print(f"\n✅  Etapa 3 concluída em {_elapsed(t0)}.")
    except Exception as exc:
        logger.exception("Falha na Etapa 3: %s", exc)
        print(f"\n❌  Etapa 3 falhou: {exc}")
        sys.exit(1)

    # ── Resumo final ──────────────────────────────────────────────────────────
    _banner(f"🏁  PIPELINE CONCLUÍDO em {_elapsed(total_start)}")
    print("  📁  Saídas geradas em output/")
    print("       ├─ mapeamento_produtos.csv")
    print("       ├─ consolidado_pedidos.xlsx")
    print("       └─ pipeline.log")
    print()

    # ── Arquivamento opcional ─────────────────────────────────────────────────
    resposta = input("  Deseja arquivar este lote agora? (s/N): ").strip().lower()
    if resposta == "s":
        from archiver import archive
        archive()


if __name__ == "__main__":
    # Garante que a pasta de log exista antes de o FileHandler tentar abri-la
    Path("output").mkdir(parents=True, exist_ok=True)
    run_pipeline()
