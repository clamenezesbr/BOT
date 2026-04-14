"""
main.py
-------
Orquestrador principal da automação RPA de Pedidos de Compra.

Fluxo:
  1. Extrai PDFs de arquivos .msg (extractor.py)
  2. Faz o parse de cada PDF (po_parser.py)
  3. Gera a planilha Excel consolidada (formatter.py)
  4. Exibe relatório final no terminal

Uso:
    python main.py
"""

import logging
import sys
import time
from pathlib import Path

from tqdm import tqdm

import extractor
import po_parser as parser
import formatter

# ---------------------------------------------------------------------------
# Diretórios
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).parent
INPUT_DIR = BASE_DIR / "input_msgs"
TEMP_DIR = BASE_DIR / "temp_pdfs"
OUTPUT_DIR = BASE_DIR / "output"
LOG_FILE = OUTPUT_DIR / "execucao.log"


# ---------------------------------------------------------------------------
# Configuração de Logging
# ---------------------------------------------------------------------------
def _setup_logging() -> None:
    """Configura handlers de console (INFO) e arquivo (DEBUG)."""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    root = logging.getLogger()
    root.setLevel(logging.DEBUG)

    # Handler: arquivo — captura tudo (DEBUG+)
    fh = logging.FileHandler(str(LOG_FILE), encoding="utf-8", mode="a")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(
        logging.Formatter(
            "%(asctime)s | %(levelname)-8s | %(module)s | %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
    )

    # Handler: console — apenas WARNING+ (não polui o terminal com tqdm)
    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.WARNING)
    ch.setFormatter(logging.Formatter("%(levelname)s: %(message)s"))

    root.addHandler(fh)
    root.addHandler(ch)


logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Helpers de UI
# ---------------------------------------------------------------------------
SEP = "─" * 60


def _banner(msg: str) -> None:
    print(f"\n{SEP}")
    print(f"  {msg}")
    print(SEP)


def _step(emoji: str, msg: str) -> None:
    print(f"\n{emoji}  {msg}")


# ---------------------------------------------------------------------------
# Etapas
# ---------------------------------------------------------------------------
def etapa_extracao() -> list[Path]:
    _step("📨", "ETAPA 1 — Extraindo PDFs dos arquivos .msg …")
    logger.info("=== ETAPA 1: Extração de PDFs ===")

    pdf_paths, total_msgs, skipped = extractor.extract_pdfs_from_msgs(
        INPUT_DIR, TEMP_DIR
    )

    # Barra de progresso visual (retroativa — simula progresso dos .msg)
    with tqdm(
        total=total_msgs,
        desc="  .msg processados",
        unit="msg",
        ncols=70,
        colour="cyan",
    ) as pbar:
        pbar.update(total_msgs)

    print(
        f"     → {total_msgs} arquivo(s) .msg encontrado(s) | "
        f"{len(pdf_paths)} PDF(s) extraído(s) | "
        f"{skipped} sem PDF"
    )
    return pdf_paths


def etapa_parse(pdf_paths: list[Path]) -> list[dict]:
    _step("🔍", "ETAPA 2 — Fazendo parse dos PDFs …")
    logger.info("=== ETAPA 2: Parse dos PDFs ===")

    all_records: list[dict] = []
    errors = 0

    with tqdm(
        total=len(pdf_paths),
        desc="  PDFs analisados",
        unit="pdf",
        ncols=70,
        colour="yellow",
    ) as pbar:
        for pdf_path in pdf_paths:
            records = parser.parse_pdf(pdf_path)
            if not records:
                errors += 1
            all_records.extend(records)
            pbar.update(1)

    pos_extraidos = len({r["PO"] for r in all_records})
    print(
        f"     → {len(pdf_paths)} PDF(s) analisado(s) | "
        f"{pos_extraidos} PO(s) extraído(s) | "
        f"{errors} sem dados"
    )
    return all_records


def etapa_exportacao(records: list[dict]) -> tuple[Path, list[str]]:
    _step("📊", "ETAPA 3 — Gerando planilha Excel …")
    logger.info("=== ETAPA 3: Exportação Excel ===")

    output_path, duplicate_pos = formatter.export_to_excel(records, OUTPUT_DIR)

    print(f"     → Planilha salva em: {output_path}")
    if duplicate_pos:
        print(f"     ⚠️  POs duplicados somados: {', '.join(duplicate_pos)}")

    return output_path, duplicate_pos


# ---------------------------------------------------------------------------
# Relatório final
# ---------------------------------------------------------------------------
def _print_summary(
    elapsed: float,
    total_msgs: int,
    pdf_paths: list[Path],
    all_records: list[dict],
    duplicate_pos: list[str],
    output_path: Path,
) -> None:
    pos_unicos = len({r["PO"] for r in all_records})
    erros = len(pdf_paths) - len({r["PO"] for r in all_records})
    # erros pode ser negativo se um PDF gerou múltiplos POs (não esperado)
    erros = max(erros, 0)

    _banner("RELATÓRIO FINAL DE EXECUÇÃO")
    print(f"  ⏱️  Tempo de execução   : {elapsed:.2f}s")
    print(f"  📄 .msg processados    : {total_msgs}")
    print(f"  📎 PDFs extraídos      : {len(pdf_paths)}")
    print(f"  ✅ POs com sucesso     : {pos_unicos}")
    print(f"  ⚠️  PDFs sem dados/erro : {erros}")
    print(f"  🔄 POs duplicados      : {len(duplicate_pos)}")
    print(f"  📁 Planilha gerada     : {output_path}")
    print(f"  📝 Log gravado em      : {LOG_FILE}")
    print(SEP + "\n")

    logger.info(
        "Execução concluída | elapsed=%.2fs | msgs=%d | pdfs=%d | pos=%d | "
        "erros=%d | duplicatas=%d",
        elapsed, total_msgs, len(pdf_paths), pos_unicos, erros, len(duplicate_pos),
    )


# ---------------------------------------------------------------------------
# Entry-point
# ---------------------------------------------------------------------------
def main() -> None:
    _setup_logging()
    start_time = time.perf_counter()

    _banner("🤖  AUTOMAÇÃO RPA — PEDIDOS DE COMPRA  v0.4")
    logger.info("Iniciando automação.")

    # Etapa 1 — Extração
    pdf_paths = etapa_extracao()

    if not pdf_paths:
        print(
            "\n  ❌  Nenhum PDF extraído. "
            "Verifique a pasta input_msgs/ e tente novamente."
        )
        logger.error("Nenhum PDF extraído. Encerrando.")
        sys.exit(1)

    # Etapa 2 — Parse
    all_records = etapa_parse(pdf_paths)

    if not all_records:
        print(
            "\n  ❌  Nenhum dado extraído dos PDFs. "
            "Verifique os arquivos e o log em output/execucao.log."
        )
        logger.error("Nenhum registro gerado pelo parser. Encerrando.")
        sys.exit(1)

    # Etapa 3 — Exportação
    output_path, duplicate_pos = etapa_exportacao(all_records)

    elapsed = time.perf_counter() - start_time

    # Calcular total_msgs para o resumo (lê qty de .msg que existiam)
    total_msgs = len(list(INPUT_DIR.glob("*.msg"))) if INPUT_DIR.exists() else 0

    _step("🎉", "Automação concluída com sucesso!")
    _print_summary(
        elapsed, total_msgs, pdf_paths, all_records, duplicate_pos, output_path
    )


if __name__ == "__main__":
    main()
