"""
extractor_ids.py
----------------
Módulo responsável por ler o texto bruto de todos os PDFs em temp_pdfs/ e
construir um mapeamento único de ID → Nome de produto.

Estratégias de extração (aplicadas em ordem de prioridade):
  1. Bloco   – "ID do Produto: XXXXXXX    Descrição: ..." (cabeçalho das specs)
  2. Inline  – "XXXXXXX - NOME DO PRODUTO" (linha de abertura de cada spec)
  3. Tabela  – "XXXXXXX NOME DO PRODUTO" sem traço (tabela pág. 1 e última)

O resultado final (sem IDs duplicados) é salvo em output/mapeamento_produtos.csv.
"""

from __future__ import annotations

import csv
import logging
import re
from pathlib import Path

import fitz  # PyMuPDF
from tqdm import tqdm
from tqdm.contrib.logging import logging_redirect_tqdm

logger = logging.getLogger(__name__)

# ─── Caminhos padrão ────────────────────────────────────────────────────────
PDF_DIR = Path("temp_pdfs")
OUTPUT_DIR = Path("output")
OUTPUT_CSV = OUTPUT_DIR / "mapeamento_produtos.csv"

# ─── Expressões regulares ────────────────────────────────────────────────────

# Estratégia 1 – BLOCO (mais confiável): cabeçalho das Especificações Técnicas
#   "ID do Produto: 7004434     Descrição: SWITCH L2 - 24 PORTAS MULTGIGA"
RE_BLOCK = re.compile(
    r"ID\s+do\s+Produto[:\s]+(\d{7})[\s\S]{0,200}?Descri[çc][aã]o[:\s]+(.*)",
    re.IGNORECASE,
)

# Estratégia 2 – INLINE com traço: linha de abertura da spec de cada produto
#   "7004434 - SWITCH L2 - 24 PORTAS MULTGIGA"
RE_INLINE = re.compile(
    r"(?<!\d)(\d{7})(?!\d)\s*[-–]\s*([A-Z0-9][A-Z0-9 \.\-\+\/\(\)]+)",
    re.UNICODE,
)

# Estratégia 3 – TABELA PRINCIPAL: formato sem traço (1ª e última páginas do PDF)
#   "7004434 SWITCH L2 - 24 PORTAS MULTGIGA"
# Âncora de fim de linha (MULTILINE) evita capturar texto além do nome do produto.
# Mínimo de 5 chars no nome para não capturar ruídos curtos.
RE_TABLE = re.compile(
    r"(?<!\d)(\d{7})(?!\d)\s+([A-Z][A-Z0-9 \-\/\+\.\(\)]{4,60})$",
    re.MULTILINE | re.UNICODE,
)


def _clean_name(raw: str) -> str:
    """Remove quebras de linha e espaços duplos do nome do produto."""
    return re.sub(r"\s+", " ", raw.replace("\n", " ")).strip()


def _extract_text_from_pdf(pdf_path: Path) -> str:
    """Retorna o texto bruto concatenado de todas as páginas do PDF."""
    try:
        doc = fitz.open(str(pdf_path))
        pages_text = [page.get_text("text") for page in doc]
        doc.close()
        return "\n".join(pages_text)
    except Exception as exc:  # noqa: BLE001
        logger.error("Erro ao ler '%s': %s", pdf_path.name, exc)
        return ""


def _parse_products(text: str) -> dict[str, str]:
    """
    Aplica três estratégias de Regex e retorna {ID: Nome}.
    Prioridade: bloco > inline-com-traço > tabela-principal.
    """
    products: dict[str, str] = {}

    # Estratégia 1 – bloco: cabeçalho "ID do Produto: X  Descrição: Y"
    for match in RE_BLOCK.finditer(text):
        prod_id = match.group(1).strip()
        prod_name = _clean_name(match.group(2))
        if prod_id not in products and prod_name:
            products[prod_id] = prod_name

    # Estratégia 2 – inline com traço: "7004434 - SWITCH L2 - 24 PORTAS MULTGIGA"
    for match in RE_INLINE.finditer(text):
        prod_id = match.group(1).strip()
        prod_name = _clean_name(match.group(2))
        if prod_id not in products and prod_name:
            products[prod_id] = prod_name

    # Estratégia 3 – tabela principal (fallback): "7004434 SWITCH L2 - 24 PORTAS MULTGIGA"
    # Cobre PDFs que não tenham seção de Especificações Técnicas.
    for match in RE_TABLE.finditer(text):
        prod_id = match.group(1).strip()
        prod_name = _clean_name(match.group(2))
        if prod_id not in products and prod_name:
            products[prod_id] = prod_name

    return products


def build_product_mapping(
    pdf_dir: Path = PDF_DIR,
    output_csv: Path = OUTPUT_CSV,
) -> dict[str, str]:
    """
    Varre *pdf_dir*, extrai pares (ID, Nome) de cada PDF e persiste o
    mapeamento consolidado em *output_csv*.

    Returns
    -------
    dict[str, str]
        Mapeamento { ID_PRODUTO: NOME_PRODUTO } sem duplicatas.
    """
    output_csv.parent.mkdir(parents=True, exist_ok=True)

    pdf_files = sorted(pdf_dir.glob("*.pdf"))
    if not pdf_files:
        logger.warning("Nenhum PDF encontrado em '%s'.", pdf_dir)
        return {}

    all_products: dict[str, str] = {}

    with logging_redirect_tqdm():
        for pdf_path in tqdm(pdf_files, desc="  🔍  Mapeando produtos", unit="pdf", dynamic_ncols=True, leave=True):
            text = _extract_text_from_pdf(pdf_path)
            if not text:
                continue

            found = _parse_products(text)
            new_count = 0
            for pid, pname in found.items():
                if pid not in all_products:
                    all_products[pid] = pname
                    new_count += 1

            logger.info(
                "  ↳ '%s': %d produto(s) encontrado(s) (%d novo(s)).",
                pdf_path.name,
                len(found),
                new_count,
            )

    # Persistência
    with output_csv.open("w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.writer(fh)
        writer.writerow(["ID_PRODUTO", "NOME_PRODUTO"])
        for pid, pname in sorted(all_products.items()):
            writer.writerow([pid, pname])

    logger.info("Mapeamento salvo: '%s' (%d produto(s)).", output_csv, len(all_products))
    return all_products


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    mapping = build_product_mapping()
    print(f"\n✅  {len(mapping)} produto(s) únicos mapeados em '{OUTPUT_CSV}'.")
