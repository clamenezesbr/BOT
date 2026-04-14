"""
parser.py
---------
Módulo responsável por:
  - Extrair o texto bruto de cada PDF em temp_pdfs/
  - Capturar o número do Pedido de Compra (PO) via regex
  - Normalizar o texto (maiúsculas, sem quebras/espaços duplos)
  - Localizar cada produto da lista estrita e capturar sua quantidade (inteiro)

Retorna uma lista de dicts: [{'PO': str, 'Produto': str, 'Qtd': int}, ...]
"""

import logging
import re
import unicodedata
from pathlib import Path
from typing import List, Dict, Optional

import pdfplumber

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Lista estrita de produtos reconhecidos
# ---------------------------------------------------------------------------
PRODUTOS: List[str] = [
    "ACCESS POINT - INDOOR 4X4",
    "ACCESS POINT - INDOOR 4X4 - SERVIÇO",
    "SERVIÇO INST. ACCESS POINT INDOOR",
    "ACCESS POINT - OUTDOOR",
    "ACCESS POINT - OUTDOOR - SERVIÇO",
    "SERVIÇO INST. ACCESS POINT OUTDOOR",
    "INTERFACE UTP 1G SWITCH L3",
    "INTERFACE ÓPTICO 1G SW L2 E L3",
    "INTERFACE ÓPTICO 10G SW L2 E L3",
    "MÓDULO DE EMPILHAMENTO DE SWITCH",
    "MÓDULO DE EMPILHAMENTO SWITCH - SERVIÇO",
    "SERVIÇO VISTORIA UN. PLANO IMPLANTAÇÃO",
    "SWITCH L2 - 24 PORTAS",
    "SWITCH L2 - 24 PORTAS - SERVIÇO",
    "SERVIÇO INST. SWITCH L2 - 24 PORTAS",
    "SWITCH L2 - 24 PORTAS MULTGIGA",
    "SWITCH L2 - 24 PORTAS MULTGIGA - SERVIÇO",
    "SERVIÇO INST. SWITCH L2 24 PORTAS MULT.",
    "SWITCH L2 - 48 PORTAS",
    "SWITCH L2 - 48 PORTAS - SERVIÇO",
    "SERVIÇO INST. SWITCH L2 - 48 PORTAS",
    "SWITCH L3 - 24 PORTAS",
    "SWITCH L3 - 24 PORTAS - SERVIÇO",
    "SERVIÇO INST. SWITCH L3 - 24 PORTAS",
    "SWITCH L3 - 48 PORTAS",
    "SWITCH L3 - 48 PORTAS - SERVIÇO",
    "SERVIÇO INST. SWITCH L3 - 48 PORTAS",
    "SOFTWARE GERENC. LAN WLAN E CONTROLLER",
    "TREINAMENTO CERTIFICAÇÃO FABRICANTE",
]

# ---------------------------------------------------------------------------
# Regex helpers
# ---------------------------------------------------------------------------

# Captura o número do pedido após "Nº do Pedido:" ou variações comuns
_RE_PO = re.compile(
    r"N[ºo°]?\s*\.?\s*(?:DO\s+)?PEDIDO\s*:?\s*(\d{7,12})",
    re.IGNORECASE,
)

# Quantidade: número inteiro (sem vírgula ou ponto decimal) seguido de "UN"
# ou precedido de "UN". Ex: "140 UN", "UN 5", "22UN"
# Exclui números que possuem vírgula/ponto antes de mais dígitos (preços).
_RE_QTD_AFTER_PRODUTO = re.compile(
    r"""
    (?<!\d[.,])          # não precedido de dígito com separador decimal
    \b
    (\d{1,6})            # captura: 1 a 6 dígitos (quantidade)
    \b
    (?![.,]\d)           # não seguido de separador decimal
    \s*
    UN\b                 # obrigatoriamente seguido de "UN"
    """,
    re.VERBOSE,
)

# Também tentamos capturar "UN \d+" (UN antes do número)
_RE_UN_BEFORE_QTD = re.compile(
    r"\bUN\s+(\d{1,6})\b(?![.,]\d)"
)


def _normalize_text(raw: str) -> str:
    """
    Converte para maiúsculas, normaliza unicode e colapsa espaços/quebras.
    Mantém a pontuação relevante (hífens, pontos finais de abreviaturas).
    """
    text = raw.upper()
    # Normaliza caracteres unicode (ex: óptico → ÓPTICO já está ok em upper)
    text = unicodedata.normalize("NFC", text)
    # Substitui quebras de linha e tabulações por espaço
    text = re.sub(r"[\r\n\t]+", " ", text)
    # Colapsa múltiplos espaços
    text = re.sub(r" {2,}", " ", text)
    return text.strip()


def _extract_text_from_pdf(pdf_path: Path) -> Optional[str]:
    """Extrai o texto completo de um PDF usando pdfplumber."""
    pages_text: List[str] = []
    try:
        with pdfplumber.open(str(pdf_path)) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text(x_tolerance=2, y_tolerance=2)
                if page_text:
                    pages_text.append(page_text)
    except Exception as exc:
        logger.error("Falha ao abrir PDF '%s': %s", pdf_path.name, exc)
        return None

    if not pages_text:
        logger.warning("PDF '%s' não retornou texto algum.", pdf_path.name)
        return None

    return "\n".join(pages_text)


def _find_po(text_normalized: str) -> Optional[str]:
    """Tenta localizar o número do PO no texto normalizado."""
    match = _RE_PO.search(text_normalized)
    if match:
        return match.group(1).strip()
    logger.warning("Número do PO não encontrado no texto.")
    return None


def _escape_for_regex(product_name: str) -> str:
    """
    Escapa o nome do produto para uso em regex, preservando hífens e pontos
    que fazem parte do nome (ex: 'SWITCH L2 - 24 PORTAS').
    """
    # re.escape escapa tudo; depois restauramos espaços como \s+
    escaped = re.escape(product_name)
    # Permite 1 ou mais espaços entre tokens
    escaped = re.sub(r"(\ |\\ )+", r"\\s+", escaped)
    return escaped


def _find_quantity_for_product(
    product_name: str,
    text_normalized: str,
) -> int:
    """
    Localiza a quantidade (inteiro) de um produto no texto normalizado.

    Estratégia:
      1. Encontra a posição do nome do produto no texto.
      2. Extrai uma janela de contexto (até 120 chars após o nome).
      3. Busca o padrão <inteiro> UN ou UN <inteiro> na janela.
      4. Se não encontrar, tenta uma janela antes do nome.
      5. Retorna 0 se nenhuma quantidade válida for encontrada.
    """
    pattern = _escape_for_regex(product_name)
    try:
        match = re.search(pattern, text_normalized)
    except re.error as exc:
        logger.error(
            "Regex inválida para produto '%s': %s", product_name, exc
        )
        return 0

    if not match:
        return 0

    # Janela após o produto (onde costuma aparecer UN e qtd)
    start = match.end()
    window_after = text_normalized[start : start + 150]

    qty = _parse_quantity_from_window(window_after)
    if qty is not None:
        return qty

    # Fallback: janela antes do produto
    end = match.start()
    window_before = text_normalized[max(0, end - 60) : end]
    qty = _parse_quantity_from_window(window_before)
    if qty is not None:
        return qty

    logger.debug(
        "Quantidade não encontrada para '%s'. Janela: '%s'",
        product_name,
        window_after,
    )
    return 0


def _parse_quantity_from_window(window: str) -> Optional[int]:
    """
    Tenta extrair a quantidade de uma janela de texto.
    Prioriza padrão <inteiro> UN, depois UN <inteiro>.
    Ignora números com vírgula/ponto (preços).
    """
    # Padrão principal: <inteiro> UN
    m = _RE_QTD_AFTER_PRODUTO.search(window)
    if m:
        return int(m.group(1))

    # Padrão alternativo: UN <inteiro>
    m = _RE_UN_BEFORE_QTD.search(window)
    if m:
        return int(m.group(1))

    return None


def parse_pdf(pdf_path: Path) -> List[Dict]:
    """
    Processa um único PDF e retorna uma lista de dicionários com
    {'PO': str, 'Produto': str, 'Qtd': int} para cada produto encontrado.

    Produtos não encontrados no PDF NÃO são incluídos (o formatter.py
    garantirá que todos os produtos apareçam com 0 na tabela final).
    """
    logger.info("Fazendo parse do PDF: %s", pdf_path.name)

    raw_text = _extract_text_from_pdf(pdf_path)
    if raw_text is None:
        return []

    text = _normalize_text(raw_text)

    po_number = _find_po(text)
    if po_number is None:
        logger.error(
            "'%s': PO não identificado — arquivo ignorado.", pdf_path.name
        )
        return []

    logger.info("PO identificado: %s", po_number)

    results: List[Dict] = []
    for produto in PRODUTOS:
        qty = _find_quantity_for_product(produto, text)
        if qty > 0:
            results.append(
                {"PO": po_number, "Produto": produto, "Qtd": qty}
            )
            logger.debug("  %s → Qtd: %d", produto, qty)

    logger.info(
        "PDF '%s' | PO %s | %d produto(s) com quantidade > 0.",
        pdf_path.name,
        po_number,
        len(results),
    )
    return results


def parse_all_pdfs(pdf_paths: List[Path]) -> List[Dict]:
    """
    Itera sobre uma lista de PDFs e consolida todos os resultados.

    Parameters
    ----------
    pdf_paths : list[Path]
        Caminhos dos PDFs a processar.

    Returns
    -------
    list[dict]
        Lista consolidada de {'PO', 'Produto', 'Qtd'}.
    """
    all_records: List[Dict] = []
    for pdf_path in pdf_paths:
        records = parse_pdf(pdf_path)
        all_records.extend(records)
    return all_records
