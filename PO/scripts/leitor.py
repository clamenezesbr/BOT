"""
acumulador_sesi.py
------------------
Lê todos os PDFs de cotação SESI da pasta 'saida/', consolida os itens
por ID de Produto (somando quantidades) e gera um log detalhado em:

    logs/acumulador_YYYYMMDD_HHMMSS.log

Dependências:
    pip install pdfplumber
"""

import re
import logging
import pdfplumber
from pathlib import Path
from datetime import datetime

# ---------------------------------------------------------------------------
# CONFIGURAÇÃO DE PASTAS
# ---------------------------------------------------------------------------
BASE          = Path(__file__).parent.parent
PASTA_ENTRADA = BASE / "saida"          # PDFs a serem lidos
PASTA_LOGS    = BASE / "logs"           # onde o .log será criado

# ---------------------------------------------------------------------------
# CONFIGURAÇÃO DO LOGGER DETALHADO
# ---------------------------------------------------------------------------

def _configurar_logger() -> tuple[logging.Logger, Path]:
    """
    Cria um logger que escreve simultaneamente:
      - Terminal    : nível INFO  (mensagens resumidas)
      - Arquivo .log: nível DEBUG (tudo, com timestamp)
    """
    PASTA_LOGS.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path  = PASTA_LOGS / f"acumulador_{timestamp}.log"

    logger = logging.getLogger("acumulador_sesi")
    logger.setLevel(logging.DEBUG)
    logger.handlers.clear()

    # Handler de arquivo — DEBUG+
    fmt_arq = logging.Formatter(
        fmt="%(asctime)s  [%(levelname)-8s]  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    fh = logging.FileHandler(log_path, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt_arq)
    logger.addHandler(fh)

    # Handler de console — INFO+
    fmt_con = logging.Formatter(fmt="%(levelname)-8s  %(message)s")
    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    ch.setFormatter(fmt_con)
    logger.addHandler(ch)

    logger.info(f"Log iniciado → {log_path.resolve()}")
    return logger, log_path


# ---------------------------------------------------------------------------
# 1. ACUMULADOR GLOBAL
# ---------------------------------------------------------------------------
# acumulador[id_produto] = {
#     "descricao":      str,
#     "quantidade":     float,
#     "valor_unitario": float,
#     "arquivos":       list[str],   # quais PDFs contribuíram
# }
acumulador: dict[str, dict] = {}


# ---------------------------------------------------------------------------
# 2. HELPERS DE LIMPEZA
# ---------------------------------------------------------------------------

def _limpar_numero(texto: str) -> float:
    """Converte strings BR ('6.727,00', '10') para float."""
    texto = texto.strip().replace(".", "").replace(",", ".")
    try:
        return float(texto)
    except ValueError:
        return 0.0


def _primeira_linha(texto: str) -> str:
    """Retorna só a 1ª linha — evita notas de garantia multi-linha."""
    return texto.strip().splitlines()[0].strip() if texto else ""


# ---------------------------------------------------------------------------
# 3. EXTRAÇÃO DE UM ÚNICO PDF
# ---------------------------------------------------------------------------

def _extrair_itens_pdf(caminho_pdf: str, log: logging.Logger) -> list[dict]:
    """
    Varre todas as tabelas do PDF e devolve lista de dicts com:
        id_produto, descricao, quantidade, valor_unitario
    """
    itens: list[dict] = []
    cabecalhos_esperados = {"id", "produto", "item", "quant", "valor"}

    log.debug(f"  Abrindo: {caminho_pdf}")

    with pdfplumber.open(caminho_pdf) as pdf:
        log.debug(f"  Total de páginas: {len(pdf.pages)}")

        for num_pag, pagina in enumerate(pdf.pages, start=1):
            tabelas = pagina.extract_tables()
            log.debug(f"  Pág. {num_pag}: {len(tabelas)} tabela(s)")

            for t_idx, tabela in enumerate(tabelas, start=1):
                if not tabela:
                    log.debug(f"    Tabela {t_idx}: vazia, ignorada")
                    continue

                # Localizar linha de cabeçalho
                idx_cab = None
                for i, linha in enumerate(tabela):
                    texto = " ".join(str(c or "").lower() for c in linha)
                    if any(h in texto for h in cabecalhos_esperados):
                        idx_cab = i
                        break

                if idx_cab is None:
                    log.debug(f"    Tabela {t_idx}: cabeçalho não reconhecido, ignorada")
                    continue

                cabecalho = [str(c or "").lower().strip() for c in tabela[idx_cab]]
                log.debug(f"    Tabela {t_idx} pág {num_pag}: cabeçalho linha {idx_cab} → {cabecalho}")

                # Mapear colunas dinamicamente
                def achar_col(*chaves):
                    for chave in chaves:
                        for j, c in enumerate(cabecalho):
                            if chave in c:
                                return j
                    return None

                col_id   = achar_col("id", "produto", "item")
                col_desc = achar_col("descri", "nome")
                col_qtd  = achar_col("quant")
                col_vl   = achar_col("valor unit", "vl unit", "unitário", "unitario")

                log.debug(
                    f"    Colunas → ID:{col_id}  DESC:{col_desc}"
                    f"  QTD:{col_qtd}  VL_UNIT:{col_vl}"
                )

                if col_id is None or col_qtd is None:
                    log.warning(
                        f"    Tabela {t_idx} pág {num_pag}: colunas mínimas ausentes"
                        f" (ID={col_id}, QTD={col_qtd}) — ignorada"
                    )
                    continue

                # Iterar linhas de dados
                extraidos = 0
                for r_idx, linha in enumerate(tabela[idx_cab + 1:], start=idx_cab + 1):
                    if not linha or all(c is None or str(c).strip() == "" for c in linha):
                        log.debug(f"      Linha {r_idx}: vazia")
                        continue

                    def cel(idx):
                        if idx is None or idx >= len(linha):
                            return ""
                        return str(linha[idx] or "").strip()

                    id_bruto  = cel(col_id)
                    match_id  = re.search(r"\b\d{5,}\b", id_bruto)

                    if not match_id:
                        log.debug(f"      Linha {r_idx}: ID inválido ('{id_bruto}'), ignorada")
                        continue

                    id_produto   = match_id.group()
                    descricao    = _primeira_linha(cel(col_desc)) if col_desc is not None else ""
                    quantidade   = _limpar_numero(cel(col_qtd))
                    valor_unit   = _limpar_numero(cel(col_vl)) if col_vl is not None else 0.0

                    if quantidade <= 0:
                        log.debug(f"      Linha {r_idx}: quantidade inválida ({quantidade})")
                        continue

                    log.debug(
                        f"      Linha {r_idx}: ID={id_produto}  "
                        f"DESC='{descricao[:40]}'  QTD={quantidade}  VL={valor_unit}"
                    )
                    itens.append({
                        "id_produto":     id_produto,
                        "descricao":      descricao,
                        "quantidade":     quantidade,
                        "valor_unitario": valor_unit,
                    })
                    extraidos += 1

                log.debug(f"    Tabela {t_idx}: {extraidos} item(ns) extraído(s)")

    log.info(f"  Itens extraídos do arquivo: {len(itens)}")
    return itens


# ---------------------------------------------------------------------------
# 4. LÓGICA DE SOMATÓRIA
# ---------------------------------------------------------------------------

def _acumular(itens: list[dict], nome_arquivo: str, log: logging.Logger) -> None:
    """Aplica decisão SOMA / NOVO sobre o acumulador global."""
    novos   = 0
    somados = 0

    for item in itens:
        pid = item["id_produto"]

        if pid in acumulador:
            antes = acumulador[pid]["quantidade"]
            acumulador[pid]["quantidade"] += item["quantidade"]
            acumulador[pid]["arquivos"].append(nome_arquivo)
            log.debug(
                f"  SOMA  ID={pid}  '{acumulador[pid]['descricao'][:30]}'  "
                f"{antes} + {item['quantidade']} = {acumulador[pid]['quantidade']}"
            )
            somados += 1
        else:
            acumulador[pid] = {
                "descricao":      item["descricao"],
                "quantidade":     item["quantidade"],
                "valor_unitario": item["valor_unitario"],
                "arquivos":       [nome_arquivo],
            }
            log.debug(
                f"  NOVO  ID={pid}  '{item['descricao'][:30]}'  "
                f"QTD={item['quantidade']}  VL={item['valor_unitario']}"
            )
            novos += 1

    log.info(f"  Acumulador atualizado: {novos} novo(s)  |  {somados} somado(s)")


# ---------------------------------------------------------------------------
# 5. RELATÓRIO FINAL
# ---------------------------------------------------------------------------

def _gerar_relatorio(log: logging.Logger, log_path: Path) -> None:
    """Grava o relatório consolidado no logger (terminal + arquivo)."""
    sep  = "=" * 90
    sep2 = "-" * 90

    total_valor = sum(
        d["quantidade"] * d["valor_unitario"] for d in acumulador.values()
    )

    linhas = [
        "",
        sep,
        f"{'RELATÓRIO CONSOLIDADO DE PRODUTOS':^90}",
        sep,
        f"  {'ID Produto':<12}  {'Descrição':<34}  {'Qtd':>8}  {'Vl Unitário':>13}  Origem (arquivos)",
        sep2,
    ]

    for pid, d in sorted(acumulador.items()):
        arquivos_str = ", ".join(sorted(set(d["arquivos"])))
        linhas.append(
            f"  {pid:<12}  {d['descricao'][:34]:<34}  "
            f"{d['quantidade']:>8.0f}  "
            f"R$ {d['valor_unitario']:>10,.2f}  "
            f"{arquivos_str}"
        )

    linhas += [
        sep2,
        f"  SKUs distintos : {len(acumulador)}",
        f"  Valor total    : R$ {total_valor:,.2f}",
        sep,
        f"  Log completo   : {log_path.resolve()}",
        sep,
    ]

    for linha in linhas:
        log.info(linha)


# ---------------------------------------------------------------------------
# 6. FUNÇÃO PRINCIPAL PÚBLICA
# ---------------------------------------------------------------------------

def processar_pasta_saida() -> dict:
    """
    Ponto de entrada principal.

    - Lê todos os *.pdf de PASTA_ENTRADA ('saida/')
    - Consolida no acumulador global
    - Grava log detalhado em PASTA_LOGS ('logs/')
    - Retorna o dicionário acumulador final
    """
    global acumulador
    acumulador = {}

    log, log_path = _configurar_logger()

    log.info(f"Pasta de entrada : {PASTA_ENTRADA.resolve()}")
    log.info(f"Pasta de logs    : {PASTA_LOGS.resolve()}")

    if not PASTA_ENTRADA.exists():
        log.error(
            f"Pasta '{PASTA_ENTRADA}' não encontrada. "
            "Crie-a e coloque os PDFs SESI dentro."
        )
        return acumulador

    pdfs = sorted(PASTA_ENTRADA.glob("*.pdf"))
    log.info(f"PDFs encontrados : {len(pdfs)}")

    if not pdfs:
        log.warning("Nenhum PDF encontrado em 'saida/'. Encerrando sem processar.")
        return acumulador

    # --- Processar cada PDF ---
    for i, pdf_path in enumerate(pdfs, start=1):
        log.info("")
        log.info(f"{'─' * 60}")
        log.info(f"[{i}/{len(pdfs)}] {pdf_path.name}")
        log.info(f"  Tamanho: {pdf_path.stat().st_size / 1024:.1f} KB")

        try:
            itens = _extrair_itens_pdf(str(pdf_path), log)
            _acumular(itens, pdf_path.name, log)
        except Exception as exc:
            log.error(f"  FALHA ao processar {pdf_path.name}: {exc}", exc_info=True)

    # --- Relatório final ---
    _gerar_relatorio(log, log_path)

    return acumulador


# ---------------------------------------------------------------------------
# EXECUÇÃO DIRETA
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    processar_pasta_saida()
