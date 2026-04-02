"""
acumulador_sesi.py
------------------
Lê todos os PDFs de cotação SESI da pasta 'saida/', filtra apenas os produtos
do catálogo oficial do projeto, consolida quantidades e gera:

  1. logs/acumulador_YYYYMMDD_HHMMSS.log  — log detalhado (DEBUG)
  2. relatorios/relatorio_YYYYMMDD_HHMMSS.pdf — relatório final em PDF

Double Check:
  Cada PDF é lido DUAS vezes de forma independente. Se as quantidades ou
  valores extraídos nas duas leituras diferirem, o item é marcado como
  ⚠ SUSPEITO no relatório PDF e destacado no log.

Dependências:
    pip install pdfplumber reportlab
"""

import re
import copy
import logging
import pdfplumber
from pathlib import Path
from datetime import datetime

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
)

# ---------------------------------------------------------------------------
# CONFIGURAÇÃO DE PASTAS
# ---------------------------------------------------------------------------
BASE          = Path(__file__).parent.parent
PASTA_ENTRADA = BASE / "saida"
PASTA_LOGS    = BASE / "logs"
PASTA_RELAT   = BASE / "relatorios"

# Tolerância para double check: diferença percentual máxima aceitável
# entre as duas leituras antes de marcar como suspeito.
TOLERANCIA_DOUBLE_CHECK = 0.0   # 0% = qualquer diferença já é suspeita

# ---------------------------------------------------------------------------
# CATÁLOGO OFICIAL DE PRODUTOS DO PROJETO
# ---------------------------------------------------------------------------
CATALOGO_OFICIAL: list[str] = [
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

def _tokenizar(texto: str) -> set[str]:
    texto = texto.upper()
    texto = re.sub(r"[^\w\s]", " ", texto)
    return set(texto.split())

_CATALOGO_TOKENS: list[tuple[set[str], str]] = [
    (_tokenizar(n), n) for n in CATALOGO_OFICIAL
]
_LIMIAR_SIMILARIDADE = 0.75

def _encontrar_no_catalogo(descricao: str) -> str | None:
    tokens_desc = _tokenizar(descricao)
    if not tokens_desc:
        return None
    melhor_nome  = None
    melhor_score = 0.0
    for tokens_cat, nome_cat in _CATALOGO_TOKENS:
        uniao = tokens_desc | tokens_cat
        score = len(tokens_desc & tokens_cat) / len(uniao) if uniao else 0.0
        if score > melhor_score:
            melhor_score = score
            melhor_nome  = nome_cat
    return melhor_nome if melhor_score >= _LIMIAR_SIMILARIDADE else None


# ---------------------------------------------------------------------------
# ACUMULADOR GLOBAL
# ---------------------------------------------------------------------------
# acumulador[nome_canonico] = {
#     "quantidade":     float,
#     "valor_unitario": float,
#     "arquivos":       list[str],
#     "suspeito":       bool,          ← definido pelo double check
#     "divergencias":   list[str],     ← descrição de cada divergência
# }
acumulador: dict[str, dict] = {}
rejeitados: list[dict]      = []


# ---------------------------------------------------------------------------
# LOGGER
# ---------------------------------------------------------------------------
def _configurar_logger() -> tuple[logging.Logger, Path]:
    PASTA_LOGS.mkdir(parents=True, exist_ok=True)
    ts       = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = PASTA_LOGS / f"acumulador_{ts}.log"

    logger = logging.getLogger("acumulador_sesi")
    logger.setLevel(logging.DEBUG)
    logger.handlers.clear()

    fmt_arq = logging.Formatter("%(asctime)s  [%(levelname)-8s]  %(message)s",
                                datefmt="%Y-%m-%d %H:%M:%S")
    fh = logging.FileHandler(log_path, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt_arq)
    logger.addHandler(fh)

    fmt_con = logging.Formatter("%(levelname)-8s  %(message)s")
    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    ch.setFormatter(fmt_con)
    logger.addHandler(ch)

    logger.info(f"Log iniciado → {log_path.resolve()}")
    return logger, log_path


# ---------------------------------------------------------------------------
# HELPERS NUMÉRICOS
# ---------------------------------------------------------------------------
def _limpar_numero(texto: str) -> float:
    texto = texto.strip().replace(".", "").replace(",", ".")
    try:
        return float(texto)
    except ValueError:
        return 0.0

def _primeira_linha(texto: str) -> str:
    return texto.strip().splitlines()[0].strip() if texto else ""


# ---------------------------------------------------------------------------
# EXTRAÇÃO CORE  (usada duas vezes por PDF)
# ---------------------------------------------------------------------------
def _extrair_itens_pdf(caminho_pdf: str, log: logging.Logger,
                       leitura: int = 1) -> list[dict]:
    """
    Extrai itens do PDF. Chamada com leitura=1 e leitura=2 para double check.
    Retorna lista de dicts: id_produto, descricao, quantidade, valor_unitario.
    """
    itens: list[dict] = []
    cabecalhos_esperados = {"id", "produto", "item", "quant", "valor"}

    log.debug(f"  [Leitura {leitura}] Abrindo: {caminho_pdf}")

    with pdfplumber.open(caminho_pdf) as pdf:
        log.debug(f"  [Leitura {leitura}] Páginas: {len(pdf.pages)}")

        for num_pag, pagina in enumerate(pdf.pages, start=1):
            tabelas = pagina.extract_tables()
            log.debug(f"  [L{leitura}] Pág. {num_pag}: {len(tabelas)} tabela(s)")

            for t_idx, tabela in enumerate(tabelas, start=1):
                if not tabela:
                    continue

                idx_cab = None
                for i, linha in enumerate(tabela):
                    texto = " ".join(str(c or "").lower() for c in linha)
                    if any(h in texto for h in cabecalhos_esperados):
                        idx_cab = i
                        break

                if idx_cab is None:
                    log.debug(f"    [L{leitura}] Tabela {t_idx}: cabeçalho não reconhecido")
                    continue

                cabecalho = [str(c or "").lower().strip() for c in tabela[idx_cab]]

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

                if col_id is None or col_qtd is None:
                    log.warning(f"    [L{leitura}] Tabela {t_idx} pág {num_pag}: colunas mínimas ausentes")
                    continue

                for r_idx, linha in enumerate(tabela[idx_cab + 1:], start=idx_cab + 1):
                    if not linha or all(c is None or str(c).strip() == "" for c in linha):
                        continue

                    def cel(idx):
                        if idx is None or idx >= len(linha):
                            return ""
                        return str(linha[idx] or "").strip()

                    id_bruto = cel(col_id)
                    match_id = re.search(r"\b\d{5,}\b", id_bruto)
                    if not match_id:
                        continue

                    id_produto   = match_id.group()
                    descricao    = _primeira_linha(cel(col_desc)) if col_desc is not None else ""
                    quantidade   = _limpar_numero(cel(col_qtd))
                    valor_unit   = _limpar_numero(cel(col_vl)) if col_vl is not None else 0.0

                    if quantidade <= 0:
                        continue

                    log.debug(
                        f"    [L{leitura}] Linha {r_idx}: ID={id_produto}"
                        f"  QTD={quantidade}  VL={valor_unit}"
                        f"  DESC='{descricao[:40]}'"
                    )
                    itens.append({
                        "id_produto":     id_produto,
                        "descricao":      descricao,
                        "quantidade":     quantidade,
                        "valor_unitario": valor_unit,
                    })

    log.info(f"  [Leitura {leitura}] Itens brutos extraídos: {len(itens)}")
    for it in itens:
        log.info(
            f"    → ID {it['id_produto']:>8}  |  "
            f"Qtd: {it['quantidade']:>6.0f}  |  "
            f"Vl: R$ {it['valor_unitario']:>10,.2f}  |  "
            f"{it['descricao']}"
        )
    return itens


# ---------------------------------------------------------------------------
# DOUBLE CHECK — compara leitura 1 vs leitura 2
# ---------------------------------------------------------------------------
def _double_check_pdf(caminho_pdf: str, log: logging.Logger
                      ) -> tuple[list[dict], list[str]]:
    """
    Executa duas leituras independentes do mesmo PDF e compara item a item.

    Retorna:
        itens_finais : lista com os itens da leitura 1 (dados primários),
                       já com flag 'suspeito' e 'divergencia' preenchidos.
        resumo_diverg: lista de strings descritivas para o log e PDF.
    """
    log.info(f"  ── DOUBLE CHECK ──────────────────────────────────────────")
    leitura1 = _extrair_itens_pdf(caminho_pdf, log, leitura=1)
    leitura2 = _extrair_itens_pdf(caminho_pdf, log, leitura=2)

    # Indexa leitura 2 por id_produto para comparação rápida
    indice_l2: dict[str, list[dict]] = {}
    for item in leitura2:
        indice_l2.setdefault(item["id_produto"], []).append(item)

    resumo_diverg: list[str] = []
    itens_finais:  list[dict] = []

    for item in leitura1:
        pid     = item["id_produto"]
        item    = dict(item)           # cópia para não mutar o original
        item["suspeito"]   = False
        item["divergencia"] = ""

        candidatos_l2 = indice_l2.get(pid, [])

        if not candidatos_l2:
            # Item presente na leitura 1 mas ausente na leitura 2
            msg = (
                f"ID {pid} ('{item['descricao'][:35]}') encontrado na "
                f"leitura 1 mas AUSENTE na leitura 2"
            )
            item["suspeito"]    = True
            item["divergencia"] = "Ausente na leitura 2"
            resumo_diverg.append(msg)
            log.warning(f"  ⚠ SUSPEITO — {msg}")

        else:
            # Usa o primeiro candidato da leitura 2 com o mesmo ID
            l2 = candidatos_l2[0]

            divergencias_item: list[str] = []

            # — Verificar quantidade —
            qtd1, qtd2 = item["quantidade"], l2["quantidade"]
            dif_qtd = abs(qtd1 - qtd2)
            tol_abs = max(qtd1, qtd2) * TOLERANCIA_DOUBLE_CHECK if max(qtd1, qtd2) > 0 else 0
            if dif_qtd > tol_abs:
                msg = (
                    f"ID {pid} ('{item['descricao'][:35]}'): "
                    f"QTD leitura1={qtd1} ≠ leitura2={qtd2} "
                    f"(Δ={dif_qtd:+.2f})"
                )
                divergencias_item.append(f"Qtd L1={qtd1} / L2={qtd2}")
                resumo_diverg.append(msg)
                log.warning(f"  ⚠ SUSPEITO — {msg}")

            # — Verificar valor unitário —
            vl1, vl2 = item["valor_unitario"], l2["valor_unitario"]
            dif_vl = abs(vl1 - vl2)
            tol_vl = max(vl1, vl2) * TOLERANCIA_DOUBLE_CHECK if max(vl1, vl2) > 0 else 0
            if dif_vl > tol_vl:
                msg = (
                    f"ID {pid} ('{item['descricao'][:35]}'): "
                    f"VL leitura1={vl1:.2f} ≠ leitura2={vl2:.2f} "
                    f"(Δ=R$ {dif_vl:+.2f})"
                )
                divergencias_item.append(f"Vl L1=R${vl1:.2f} / L2=R${vl2:.2f}")
                resumo_diverg.append(msg)
                log.warning(f"  ⚠ SUSPEITO — {msg}")

            if divergencias_item:
                item["suspeito"]    = True
                item["divergencia"] = " | ".join(divergencias_item)
            else:
                log.debug(f"  ✔ OK — ID {pid}  QTD={qtd1}  VL={vl1:.2f}")

        itens_finais.append(item)

    # Itens que apareceram na leitura 2 mas não na leitura 1
    ids_l1 = {i["id_produto"] for i in leitura1}
    for pid, candidatos in indice_l2.items():
        if pid not in ids_l1:
            l2  = candidatos[0]
            msg = (
                f"ID {pid} ('{l2['descricao'][:35]}') AUSENTE na leitura 1 "
                f"mas encontrado na leitura 2 — item adicionado como suspeito"
            )
            item_extra = dict(l2)
            item_extra["suspeito"]    = True
            item_extra["divergencia"] = "Ausente na leitura 1"
            itens_finais.append(item_extra)
            resumo_diverg.append(msg)
            log.warning(f"  ⚠ SUSPEITO — {msg}")

    total_susp = sum(1 for i in itens_finais if i["suspeito"])
    log.info(
        f"  Double check: {len(itens_finais)} item(ns)  |  "
        f"{total_susp} suspeito(s)  |  "
        f"{len(itens_finais) - total_susp} confirmado(s)"
    )
    log.info(f"  ─────────────────────────────────────────────────────────")

    return itens_finais, resumo_diverg


# ---------------------------------------------------------------------------
# FILTRO POR CATÁLOGO
# ---------------------------------------------------------------------------
def _filtrar_catalogo(itens: list[dict], nome_arquivo: str,
                      log: logging.Logger) -> list[dict]:
    aceitos: list[dict] = []
    for item in itens:
        nome_canonico = _encontrar_no_catalogo(item["descricao"])
        if nome_canonico:
            log.debug(f"  ACEITO  '{item['descricao'][:40]}' → '{nome_canonico}'")
            item["nome_canonico"] = nome_canonico
            aceitos.append(item)
        else:
            log.debug(f"  REJEIT. '{item['descricao'][:40]}' — fora do catálogo")
            rejeitados.append({**item, "arquivo": nome_arquivo})
    log.info(
        f"  Filtro catálogo: {len(aceitos)} aceito(s)  |  "
        f"{len(itens)-len(aceitos)} rejeitado(s)"
    )
    return aceitos


# ---------------------------------------------------------------------------
# SOMATÓRIA
# ---------------------------------------------------------------------------
def _acumular(itens: list[dict], nome_arquivo: str,
              log: logging.Logger) -> None:
    novos = somados = 0
    for item in itens:
        chave   = item["nome_canonico"]
        susp    = item.get("suspeito", False)
        diverg  = item.get("divergencia", "")

        if chave in acumulador:
            antes = acumulador[chave]["quantidade"]
            acumulador[chave]["quantidade"] += item["quantidade"]
            acumulador[chave]["arquivos"].append(nome_arquivo)
            if susp:
                acumulador[chave]["suspeito"] = True
                if diverg:
                    acumulador[chave]["divergencias"].append(
                        f"[{nome_arquivo}] {diverg}"
                    )
            log.debug(
                f"  SOMA  '{chave[:40]}'  "
                f"{antes} + {item['quantidade']} = {acumulador[chave]['quantidade']}"
                + (f"  ⚠ SUSPEITO" if susp else "")
            )
            somados += 1
        else:
            acumulador[chave] = {
                "quantidade":     item["quantidade"],
                "valor_unitario": item["valor_unitario"],
                "arquivos":       [nome_arquivo],
                "suspeito":       susp,
                "divergencias":   [f"[{nome_arquivo}] {diverg}"] if (susp and diverg) else [],
            }
            log.debug(
                f"  NOVO  '{chave[:40]}'  "
                f"QTD={item['quantidade']}  VL={item['valor_unitario']}"
                + (f"  ⚠ SUSPEITO" if susp else "")
            )
            novos += 1

    log.info(f"  Acumulador: {novos} novo(s)  |  {somados} somado(s)")


# ---------------------------------------------------------------------------
# GERAÇÃO DO PDF RELATÓRIO
# ---------------------------------------------------------------------------
def _gerar_pdf(log: logging.Logger) -> Path:
    PASTA_RELAT.mkdir(parents=True, exist_ok=True)
    ts       = datetime.now().strftime("%Y%m%d_%H%M%S")
    pdf_path = PASTA_RELAT / f"relatorio_{ts}.pdf"

    doc = SimpleDocTemplate(
        str(pdf_path), pagesize=A4,
        leftMargin=2*cm, rightMargin=2*cm,
        topMargin=2*cm,  bottomMargin=2*cm,
    )

    styles = getSampleStyleSheet()
    COR_AZUL_ESC  = colors.HexColor("#1A3A5C")
    COR_AZUL_MED  = colors.HexColor("#2E7D9C")
    COR_AZUL_CLAR = colors.HexColor("#EAF2F8")
    COR_AMARELO   = colors.HexColor("#FFF3CD")  # fundo linha suspeita
    COR_LARANJA   = colors.HexColor("#E67E22")  # texto alerta suspeito
    COR_TOTAL_BG  = colors.HexColor("#1A3A5C")

    def estilo(nome, **kw):
        return ParagraphStyle(nome, parent=styles["Normal"], **kw)

    e_titulo  = estilo("Tit", fontSize=16, textColor=COR_AZUL_ESC,
                       spaceAfter=4, alignment=TA_CENTER)
    e_sub     = estilo("Sub", fontSize=9,  textColor=colors.HexColor("#555555"),
                       alignment=TA_CENTER, spaceAfter=12)
    e_nota    = estilo("Nota", fontSize=8, textColor=colors.HexColor("#777777"),
                       spaceBefore=10)
    e_cel     = estilo("Cel",  fontSize=8, leading=10)
    e_cel_dir = estilo("CelD", fontSize=8, leading=10, alignment=TA_RIGHT)
    e_susp    = estilo("Susp", fontSize=8, leading=10, textColor=COR_LARANJA)
    e_susp_d  = estilo("SuspD",fontSize=8, leading=10, textColor=COR_LARANJA,
                       alignment=TA_RIGHT)
    e_diverg  = estilo("Div",  fontSize=7, leading=9,
                       textColor=colors.HexColor("#AA5500"))

    def hdr(txt):
        return Paragraph(
            f"<b>{txt}</b>",
            estilo("H", fontSize=8, textColor=colors.white, leading=10,
                   alignment=TA_RIGHT if txt != "Produto" else TA_CENTER)
        )

    n_suspeitos = sum(1 for d in acumulador.values() if d["suspeito"])

    story = []
    story.append(Paragraph("RELATÓRIO CONSOLIDADO DE PRODUTOS", e_titulo))
    story.append(Paragraph(
        f"Gerado em: {datetime.now().strftime('%d/%m/%Y às %H:%M:%S')}  |  "
        f"SKUs: {len(acumulador)}  |  "
        f"⚠ Suspeitos: {n_suspeitos}",
        e_sub,
    ))
    story.append(HRFlowable(width="100%", thickness=2, color=COR_AZUL_ESC, spaceAfter=14))

    # ── Cabeçalho da tabela ──
    dados = [[hdr("Produto"), hdr("Qtd."), hdr("Vl. Unitário"), hdr("Total"), hdr("DC")]]

    total_geral = 0.0
    estilos_cond: list[tuple] = []   # estilos condicionais por linha

    for idx, (nome, d) in enumerate(sorted(acumulador.items()), start=1):
        qtd   = d["quantidade"]
        vl    = d["valor_unitario"]
        total = qtd * vl
        total_geral += total
        susp  = d["suspeito"]
        diverg_txt = "; ".join(d["divergencias"])

        # Célula de nome: adiciona detalhe da divergência abaixo se suspeito
        if susp and diverg_txt:
            cel_nome = Paragraph(
                f"<b>{nome}</b><br/>"
                f"<font size='6.5' color='#AA5500'>⚠ {diverg_txt[:90]}</font>",
                e_susp if susp else e_cel,
            )
        elif susp:
            cel_nome = Paragraph(nome, e_susp)
        else:
            cel_nome = Paragraph(nome, e_cel)

        ec = e_susp_d if susp else e_cel_dir
        ic = "⚠" if susp else "✔"

        dados.append([
            cel_nome,
            Paragraph(f"{qtd:,.0f}",       ec),
            Paragraph(f"R$ {vl:,.2f}",     ec),
            Paragraph(f"R$ {total:,.2f}",  ec),
            Paragraph(ic,                  e_susp_d if susp else e_cel_dir),
        ])

        # Fundo amarelo para linhas suspeitas
        if susp:
            linha_pdf = idx   # linha 0 = cabeçalho, dados começam em 1
            estilos_cond.append(
                ("BACKGROUND", (0, linha_pdf), (-1, linha_pdf), COR_AMARELO)
            )

    # Linha de total geral
    n_linhas = len(dados)
    e_tot = estilo("Tot", fontSize=9, textColor=colors.white,
                   leading=11, alignment=TA_RIGHT)
    e_tot_l = estilo("TotL", fontSize=9, textColor=colors.white, leading=11)
    dados.append([
        Paragraph("<b>TOTAL GERAL</b>", e_tot_l),
        Paragraph("", e_tot),
        Paragraph("", e_tot),
        Paragraph(f"<b>R$ {total_geral:,.2f}</b>", e_tot),
        Paragraph("", e_tot),
    ])

    larg = A4[0] - 4*cm
    col_w = [larg*0.46, larg*0.10, larg*0.18, larg*0.18, larg*0.08]

    tabela_pdf = Table(dados, colWidths=col_w, repeatRows=1)
    estilo_base = TableStyle([
        ("BACKGROUND",     (0, 0),         (-1, 0),          COR_AZUL_ESC),
        ("TEXTCOLOR",      (0, 0),         (-1, 0),          colors.white),
        ("ROWBACKGROUNDS", (0, 1),         (-1, n_linhas-1), [colors.white, COR_AZUL_CLAR]),
        ("BACKGROUND",     (0, n_linhas),  (-1, n_linhas),   COR_TOTAL_BG),
        ("LINEABOVE",      (0, n_linhas),  (-1, n_linhas),   1.5, COR_AZUL_MED),
        ("GRID",           (0, 0),         (-1, n_linhas-1), 0.4, colors.HexColor("#CCCCCC")),
        ("TOPPADDING",     (0, 0),         (-1, -1),         5),
        ("BOTTOMPADDING",  (0, 0),         (-1, -1),         5),
        ("LEFTPADDING",    (0, 0),         (-1, -1),         6),
        ("RIGHTPADDING",   (0, 0),         (-1, -1),         6),
        ("VALIGN",         (0, 0),         (-1, -1),         "MIDDLE"),
    ])
    # Aplicar estilos condicionais (linhas suspeitas sobrescrevem rowbackground)
    for cond in estilos_cond:
        estilo_base.add(*cond)

    tabela_pdf.setStyle(estilo_base)
    story.append(tabela_pdf)

    # ── Legenda double check ──
    story.append(Spacer(1, 0.4*cm))
    story.append(Paragraph(
        "<b>Legenda Double Check:</b>  ✔ = leituras consistentes  |  "
        "⚠ = divergência entre as duas leituras — revisar manualmente",
        estilo("Leg", fontSize=8, textColor=COR_LARANJA, spaceBefore=4),
    ))

    if rejeitados:
        ids_rej = sorted(set(r["id_produto"] for r in rejeitados))
        story.append(Paragraph(
            f"<i>* {len(rejeitados)} item(ns) ignorado(s) por não pertencerem "
            f"ao catálogo (IDs: {', '.join(ids_rej[:15])}"
            f"{'…' if len(ids_rej) > 15 else ''}).</i>",
            e_nota,
        ))

    story.append(Spacer(1, 0.5*cm))
    story.append(HRFlowable(width="100%", thickness=1, color=COR_AZUL_MED))
    fontes = sorted(set(a for d in acumulador.values() for a in d["arquivos"]))
    story.append(Paragraph(f"<i>Fontes: {', '.join(fontes)}</i>", e_nota))

    doc.build(story)
    log.info(f"PDF gerado → {pdf_path.resolve()}")
    return pdf_path


# ---------------------------------------------------------------------------
# LOG FINAL
# ---------------------------------------------------------------------------
def _log_relatorio(log: logging.Logger, log_path: Path, pdf_path: Path) -> None:
    sep  = "=" * 100
    sep2 = "-" * 100
    total_geral = sum(d["quantidade"] * d["valor_unitario"] for d in acumulador.values())
    n_susp      = sum(1 for d in acumulador.values() if d["suspeito"])

    linhas = [
        "", sep,
        f"{'RELATÓRIO CONSOLIDADO — CATÁLOGO DO PROJETO':^100}",
        sep,
        f"  {'Produto':<48}  {'Qtd':>6}  {'Vl Unit.':>13}  {'Total':>13}  DC  Arquivos",
        sep2,
    ]
    for nome, d in sorted(acumulador.items()):
        total = d["quantidade"] * d["valor_unitario"]
        arqs  = ", ".join(sorted(set(d["arquivos"])))
        flag  = "⚠" if d["suspeito"] else "✔"
        linhas.append(
            f"  {nome[:48]:<48}  {d['quantidade']:>6.0f}  "
            f"R$ {d['valor_unitario']:>10,.2f}  R$ {total:>10,.2f}  {flag}   {arqs}"
        )
        for div in d["divergencias"]:
            linhas.append(f"    ↳ DIVERGÊNCIA: {div}")
    linhas += [
        sep2,
        f"  SKUs distintos  : {len(acumulador)}",
        f"  Itens suspeitos : {n_susp}",
        f"  Valor total     : R$ {total_geral:,.2f}",
        sep,
        f"  Log  → {log_path.resolve()}",
        f"  PDF  → {pdf_path.resolve()}",
        sep,
    ]
    for l in linhas:
        log.info(l)


# ---------------------------------------------------------------------------
# FUNÇÃO PRINCIPAL PÚBLICA
# ---------------------------------------------------------------------------
def processar_pasta_saida() -> dict:
    """
    Lê todos os PDFs de 'saida/', aplica double check (2 leituras independentes),
    filtra pelo catálogo, consolida e gera log + PDF.
    """
    global acumulador, rejeitados
    acumulador = {}
    rejeitados = []

    log, log_path = _configurar_logger()
    log.info(f"Pasta de entrada : {PASTA_ENTRADA.resolve()}")

    if not PASTA_ENTRADA.exists():
        log.error(f"Pasta '{PASTA_ENTRADA}' não encontrada.")
        return acumulador

    pdfs = sorted(PASTA_ENTRADA.glob("*.pdf"))
    log.info(f"PDFs encontrados : {len(pdfs)}")

    if not pdfs:
        log.warning("Nenhum PDF em 'saida/'. Encerrando.")
        return acumulador

    for i, pdf_path in enumerate(pdfs, start=1):
        log.info("")
        log.info(f"{'─' * 65}")
        log.info(f"[{i}/{len(pdfs)}] {pdf_path.name}  ({pdf_path.stat().st_size/1024:.1f} KB)")

        try:
            itens_dc, diverg = _double_check_pdf(str(pdf_path), log)
            aceitos = _filtrar_catalogo(itens_dc, pdf_path.name, log)
            _acumular(aceitos, pdf_path.name, log)
        except Exception as exc:
            log.error(f"  FALHA: {exc}", exc_info=True)

    pdf_out = _gerar_pdf(log)
    _log_relatorio(log, log_path, pdf_out)
    return acumulador


if __name__ == "__main__":
    processar_pasta_saida()
