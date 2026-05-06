# BotPO — Processador de Pedidos de Compra

Automação RPA em Python para extrair, consolidar e exportar dados de Pedidos de Compra a partir de emails Outlook (`.msg`).  
Disponível em duas interfaces: **web** (recomendada) e **CLI**.

---

## Como funciona

O pipeline executa três etapas em sequência:

| Etapa | O que faz | Saída |
|-------|-----------|-------|
| **1 — Extração de PDFs** | Abre cada `.msg` e extrai os anexos PDF | `temp_pdfs/` |
| **2 — Mapeamento de produtos** | Lê os PDFs e mapeia ID (7 dígitos) → Nome do produto | `output/mapeamento_produtos.csv` |
| **3 — Consolidação e Excel** | Cruza IDs, POs e quantidades; gera tabela pivot estilizada | `output/consolidado_pedidos.xlsx` |

---

## Interface Web (recomendada)

```bash
python app.py
```

O navegador abre automaticamente em `http://localhost:5000`.

**Fluxo:**
1. Arraste ou selecione os arquivos `.msg`
2. Clique em **Processar** — o progresso de cada etapa aparece em tempo real
3. Baixe o relatório Excel ao final

> Não é necessário configurar nada. O servidor roda localmente na porta 5000.

---

## Interface CLI

```bash
python main.py
```

Ao final do pipeline, o sistema pergunta se deseja arquivar o lote em `archives/`.

---

## Estrutura do projeto

```
PO/
├── app.py                        # Servidor web Flask (interface principal)
├── main.py                       # Orquestrador CLI
├── extractor_msg.py              # Etapa 1: extrai PDFs dos .msg
├── extractor_ids.py              # Etapa 2: mapeia ID → Nome de produto
├── extractor_quantidades.py      # Etapa 3: consolida quantidades e gera Excel
├── archiver.py                   # Arquivamento de lotes (CLI)
├── templates/
│   └── index.html                # Página única da interface web
├── input_msgs/                   # ← coloque os .msg aqui (ignorado pelo git)
├── temp_pdfs/                    # PDFs extraídos (ignorado pelo git)
└── output/                       # Relatórios gerados (ignorado pelo git)
```

---

## Dependências

```bash
python -m pip install flask extract-msg PyMuPDF pandas openpyxl tqdm
```

| Pacote | Uso |
|--------|-----|
| `flask` | Servidor web da interface |
| `extract-msg` | Leitura de arquivos `.msg` do Outlook |
| `PyMuPDF` | Extração de texto dos PDFs |
| `pandas` | Pivot table e consolidação de dados |
| `openpyxl` | Geração e estilização do Excel |
| `tqdm` | Barras de progresso no CLI |

---

## Formato do Excel gerado

| Coluna | Conteúdo |
|--------|----------|
| **PRODUTO** | Nome do produto |
| **TOTAL** | Soma de todos os POs |
| **PO-XXXXX** | Quantidade por Pedido de Compra |

Produtos sem quantidade em determinado PO recebem valor `0`.  
Cabeçalho azul escuro, linhas alternadas, colunas com largura ajustada e painéis congelados.

---

## Padrões reconhecidos nos PDFs

Os PDFs devem conter:
- **Número do PO:** `Nº do Pedido: 12345`
- **ID do produto:** código de 7 dígitos (ex: `7004434`)
- **Quantidade:** linha com `dd.mm.aaaa  QTD  UN`

O mapeamento ID → Nome suporta três formatos de layout (bloco, inline com traço e tabela).
