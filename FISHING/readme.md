# 🎣🤖 Fishing Bot

> Automação inteligente para o minigame de pesca do FiveM — detecção visual em tempo real via OpenCV com movimentação de cursor de baixa latência (~120 fps).

---

## 📌 Sobre o Projeto

O **Fishing Bot** é um bot em Python desenvolvido para **automatizar o minigame de pesca em servidores FiveM**. Ele utiliza **visão computacional** para identificar o alvo cinza do minigame em tempo real na tela, mover o cursor automaticamente até ele e, ao detectar o fim da sessão, **relançar a pescaria de forma autônoma** com a sequência de teclas correta.

Ideal para farm contínuo de itens de pesca em servidores RP sem necessidade de interação manual.

---

## ✨ Funcionalidades

- 🖥️ **Captura de tela de alta performance** via `mss` com buffers pré-alocados (~120 fps)
- 🔍 **Detecção de alvo por visão computacional** — filtro de cor cinza neutro + análise de contornos com OpenCV
- 🖱️ **Movimentação de cursor via `SendInput`** (ctypes) — compatível com jogo em foco
- 🔄 **Relançamento automático** da pescaria ao detectar fim do minigame
- 🐛 **Modo debug** com janela de visualização em tempo real (máscara + contornos + alvo)
- ⌨️ **Atalhos de teclado** para ativar/desativar e encerrar sem fechar o jogo

---

## 🗂️ Estrutura do Projeto

```
Fishing/
│
├── 🔨 build.bat            # Script de build automatizado (limpa e gera o .exe)
└── 📄 fishing_bot.py       # Script principal — captura, detecção, controle e debug
```

### 📄 Descrição dos Arquivos

| Arquivo | Função |
|---|---|
| `fishing_bot.py` | Código-fonte principal — contém toda a lógica de captura, detecção, controle e debug |
| `build.bat` | Instalação completa e automatizada: instala todas as dependências via `pip`, limpa builds anteriores e gera `dist/fishing_bot.exe` via PyInstaller |

### 📄 Descrição dos Módulos Internos

| Função / Módulo | Responsabilidade |
|---|---|
| `capture()` | Captura a região do minigame via `mss` de forma thread-safe, sem alocação por frame |
| `build_mask()` | Gera máscara binária dos pixels cinza neutro usando operações NumPy totalmente in-place |
| `find_target()` | Analisa contornos da máscara e seleciona o melhor blob por score de circularidade × área |
| `bot_loop()` | Loop principal (~120 fps) que rastreia o alvo e aciona o relançamento automático |
| `fishing_again()` | Executa a sequência de teclas `E → 4 → E` para reiniciar a pescaria |
| `debug_loop()` | Exibe janela OpenCV com overlay da máscara, contornos e posição do alvo em tempo real |
| `_send_vk()` / `_press_key()` | Envia eventos de teclado via `SendInput` da WinAPI para garantir compatibilidade com o jogo |

---

## ⌨️ Controles

| Tecla | Ação |
|---|---|
| `Home` ou `L` | Ativar / Desativar o bot |
| `F9` | Encerrar o script completamente |
| `Q` *(modo debug)* | Fechar a janela de debug |

---

## ⚙️ Configurações

As principais constantes ficam no topo do arquivo e podem ser ajustadas conforme a resolução e o servidor:

| Parâmetro | Padrão | Descrição |
|---|---|---|
| `MINIGAME_REGION` | `left:1185, top:430, w:300, h:250` | Região da tela onde o minigame aparece |
| `GRAY_MAX_CHANNEL_DIFF` | `18` | Tolerância máxima entre canais BGR para considerar um pixel cinza |
| `GRAY_BRIGHTNESS_MIN` | `15` | Brilho mínimo do pixel alvo |
| `GRAY_BRIGHTNESS_MAX` | `92` | Brilho máximo do pixel alvo |
| `BLOB_AREA_MIN` | `15` | Área mínima do contorno detectado (px²) |
| `BLOB_AREA_MAX` | `4000` | Área máxima do contorno detectado (px²) |
| `LOOP_INTERVAL` | `0.005` | Intervalo entre frames (~120 fps) |
| `MAX_MISS_FRAMES` | `45` | Frames sem detecção antes de considerar o minigame encerrado |

> ⚠️ Se o minigame não estiver sendo detectado, ajuste `MINIGAME_REGION` para a posição correta na sua resolução de tela.

---

## 🚀 Como Usar

### Pré-requisitos

- Python **3.8+** instalado e disponível no `PATH`
- **Windows** (obrigatório — usa WinAPI via `ctypes` e `pywin32`)
- Resolução de tela configurada de acordo com `MINIGAME_REGION`

> 💡 Se for rodar pelo `.exe` gerado pelo `build.bat`, **não é necessário ter Python nem instalar nada** na máquina de destino.

### Execução — Modo Normal

```bash
python fishing_bot.py
```

1. Com o script rodando, entre no servidor FiveM e inicie uma pescaria manualmente.
2. Pressione **`Home`** (ou **`L`**) para ativar o bot.
3. O bot passará a rastrear o alvo e mover o cursor automaticamente.
4. Ao término de cada pescaria, o bot relançará sozinho com a sequência `E → 4 → E`.
5. Pressione **`Home`** novamente para pausar, ou **`F9`** para encerrar.

### Execução — Modo Debug

```bash
python fishing_bot.py --debug
```

Abre uma janela OpenCV em tempo real exibindo:
- **Verde** — pixels detectados como cinza neutro (máscara ativa)
- **Amarelo** — centróides de contornos dentro do intervalo de área válido
- **Azul** — centróides fora do intervalo (descartados)
- **Círculo verde grande** — alvo selecionado com sua posição (x, y)

Use este modo para calibrar `MINIGAME_REGION` e os parâmetros de cor na sua resolução.

---

## 📦 Gerar Executável (.exe)

O projeto inclui o `build.bat` que **faz tudo automaticamente em uma única etapa** — instala as dependências e gera o executável. Não é necessário instalar nada manualmente antes.

### Pré-requisito

Apenas ter o **Python 3.8+** instalado e disponível no `PATH` do sistema.

### Build

Dê um duplo clique no `build.bat` ou execute via terminal na raiz do projeto:

```bat
build.bat
```

O script executa automaticamente na seguinte ordem:
1. **Instala todas as dependências** via `pip` (opencv-python, numpy, pyautogui, mss, keyboard, pywin32, pyinstaller)
2. Remove as pastas `build/` e `dist/` antigas
3. Remove o `.spec` anterior
4. Gera o `fishing_bot.exe` em `dist/`

Ao final, o executável estará em `dist/fishing_bot.exe` — **standalone, sem necessidade de Python instalado** na máquina de destino.

> O executável gerado tem ~69 MB pois empacota todas as dependências (OpenCV, NumPy, etc.) em um único arquivo.

> ⚠️ O `.exe` pode ser detectado como falso positivo por alguns antivírus devido ao uso de `SendInput` e captura de tela — isso é comportamento esperado para automação Windows.

---

## 🛠️ Tecnologias Utilizadas

- **Python 3** — Linguagem principal
- **OpenCV (`cv2`)** — Detecção de contornos, morfologia e visualização debug
- **NumPy** — Operações vetorizadas in-place para filtragem de cor sem alocação heap
- **mss** — Captura de tela de alta performance thread-safe
- **pywin32 / ctypes** — Movimentação de cursor e envio de teclas via WinAPI (`SendInput`)
- **keyboard** — Registro de hotkeys globais
- **PyInstaller** — Empacotamento em executável standalone `.exe`

---

## ⚠️ Avisos

- Este script foi desenvolvido para uso pessoal em servidores FiveM que permitem automação.
- Use com responsabilidade e respeite as regras do servidor em que jogar.
- O bot **não bypassa anti-cheats** — funciona apenas via movimentação de cursor e teclas, como um usuário normal.
