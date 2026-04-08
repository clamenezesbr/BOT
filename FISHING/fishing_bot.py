"""
FiveM Fishing Minigame Bot v6
==============================
Dependências:
    pip install opencv-python numpy pyautogui mss keyboard pywin32

Controles:
    Home  → Ativar / Desativar bot
    F9  → Encerrar script

Uso:
    python fishing_bot.py           # modo normal
    python fishing_bot.py --debug   # janela de visualização em tempo real
"""

import time
import threading
import ctypes
from ctypes import wintypes
import numpy as np
import cv2
import pyautogui
import win32api
import pywintypes
import mss
import keyboard

# ══════════════════════════════════════════════════════════════
#  CONFIGURAÇÕES
# ══════════════════════════════════════════════════════════════

MINIGAME_REGION = {"left": 1185, "top": 430, "width": 300, "height": 250}

# Detecção BGR — cinza escuro neutro (R ≈ G ≈ B, brilho médio-baixo)
GRAY_MAX_CHANNEL_DIFF = 18
GRAY_BRIGHTNESS_MIN   = 15
GRAY_BRIGHTNESS_MAX   = 92

# Área do blob em pixels (ROI real 300x250)
BLOB_AREA_MIN = 15
BLOB_AREA_MAX = 4000

# ~120fps = 0.005s por frame
LOOP_INTERVAL = 0.005

# Frames sem detectar antes de considerar o minigame encerrado
MAX_MISS_FRAMES = 45

# ══════════════════════════════════════════════════════════════

bot_active  = False
bot_running = True

# Pré-computados para evitar recriação a cada frame
_DILATE_KERNEL      = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (5, 5))
_BRIGHTNESS_SUM_MIN = GRAY_BRIGHTNESS_MIN * 3
_BRIGHTNESS_SUM_MAX = GRAY_BRIGHTNESS_MAX * 3

# Buffers pré-alocados — reutilizados a cada frame, eliminam ~2.4 MB de alocação/frame
# (300×250 px a ~120fps = ~288 MB/s de pressão no heap evitados)
_H = MINIGAME_REGION["height"]
_W = MINIGAME_REGION["width"]
_frame_bgr  = np.empty((_H, _W, 3), dtype=np.uint8)   # captura BGR contígua
_work_i16   = np.empty((_H, _W, 3), dtype=np.int16)   # canais uint8→int16
_diff_tmp   = np.empty((_H, _W),    dtype=np.int16)   # max_diff inter-canais
_bsum       = np.empty((_H, _W),    dtype=np.int16)   # soma de brilho (b+g+r)
_mask_bool  = np.empty((_H, _W),    dtype=np.bool_)   # máscara booleana
_cond_buf   = np.empty((_H, _W),    dtype=np.bool_)   # condição auxiliar
_mask_buf   = np.empty((_H, _W),    dtype=np.uint8)   # máscara de saída (0/255)


def toggle_bot():
    global bot_active
    bot_active = not bot_active
    print(f"[BOT] {'ATIVADO ✓' if bot_active else 'DESATIVADO ✗'}")


def quit_bot():
    global bot_running, bot_active
    bot_running = False
    bot_active  = False
    print("[BOT] Encerrando...")


# ──────────────────────────────────────────────────────────────
#  CAPTURA  (thread-safe no Windows)
# ──────────────────────────────────────────────────────────────

_local = threading.local()

def capture() -> np.ndarray:
    try:
        if not hasattr(_local, "sct"):
            _local.sct = mss.mss()
        shot = _local.sct.grab(MINIGAME_REGION)
        # Copia BGR do raw BGRA para o buffer pré-alocado contíguo (sem cvtColor, sem alloc)
        raw = np.frombuffer(shot.raw, dtype=np.uint8).reshape((_H, _W, 4))
        np.copyto(_frame_bgr, raw[:, :, :3])
        return _frame_bgr
    except Exception as e:
        print(f"[ERRO] Falha ao capturar a tela — verifique se o monitor está ativo e a resolucao nao mudou. ({e})")
        return _frame_bgr


# ──────────────────────────────────────────────────────────────
#  DETECÇÃO
# ──────────────────────────────────────────────────────────────

def build_mask(frame: np.ndarray) -> np.ndarray:
    # uint8 → int16 sem alocar: copia direto no buffer pré-alocado
    np.copyto(_work_i16, frame, casting='unsafe')
    b = _work_i16[:, :, 0]
    g = _work_i16[:, :, 1]
    r = _work_i16[:, :, 2]

    # max(|r-g|, |r-b|, |g-b|) — totalmente in-place, zero alocação heap
    np.subtract(r, g, out=_diff_tmp);  np.abs(_diff_tmp, out=_diff_tmp)
    np.subtract(r, b, out=_bsum);      np.abs(_bsum,     out=_bsum)
    np.maximum(_diff_tmp, _bsum, out=_diff_tmp)
    np.subtract(g, b, out=_bsum);      np.abs(_bsum,     out=_bsum)
    np.maximum(_diff_tmp, _bsum, out=_diff_tmp)

    # b + g + r — in-place no _bsum
    np.add(b, g, out=_bsum);  np.add(_bsum, r, out=_bsum)

    # AND das três condições sem temporários heap
    np.less(_diff_tmp, GRAY_MAX_CHANNEL_DIFF,    out=_mask_bool)
    np.greater_equal(_bsum, _BRIGHTNESS_SUM_MIN, out=_cond_buf)
    np.bitwise_and(_mask_bool, _cond_buf, out=_mask_bool)
    np.less_equal(_bsum,    _BRIGHTNESS_SUM_MAX, out=_cond_buf)
    np.bitwise_and(_mask_bool, _cond_buf, out=_mask_bool)

    # bool (0/1) → uint8 (0/255) no buffer de saída; dilate in-place
    np.multiply(_mask_bool, np.uint8(255), out=_mask_buf)
    cv2.dilate(_mask_buf, _DILATE_KERNEL, dst=_mask_buf, iterations=1)
    return _mask_buf


def find_target(frame: np.ndarray):
    mask = build_mask(frame)
    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not contours:
        return None, mask, contours
    best       = None
    best_score = -1
    for c in contours:
        area = cv2.contourArea(c)
        if not (BLOB_AREA_MIN < area < BLOB_AREA_MAX):
            continue
        perimeter = cv2.arcLength(c, True)
        if perimeter == 0:
            continue
        circularity = (4 * np.pi * area) / (perimeter ** 2)
        score = circularity * area
        if score > best_score:
            best_score = score
            best = c
    if best is None:
        return None, mask, contours
    M = cv2.moments(best)
    if M["m00"] == 0:
        return None, mask, contours
    return (int(M["m10"] / M["m00"]), int(M["m01"] / M["m00"])), mask, contours


def roi_to_screen(rx: int, ry: int):
    return MINIGAME_REGION["left"] + rx, MINIGAME_REGION["top"] + ry


# ──────────────────────────────────────────────────────────────
#  SENDINPUT (ctypes) — funciona com jogo em foco
# ──────────────────────────────────────────────────────────────

class _KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk",         wintypes.WORD),
        ("wScan",       wintypes.WORD),
        ("dwFlags",     wintypes.DWORD),
        ("time",        wintypes.DWORD),
        ("dwExtraInfo", ctypes.c_size_t),   # ULONG_PTR — 8 bytes em 64-bit
    ]

class _MOUSEINPUT(ctypes.Structure):        # necessário para dimensionar a union corretamente
    _fields_ = [
        ("dx",          wintypes.LONG),
        ("dy",          wintypes.LONG),
        ("mouseData",   wintypes.DWORD),
        ("dwFlags",     wintypes.DWORD),
        ("time",        wintypes.DWORD),
        ("dwExtraInfo", ctypes.c_size_t),
    ]

class _INPUT_UNION(ctypes.Union):
    _fields_ = [("mi", _MOUSEINPUT), ("ki", _KEYBDINPUT)]

class _INPUT(ctypes.Structure):
    _fields_ = [("type", wintypes.DWORD), ("_input", _INPUT_UNION)]

_KEYEVENTF_KEYUP  = 0x0002
_INPUT_KEYBOARD   = 1
_user32_SendInput = ctypes.windll.user32.SendInput


def _send_vk(vk: int, flags: int) -> None:
    inp = _INPUT(type=_INPUT_KEYBOARD)
    inp._input.ki = _KEYBDINPUT(wVk=vk, wScan=0, dwFlags=flags, time=0, dwExtraInfo=0)
    result = _user32_SendInput(1, ctypes.byref(inp), ctypes.sizeof(_INPUT))
    if result == 0:
        print(f"[ERRO] SendInput falhou ao enviar tecla VK=0x{vk:02X} — o jogo pode estar bloqueando inputs externos.")


def _press_key(vk: int) -> None:
    _send_vk(vk, 0)
    time.sleep(0.05)
    _send_vk(vk, _KEYEVENTF_KEYUP)


def fishing_again() -> None:
    time.sleep(4.0)
    _press_key(0x45)   # E

    time.sleep(3.5)
    _press_key(0x34)   # 4

    time.sleep(1.7)
    _press_key(0x45)   # E


# ──────────────────────────────────────────────────────────────
#  LOOP PRINCIPAL
# ──────────────────────────────────────────────────────────────

def bot_loop():
    pyautogui.FAILSAFE = False
    misses       = 0
    was_tracking = False

    print("[BOT] Pronto!  HOME = ativar/desativar  |  F9 = sair")

    while bot_running:
        if not bot_active:
            time.sleep(0.05)
            misses       = 0
            was_tracking = False
            continue

        t0 = time.perf_counter()

        frame        = capture()
        target, _, _ = find_target(frame)

        if target:
            misses       = 0
            was_tracking = True
            sx, sy = roi_to_screen(*target)
            try:
                win32api.SetCursorPos((sx, sy))
            except pywintypes.error:
                print(f"[ERRO] Nao foi possivel mover o mouse para ({sx}, {sy}) — verifique se o mouse esta conectado e nao ha software bloqueando o controle do cursor.")
        else:
            misses += 1
            if misses == MAX_MISS_FRAMES and was_tracking:
                print("[BOT] Minigame encerrado — relançando...")
                was_tracking = False
                misses       = 0
                fishing_again()

        elapsed = time.perf_counter() - t0
        sleep_t = LOOP_INTERVAL - elapsed
        if sleep_t > 0:
            time.sleep(sleep_t)


# ──────────────────────────────────────────────────────────────
#  MODO DEBUG
# ──────────────────────────────────────────────────────────────

def debug_loop():
    print("[DEBUG] Janela aberta. Pressione Q para fechar.")
    scale = 3

    while bot_running:
        frame                  = capture()
        target, mask, all_ctrs = find_target(frame)

        overlay = frame.copy()
        overlay[mask > 0] = [0, 200, 0]
        display = cv2.addWeighted(frame, 0.5, overlay, 0.5, 0)

        for c in all_ctrs:
            area = cv2.contourArea(c)
            if area < 5:
                continue
            M = cv2.moments(c)
            if M["m00"] == 0:
                continue
            cx = int(M["m10"] / M["m00"])
            cy = int(M["m01"] / M["m00"])
            in_range = BLOB_AREA_MIN < area < BLOB_AREA_MAX
            color = (0, 255, 255) if in_range else (0, 80, 200)
            cv2.circle(display, (cx, cy), 4, color, -1)
            cv2.putText(display, f"{area:.0f}", (cx+4, cy),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.32, color, 1)

        if target:
            cv2.circle(display, target, 14, (0, 255, 0), 2)
            cv2.circle(display, target,  3, (0, 255, 0), -1)
            cv2.putText(display, f"ALVO {target}", (4, 15),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.45, (0, 255, 0), 1)
        else:
            cv2.putText(display, "sem alvo", (4, 15),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.45, (0, 60, 255), 1)

        big = cv2.resize(display, (display.shape[1]*scale, display.shape[0]*scale),
                         interpolation=cv2.INTER_NEAREST)
        cv2.imshow("Fishing Bot v6 DEBUG", big)
        if cv2.waitKey(1) & 0xFF == ord("q"):
            break

    cv2.destroyAllWindows()


# ──────────────────────────────────────────────────────────────
#  ENTRY POINT (Teclas)
# ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys

    keyboard.add_hotkey("home", togglllleee_bot)
    keyboard.add_hotkey("l", toggle_bot)
    keyboard.add_hotkey("f9", quit_bot)

    t = threading.Thread(target=bot_loop, daemon=True)
    t.start()

    if "--debug" in sys.argv:
        debug_loop()
    else:
        while bot_running:
            time.sleep(0.5)

    print("[BOT] Encerrado.")
