#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Simple Live Translator with WordFlash and skip/re-display behavior.
This is the full script with focus-preserving show() and corrected on_redisplay_word().
"""

import os
import re
import csv
import json
import time
import threading
from collections import OrderedDict
from datetime import datetime
try:
    from zoneinfo import ZoneInfo
except Exception:
    ZoneInfo = None

import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.scrolledtext import ScrolledText
import tkinter.font as tkfont
from queue import Queue, Empty
from threading import Lock

# Pillow for screenshot (optional)
try:
    from PIL import ImageGrab, Image
except Exception:
    ImageGrab = None
    Image = None

# ---------- lmstudio ----------
try:
    import lmstudio as lms
except Exception:
    lms = None

# ---------- Config / Disabled langs ----------
DISABLED_LANGS = {"es", "fr"}  # スペイン語とフランス語を無効化
DEFAULT_CONFIG = {
    "SERVER_API_HOST": "localhost:1234",
    "MODEL_NAME": "google/gemma-3n-e4b",
    #"MODEL_NAME": "openai/gpt-oss-20b",
    "TEMPERATURE": 0.1,
    "MAX_TOKENS": 0,  # 0=auto → omit maxTokens (server default)
}
CONFIG_PATH = os.path.join(os.path.dirname(__file__), "simple_live_translator.config.json")

# ---------- colors ----------
COLOR_INPUT = "#FFFFFF"
COLOR_JA    = "#FFD700"
COLOR_EN    = "#00FFFF"
COLOR_ES    = "#777777"   # disabled dim
COLOR_FR    = "#777777"   # disabled dim
COLOR_QA    = "#FFD700"
COLOR_DIM   = "#AAAAAA"
BG_DISPLAY  = "#000000"

# ---------- prompts ----------
def prompt_ja(text: str) -> str: return f"簡潔に日本語に訳した文だけ記載してください。\n「{text}」"
def prompt_en(text: str) -> str: return f"簡潔に英語に訳した文だけ記載してください。\n「{text}」"
def prompt_es(text: str) -> str: return f"簡潔にスペイン語に訳した文だけ記載してください。\n「{text}」"
def prompt_fr(text: str) -> str: return f"簡潔にフランス語に訳した文だけ記載してください。\n「{text}」"

def prompt_word_en(word: str) -> str:
    return f"簡潔にこのドイツ語に最も近い英語をセミコロン区切りの形式で3つ列挙してください。\n「{word}」"
def prompt_word_ja(word: str) -> str:
    return f"簡潔にこのドイツ語に最も近い日本語語をセミコロン区切りの形式で3つ列挙してください。\n「{word}」"

# ---------- utils ----------
PUNCT_STRIP = ".,!?;:\"“”„()[]{}<>/\\|—–-+*=~_^`…，。！？；：『』「」【】（）«»"

def tz_now_jst():
    if ZoneInfo is not None:
        try: return datetime.now(ZoneInfo("Asia/Tokyo"))
        except Exception: pass
    return datetime.now()

def ensure_dir(p: str):
    os.makedirs(p, exist_ok=True)

def write_text(path: str, text: str):
    ensure_dir(os.path.dirname(path))
    with open(path, "w", encoding="utf-8", newline="") as f:
        f.write("" if text is None else str(text))

def append_csv_row(path: str, row: list, header: list = None):
    ensure_dir(os.path.dirname(path))
    need_header = not os.path.exists(path)
    with open(path, "a", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        if need_header and header: w.writerow(header)
        w.writerow(row)

def extract_final_message_only(s: str) -> str:
    tokens = ["final<|message|>", "<|channel|>final<|message|>", "<|start|>assistant<|channel|>final<|message|>"]
    last_idx = -1; last_tok = None
    for t in tokens:
        i = s.rfind(t)
        if i > last_idx: last_idx, last_tok = i, t
    out = s[last_idx + len(last_tok):] if last_idx != -1 else s
    out = re.sub(r'^\s*(?:<\|[^|]+\|>)+', '', out, flags=re.S)
    return out.strip()

def normalize_word(w: str) -> str:
    w = w.strip(PUNCT_STRIP)
    if not w: return ""
    return w[:1].upper() + w[1:].lower()

def tokenize_german(text: str):
    seen = set(); ordered = []
    for tok in text.split():
        n = normalize_word(tok)
        if n and n not in seen:
            seen.add(n); ordered.append(n)
    return ordered

def sanitize_filename(name: str) -> str:
    # Replace problematic chars with _
    return re.sub(r'[^A-Za-z0-9_\-\.]', '_', name)

# ---------- word.csv ----------
WORD_CSV_PATH = os.path.join(os.path.dirname(__file__), "log", "word.csv")
# New header includes 'skip' as 5th column (0/1)
WORD_HEADER   = ["word", "en", "ja", "count", "skip"]

def load_word_db():
    """
    Supports both old 4-column format and new 5-column format.
    Returns dict: {word: {"en":..., "ja":..., "count":int, "skip":0/1}}
    """
    db = {}
    if not os.path.exists(WORD_CSV_PATH): return db
    with open(WORD_CSV_PATH, "r", encoding="utf-8", newline="") as f:
        rows = list(csv.reader(f))
    if not rows: return db
    header = [c.strip().lower() for c in rows[0]]
    # 2-column minimal header ["word","count"] legacy
    if header == ["word","count"]:
        for row in rows[1:]:
            if not row: continue
            w = row[0].strip()
            cnt = int(row[1]) if len(row)>1 and row[1].isdigit() else 0
            db[w] = {"en":"", "ja":"", "count":cnt, "skip":0}
    # old 4-col format
    elif header[:4] == ["word","en","ja","count"]:
        for row in rows[1:]:
            if not row: continue
            w = row[0].strip()
            en = row[1].strip() if len(row)>1 else ""
            ja = row[2].strip() if len(row)>2 else ""
            try: cnt = int(row[3]) if len(row)>3 else 0
            except Exception: cnt = 0
            # if skip column exists use it, else 0
            try: sk = int(row[4]) if len(row)>4 and row[4] in ("0","1") else 0
            except Exception: sk = 0
            db[w] = {"en":en, "ja":ja, "count":cnt, "skip":sk}
    # new 5-col format
    elif header[:5] == ["word","en","ja","count","skip"]:
        for row in rows[1:]:
            if not row: continue
            w = row[0].strip()
            en = row[1].strip() if len(row)>1 else ""
            ja = row[2].strip() if len(row)>2 else ""
            try: cnt = int(row[3]) if len(row)>3 else 0
            except Exception: cnt = 0
            try: sk = int(row[4]) if len(row)>4 and str(row[4]).strip().isdigit() else 0
            except Exception: sk = 0
            db[w] = {"en":en, "ja":ja, "count":cnt, "skip":sk}
    else:
        # fallback: try to parse rows len>=2
        for row in rows[1:]:
            if not row: continue
            w = row[0].strip()
            en = row[1].strip() if len(row)>1 else ""
            ja = row[2].strip() if len(row)>2 else ""
            try: cnt = int(row[3]) if len(row)>3 else 0
            except Exception: cnt = 0
            try: sk = int(row[4]) if len(row)>4 and str(row[4]).strip().isdigit() else 0
            except Exception: sk = 0
            db[w] = {"en":en, "ja":ja, "count":cnt, "skip":sk}
    return db

def save_word_db(db: dict):
    ensure_dir(os.path.dirname(WORD_CSV_PATH))
    with open(WORD_CSV_PATH, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f); w.writerow(WORD_HEADER)
        for word, info in sorted(db.items(), key=lambda x: (-x[1].get("count",0), x[0])):
            skip_val = int(info.get("skip",0))
            w.writerow([word, info.get("en",""), info.get("ja",""), int(info.get("count",0)), skip_val])

# ---------- Display window ----------
class DisplayWindow(tk.Toplevel):
    BOX_SIZE = 24
    BLINK_MS = 500
    def __init__(self, master):
        super().__init__(master)
        self.title("Display")
        self.configure(bg=BG_DISPLAY)
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        # set initial height to 1/4 of master's height if possible
        try:
            master.update_idletasks()
            w = master.winfo_width() or 1200
            h = master.winfo_height() or 800
            self.geometry(f"{w}x{max(120, int(h//4))}")
        except Exception:
            try: self.geometry("1000x200")
            except Exception: pass

        base = tkfont.nametofont("TkDefaultFont")
        self.big_bold   = tkfont.Font(family=base.actual("family"), size=max(16, base.actual("size") * 2), weight="bold")
        self.small_bold = tkfont.Font(family=base.actual("family"), size=max(12, int(base.actual("size") * 1.2)), weight="bold")

        # status bar (英→日 ; 西/仏 kept but dimmed when disabled)
        bar = tk.Frame(self, bg=BG_DISPLAY)
        bar.pack(fill=tk.X, padx=8, pady=(6, 0))

        def make_box(parent, char, fg, disabled=False):
            c = tk.Canvas(parent, width=self.BOX_SIZE, height=self.BOX_SIZE, bg=BG_DISPLAY, highlightthickness=0)
            outline = COLOR_DIM if disabled else fg
            c.create_rectangle(2, 2, self.BOX_SIZE-2, self.BOX_SIZE-2, outline=outline, width=2)
            tid = c.create_text(self.BOX_SIZE//2, self.BOX_SIZE//2, text=char, fill=outline, font=self.small_bold, state="hidden")
            c.pack(side=tk.LEFT, padx=6)
            return c, tid

        # status dict (no per-language queue-count label anymore)
        self.status = OrderedDict([
            ("en", {"canvas":None,"text_id":None,"blink":False,"after_id":None,"fg":COLOR_EN,"busy":0}),
            ("ja", {"canvas":None,"text_id":None,"blink":False,"after_id":None,"fg":COLOR_JA,"busy":0}),
            ("es", {"canvas":None,"text_id":None,"blink":False,"after_id":None,"fg":COLOR_ES,"busy":0}),
            ("fr", {"canvas":None,"text_id":None,"blink":False,"after_id":None,"fg":COLOR_FR,"busy":0}),
        ])

        # create the language boxes (ES/FR visually dimmed if disabled)
        for code, char in [("en","英"), ("ja","日"), ("es","西"), ("fr","仏")]:
            disabled = (code in DISABLED_LANGS)
            c, tid = make_box(bar, char, self.status[code]["fg"], disabled=disabled)
            self.status[code]["canvas"] = c
            self.status[code]["text_id"] = tid

        # single global queue-count label (white) placed at right side of the bar
        self.queue_counts_label = tk.Label(bar, text="S 0  W 0", bg=BG_DISPLAY, fg="#FFFFFF", font=self.small_bold)
        # pack to the right so it is visually separated from the language boxes
        self.queue_counts_label.pack(side=tk.RIGHT, padx=(6,8))

        # main text
        self.text = ScrolledText(self, height=24, wrap=tk.WORD, bg=BG_DISPLAY, fg=COLOR_INPUT, insertbackground="#FFFFFF")
        self.text.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        self.text.configure(font=self.big_bold)
        self.text.tag_config("in", foreground=COLOR_INPUT, font=self.big_bold)
        self.text.tag_config("ja", foreground=COLOR_JA,    font=self.big_bold)
        self.text.tag_config("en", foreground=COLOR_EN,    font=self.big_bold)
        self.text.tag_config("es", foreground=COLOR_ES,    font=self.big_bold)
        self.text.tag_config("fr", foreground=COLOR_FR,    font=self.big_bold)
        self.text.tag_config("qa_in",  foreground=COLOR_INPUT, font=self.small_bold)
        self.text.tag_config("qa_out", foreground=COLOR_QA,    font=self.small_bold)
        self.text.tag_config("dim", foreground=COLOR_DIM)

    # status control
    def start_status(self, lang: str):
        if lang in DISABLED_LANGS: return  # no-op for disabled languages
        info = self.status.get(lang);
        if not info: return
        info["busy"] += 1
        if info["busy"] == 1:
            can, tid = info["canvas"], info["text_id"]
            can.itemconfigure(tid, state="normal")
            info["blink"] = True
            def _tick():
                if not info["blink"]: return
                state = can.itemcget(tid, "state")
                can.itemconfigure(tid, state=("hidden" if state == "normal" else "normal"))
                info["after_id"] = self.after(self.BLINK_MS, _tick)
            _tick()

    def done_status(self, lang: str):
        if lang in DISABLED_LANGS: return
        info = self.status.get(lang);
        if not info: return
        info["busy"] = max(0, info["busy"] - 1)
        if info["busy"] == 0:
            if info["after_id"] is not None:
                try: self.after_cancel(info["after_id"])
                except Exception: pass
                info["after_id"] = None
            info["blink"] = False
            can, tid = info["canvas"], info["text_id"]
            can.itemconfigure(tid, state="normal")  # keep visible

    def update_queue_counts(self, sentence_q_size: int, word_q_size: int):
        # update single global queue-count label (called periodically from App)
        try:
            self.queue_counts_label.config(text=f"S {sentence_q_size}  W {word_q_size}")
        except Exception:
            pass

    def on_close(self): self.withdraw()

    def show(self):
        # フォーカスを奪わずにウィンドウを表示／復元する
        try:
            if self.state() == "withdrawn":
                self.deiconify()
            # 描画アップデートは行うがフォーカスは奪わない
            self.update_idletasks()
        except Exception:
            pass

    def append_line(self, line: str, tag: str = "in"):
        self.text.insert(tk.END, line + "\n", tag); self.text.see(tk.END)

# ---------- Word Flash (marquee for EN/JA) ----------
class WordFlash(tk.Toplevel):
    SCROLL_MS = 25
    SCROLL_DX = 2
    WIDTH = 520

    def __init__(self, master):
        super().__init__(master)
        self.title("Word Flash")
        self.configure(bg=BG_DISPLAY)
        self.geometry(f"{self.WIDTH}x220+80+80")
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        base = tkfont.nametofont("TkDefaultFont")
        self.bold = tkfont.Font(family=base.actual("family"), size=max(16, int(base.actual("size")*1.6)), weight="bold")

        # count & German (Labels)
        self.lbl_count = tk.Label(self, text="", bg=BG_DISPLAY, fg=COLOR_INPUT, font=self.bold)
        self.lbl_de    = tk.Label(self, text="", bg=BG_DISPLAY, fg=COLOR_INPUT, font=self.bold)

        # EN/JA scroll areas (Canvas)
        self.en_canvas = tk.Canvas(self, width=self.WIDTH-40, height=40, bg=BG_DISPLAY, highlightthickness=0)
        self.ja_canvas = tk.Canvas(self, width=self.WIDTH-40, height=40, bg=BG_DISPLAY, highlightthickness=0)

        self.lbl_count.pack(pady=(18,6))
        self.lbl_de.pack(pady=2)
        self.en_canvas.pack(pady=2)
        self.ja_canvas.pack(pady=2)

        # text ids and timers
        self._en_text_id = None
        self._ja_text_id = None
        self._en_timer = None
        self._ja_timer = None

        self.font_en = self.bold
        self.font_ja = self.bold

    def on_close(self): self.withdraw()

    def show(self):
        # フォーカスを奪わずに表示
        try:
            if self.state() == "withdrawn":
                self.deiconify()
            self.update_idletasks()
        except Exception:
            pass

    def _clear_marquee(self, canvas, which):
        if which == "en" and self._en_timer:
            try: self.after_cancel(self._en_timer)
            except Exception: pass
            self._en_timer = None
        if which == "ja" and self._ja_timer:
            try: self.after_cancel(self._ja_timer)
            except Exception: pass
            self._ja_timer = None
        canvas.delete("all")

    def _start_marquee(self, canvas, text, color, font, which):
        self._clear_marquee(canvas, which)
        # draw once to measure
        tid = canvas.create_text(0, 20, text=text, anchor="w", fill=color, font=font)
        bbox = canvas.bbox(tid)
        text_w = (bbox[2]-bbox[0]) if bbox else 0
        area_w = int(canvas.cget("width"))

        if text_w <= area_w:
            # center static
            canvas.coords(tid, (area_w-text_w)//2, 20)
            if which == "en": self._en_text_id = tid
            else: self._ja_text_id = tid
            return

        # place starting at right edge + small gap
        x = area_w
        canvas.coords(tid, x, 20)
        if which == "en": self._en_text_id = tid
        else: self._ja_text_id = tid

        def tick():
            nonlocal x
            x -= self.SCROLL_DX
            canvas.coords(tid, x, 20)
            if x + text_w < 0:
                x = area_w  # reset to right
            timer = self.after(self.SCROLL_MS, tick)
            if which == "en": self._en_timer = timer
            else: self._ja_timer = timer
        tick()

    def show_word(self, count:int, de:str, en:str, ja:str, starred:bool=False):
        # show window and update labels
        self.show()
        star = "★" if starred else ""
        self.lbl_count.config(text=f"登場回数: {star}{count}")
        self.lbl_de.config(text=f"{de}")
        self._start_marquee(self.en_canvas, en or "", COLOR_EN, self.font_en, "en")
        self._start_marquee(self.ja_canvas, ja or "", COLOR_JA, self.font_ja, "ja")

# ---------- Main app ----------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Simple Live Translator (LM Studio) - EN/JA only")
        # base size
        self.geometry("1000x820")
        self.minsize(920, 740)

        self.config_dict = self._load_config()
        self._init_client()
        self._build_ui()

        self.display = DisplayWindow(self); self.display.show()
        self.wordhud = WordFlash(self); self.wordhud.show()  # ← 起動時から表示

        self.pending = {}  # ts14 -> {"de":str, "ja":..., "en":..., "es":..., "fr":...}

        # --- Queues & workers ---
        self.sentence_q = Queue()
        self.word_q = Queue()
        self._current_sentence = None
        self._current_word = None
        self._cur_lock = Lock()
        # for wordflash buffer control
        self._last_wordflash_time = 0.0

        threading.Thread(target=self._sentence_worker, daemon=True).start()
        threading.Thread(target=self._word_worker_queue, daemon=True).start()

        # periodic UI updater for queue status
        self._update_queue_labels()

    # --- config ---
    def _load_config(self):
        cfg = DEFAULT_CONFIG.copy()
        if os.path.exists(CONFIG_PATH):
            try:
                with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                    disk = json.load(f); cfg.update(disk)
            except Exception: pass
        return cfg
    def _save_config(self):
        try:
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(self.config_dict, f, ensure_ascii=False, indent=2)
        except Exception: pass

    # --- lmstudio ---
    def _init_client(self):
        self.client_ok = False
        if lms is None: return
        try:
            server = self.config_dict.get("SERVER_API_HOST", DEFAULT_CONFIG["SERVER_API_HOST"]).strip()
            if server: lms.configure_default_client(server)
            self.client_ok = True
        except Exception as e:
            print("LM Studio client init failed:", e); self.client_ok = False

    # --- UI ---
    def _build_ui(self):
        base = tkfont.nametofont("TkDefaultFont")
        self.big_bold   = tkfont.Font(family=base.actual("family"), size=max(16, base.actual("size") * 2), weight="bold")
        self.small_bold = tkfont.Font(family=base.actual("family"), size=max(12, int(base.actual("size") * 1.2)), weight="bold")

        top = ttk.LabelFrame(self, text="Connection")
        top.pack(fill=tk.X, padx=10, pady=8)

        ttk.Label(top, text="Server:").grid(row=0, column=0, sticky="w", padx=6)
        self.var_server = tk.StringVar(value=self.config_dict.get("SERVER_API_HOST"))
        ttk.Entry(top, textvariable=self.var_server, width=18).grid(row=0, column=1, sticky="w")

        ttk.Label(top, text="Model:").grid(row=0, column=2, sticky="w", padx=(12,4))
        self.var_model = tk.StringVar(value=self.config_dict.get("MODEL_NAME"))
        ttk.Entry(top, textvariable=self.var_model, width=28).grid(row=0, column=3, sticky="w")

        ttk.Label(top, text="Temp:").grid(row=0, column=4, sticky="e", padx=(12,4))
        self.var_temp = tk.DoubleVar(value=float(self.config_dict.get("TEMPERATURE")))
        ttk.Spinbox(top, from_=0.0, to=1.5, increment=0.05, textvariable=self.var_temp, width=6).grid(row=0, column=5, sticky="w")

        ttk.Label(top, text="MaxTokens (0=auto):").grid(row=0, column=6, sticky="e", padx=(12,4))
        self.var_maxtok = tk.IntVar(value=int(self.config_dict.get("MAX_TOKENS", 0) or 0))
        ttk.Spinbox(top, from_=0, to=32768, increment=64, textvariable=self.var_maxtok, width=8).grid(row=0, column=7, sticky="w")

        ttk.Button(top, text="Save", command=self.on_save).grid(row=0, column=8, padx=8)
        ttk.Button(top, text="Reconnect", command=self.on_reconnect).grid(row=0, column=9, padx=2)

        # --- MOVED: Show Display/WordFlash buttons to next row so they don't overflow ---
        ttk.Button(top, text="Show Display Window", command=self.on_show_display).grid(row=1, column=0, columnspan=2, sticky="w", padx=(6,4), pady=(6,0))
        ttk.Button(top, text="Show Word Flash Window", command=self.on_show_wordflash).grid(row=1, column=2, columnspan=2, sticky="w", padx=(6,4), pady=(6,0))

        # 翻訳 (EN/JA only)
        frm_in = ttk.LabelFrame(self, text="German → EN/JA (sequenced EN→JA) — ES/FR disabled")
        frm_in.pack(fill=tk.BOTH, expand=False, padx=10, pady=(6, 4))
        self.txt_in = ScrolledText(frm_in, height=4, wrap=tk.WORD)
        self.txt_in.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        self.txt_in.configure(font=self.big_bold)
        self.txt_in.bind("<Control-Return>", lambda e: self.on_send_translate())

        btns1 = ttk.Frame(self)
        btns1.pack(fill=tk.X, padx=10)
        ttk.Button(btns1, text="Translate (Ctrl+Enter)", command=self.on_send_translate).pack(side=tk.LEFT, padx=4)

        # Queue / progress display
        qfrm = ttk.Frame(self)
        qfrm.pack(fill=tk.X, padx=10, pady=(6,4))
        ttk.Label(qfrm, text="Sentence Queue:").grid(row=0, column=0, sticky="w")
        self.lbl_sentence_queue = ttk.Label(qfrm, text="0")
        self.lbl_sentence_queue.grid(row=0, column=1, sticky="w", padx=(6,20))
        ttk.Label(qfrm, text="Current Sentence:").grid(row=0, column=2, sticky="w")
        self.lbl_current_sentence = ttk.Label(qfrm, text="-")
        self.lbl_current_sentence.grid(row=0, column=3, sticky="w", padx=(6,20))

        ttk.Label(qfrm, text="Word Queue:").grid(row=1, column=0, sticky="w")
        self.lbl_word_queue = ttk.Label(qfrm, text="0")
        self.lbl_word_queue.grid(row=1, column=1, sticky="w", padx=(6,20))
        ttk.Label(qfrm, text="Current Word:").grid(row=1, column=2, sticky="w")
        self.lbl_current_word = ttk.Label(qfrm, text="-")
        self.lbl_current_word.grid(row=1, column=3, sticky="w", padx=(6,20))

        # Add combobox for existing words and redisplay / skip / save buttons
        ttk.Label(qfrm, text="Saved Words:").grid(row=2, column=0, sticky="w", pady=(6,0))
        self.word_combo = ttk.Combobox(qfrm, values=[], width=30)
        self.word_combo.grid(row=2, column=1, columnspan=1, sticky="w", pady=(6,0))
        ttk.Button(qfrm, text="Re-display Word", command=self.on_redisplay_word).grid(row=2, column=2, sticky="w", pady=(6,0))
        ttk.Button(qfrm, text="Skip Flash", command=self.on_toggle_skip).grid(row=2, column=3, sticky="w", pady=(6,0))
        ttk.Button(qfrm, text="Save PNG", command=self.on_save_png_for_selected).grid(row=2, column=4, sticky="w", pady=(6,0))

        # Q&A
        frm_q = ttk.LabelFrame(self, text="Ask GPT-oss (general question)")
        frm_q.pack(fill=tk.BOTH, expand=False, padx=10, pady=(4, 8))
        self.txt_ask = ScrolledText(frm_q, height=3, wrap=tk.WORD)
        self.txt_ask.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        self.txt_ask.configure(font=self.small_bold)
        self.txt_ask.bind("<Control-Return>", lambda e: self.on_send_ask())

        btns2 = ttk.Frame(self)
        btns2.pack(fill=tk.X, padx=10)
        ttk.Button(btns2, text="Ask (Ctrl+Enter)", command=self.on_send_ask).pack(side=tk.LEFT, padx=4)
        ttk.Button(btns2, text="Clear Display", command=self.on_clear_display).pack(side=tk.LEFT, padx=4)

        hint = ttk.Label(self, text="Tip: Ctrl+Enter で送信。ES/FR は現在無効化されています。", foreground="#666")
        hint.pack(side=tk.BOTTOM, pady=(0,8))

        # initialize saved words list
        self.refresh_word_list()

    # --- actions ---
    def on_show_display(self): self.display.show()
    def on_show_wordflash(self): self.wordhud.show()
    def on_clear_display(self): self.display.text.delete("1.0", tk.END)

    def on_save(self):
        self.config_dict["SERVER_API_HOST"] = self.var_server.get().strip()
        self.config_dict["MODEL_NAME"] = self.var_model.get().strip()
        self.config_dict["TEMPERATURE"] = float(self.var_temp.get())
        self.config_dict["MAX_TOKENS"] = int(self.var_maxtok.get())
        try:
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(self.config_dict, f, ensure_ascii=False, indent=2)
        except Exception: pass
        messagebox.showinfo("Saved", "Settings saved.")

    def on_reconnect(self):
        self.on_save(); self._init_client()
        if self.client_ok: messagebox.showinfo("Reconnected", "LM Studio client configured.")
        else: messagebox.showwarning("Connection", "Client not available. Check lmstudio server.")

    # --- queue status updater ---
    def _update_queue_labels(self):
        try:
            self.lbl_sentence_queue.config(text=str(self.sentence_q.qsize()))
            self.lbl_word_queue.config(text=str(self.word_q.qsize()))
            with self._cur_lock:
                self.lbl_current_sentence.config(text=self._current_sentence or "-")
                self.lbl_current_word.config(text=self._current_word or "-")
            # update Display window's S/W counters next to language boxes
            try:
                self.display.update_queue_counts(self.sentence_q.qsize(), self.word_q.qsize())
            except Exception:
                pass
        except Exception:
            pass
        self.after(500, self._update_queue_labels)

    # --- sentence worker (single worker ensures sequential processing) ---
    def _sentence_worker(self):
        while True:
            try:
                task = self.sentence_q.get()
            except Exception:
                time.sleep(0.1); continue
            if not task:
                self.sentence_q.task_done(); continue
            # task: dict with keys: model_name,prompt,ts14,lang,log_path,cfg,de_text
            model_name = task.get("model_name")
            prompt = task.get("prompt")
            ts14 = task.get("ts14")
            lang = task.get("lang")
            log_path = task.get("log_path")
            cfg = task.get("cfg")
            de_text = task.get("de_text")
            # set current
            with self._cur_lock:
                self._current_sentence = f"{lang}:{de_text[:30]}"
            # start status on GUI thread
            try:
                self.after(0, lambda l=lang: self.display.start_status(l))
                # perform translation (blocking)
                model = lms.llm(model_name) if model_name else lms.llm()
                res = model.respond(prompt, config=cfg)
                raw = res if isinstance(res, str) else getattr(res, "content", str(res))
            except Exception as e:
                raw = f"(error) {e}"
            # write logs and update UI on main thread
            write_text(log_path, raw)
            final_text = extract_final_message_only(raw) or raw.strip()
            self.after(0, lambda t=final_text, ts=ts14, l=lang: self._apply_translate_result(ts, l, t))
            with self._cur_lock:
                self._current_sentence = None
            self.sentence_q.task_done()

    # --- word queue worker (single worker) ---
    def _word_worker_queue(self):
        while True:
            try:
                task = self.word_q.get()
            except Exception:
                time.sleep(0.1); continue
            if not task:
                self.word_q.task_done(); continue
            # task: dict with keys: model_name, word, cfg
            model_name = task.get("model_name")
            word = task.get("word")
            cfg = task.get("cfg")
            # set current
            with self._cur_lock:
                self._current_word = word
            try:
                # perform two sub-requests (en, ja)
                en = self._ask_word(model_name, prompt_word_en(word), cfg)
                ja = self._ask_word(model_name, prompt_word_ja(word), cfg)
            except Exception as e:
                en = f"(error) {e}"; ja = ""
            # update DB (increment count)
            db = load_word_db()
            prev = int(db.get(word, {}).get("count", 0))
            sk = int(db.get(word, {}).get("skip", 0))
            db[word] = {"en": en, "ja": ja, "count": prev + 1, "skip": sk}
            save_word_db(db)
            count_now = prev + 1

            # ensure 2s buffer since last wordflash display
            now_t = time.time()
            elapsed = now_t - self._last_wordflash_time
            if elapsed < 2.0:
                time.sleep(2.0 - elapsed)
            self._last_wordflash_time = time.time()

            # show on UI thread unless skip flag is set
            if int(db[word].get("skip",0)) == 0:
                # normal automatic display (no star)
                self.after(0, lambda c=count_now, de=word, e=en, j=ja: self.wordhud.show_word(c, de, e, j, starred=False))
                # if this is the first time (count_now == 1), save PNG automatically
                if count_now == 1:
                    try:
                        self._save_wordflash_png(de=word)
                    except Exception:
                        pass
            else:
                # skip==1: do not display automatically
                pass

            # refresh word combobox in main thread
            self.after(0, self.refresh_word_list)

            with self._cur_lock:
                self._current_word = None
            self.word_q.task_done()
            # small gap to keep loop responsive
            time.sleep(0.25)

    # --- translate sequence (enqueue only EN and JA) ---
    def on_send_translate(self):
        if not self.client_ok or lms is None:
            messagebox.showerror("Client", "lmstudio client not available.")
            return
        text = self.txt_in.get("1.0", tk.END).strip()
        if not text:
            messagebox.showwarning("Input", "Please enter German text.")
            return

        now = tz_now_jst()
        ts_hm  = now.strftime("%H:%M")
        ts12   = now.strftime("%y%m%d%H%M%S")
        ts14   = now.strftime("%Y%m%d%H%M%S")
        ymd    = now.strftime("%Y%m%d")

        day_dir = os.path.join(os.path.dirname(__file__), "log", ymd)
        ensure_dir(day_dir)
        f_input = os.path.join(day_dir, f"{ts12}_001_Input.log")
        f_ja    = os.path.join(day_dir, f"{ts12}_002_Japanese.log")
        f_en    = os.path.join(day_dir, f"{ts12}_003_English.log")
        # ES/FR logs left for compatibility but will not be used
        f_es    = os.path.join(day_dir, f"{ts12}_004_Spanish.log")
        f_fr    = os.path.join(day_dir, f"{ts12}_005_French.log")
        archive = os.path.join(os.path.dirname(__file__), "log", "archive.csv")

        p_en = prompt_en(text); p_ja = prompt_ja(text)
        write_text(f_input, f"[{ts14}] INPUT: {text}\n\n[ENGLISH]\n{p_en}\n\n[JAPANESE]\n{p_ja}\n")

        # UI display initial
        self.display.append_line("───", tag="in")
        self.display.append_line(f"{ts_hm} : {text}", tag="in")

        # pending record for archiving; keep es/fr as None but ES/FR are disabled
        self.pending[ts14] = {"de": text, "ja": None, "en": None, "es": None, "fr": None}

        cfg = {"temperature": float(self.var_temp.get())}
        mt_val = int(self.var_maxtok.get())
        if mt_val > 0: cfg["maxTokens"] = mt_val
        model_name = (self.var_model.get().strip() or self.config_dict.get("MODEL_NAME"))

        # Enqueue EN then JA to preserve sequence
        self.sentence_q.put({
            "model_name": model_name,
            "prompt": p_en,
            "ts14": ts14,
            "lang": "en",
            "log_path": f_en,
            "cfg": cfg,
            "de_text": text
        })
        self.sentence_q.put({
            "model_name": model_name,
            "prompt": p_ja,
            "ts14": ts14,
            "lang": "ja",
            "log_path": f_ja,
            "cfg": cfg,
            "de_text": text
        })

        # enqueue word translations to word_q (tokenize and queue)
        words = tokenize_german(text)
        if words:
            for w in words:
                self.word_q.put({"model_name": model_name, "word": w, "cfg": cfg})

    def _apply_translate_result(self, ts14: str, lang: str, text: str):
        # mark done status (this will be called from main thread via after)
        self.display.done_status(lang)
        tag_for_lang = {"ja":"ja","en":"en","es":"es","fr":"fr"}[lang]
        self.display.append_line(text, tag_for_lang)

        if ts14 in self.pending:
            self.pending[ts14][lang] = text
            d = self.pending[ts14]
            # since es/fr disabled, archive when en & ja present
            if all(d.get(k) is not None for k in ("ja","en")):
                row = [ts14, d["de"], d["ja"], d["en"], d.get("es") or "", d.get("fr") or ""]
                append_csv_row(os.path.join(os.path.dirname(__file__), "log", "archive.csv"),
                               row,
                               header=["Time","Germany","Japanese","English","Spanish","French"])
                self.pending.pop(ts14, None)

    # --- Word flow legacy removed; now handled via queue worker using _ask_word ---
    def _ask_word(self, model_name: str, prompt: str, cfg: dict) -> str:
        try:
            model = lms.llm(model_name) if model_name else lms.llm()
            res = model.respond(prompt, config=cfg)
            raw = res if isinstance(res, str) else getattr(res, "content", str(res))
        except Exception as e:
            raw = f"(error) {e}"
        return extract_final_message_only(raw) or ""

    # --- refresh combobox of saved words ---
    def refresh_word_list(self):
        """
        Updated: show combobox items as "word (count)".
        The selected value is preserved where possible.
        """
        db = load_word_db()
        # build list of "word (count)" strings sorted by count desc then word
        values = [f"{w} ({int(db[w].get('count',0))})" for w in sorted(db.keys(), key=lambda x: (-db[x].get("count",0), x))]
        try:
            cur = self.word_combo.get()
            self.word_combo['values'] = values
            # try to preserve selection: if cur is "word (n)" or "word", find matching item
            if cur:
                base = cur
                if " (" in cur and cur.endswith(")"):
                    base = cur.rsplit(" (", 1)[0]
                for v in values:
                    if v.startswith(base + " (") or v == base:
                        self.word_combo.set(v)
                        break
        except Exception:
            pass

    # --- re-display selected word (no count change) ---
    def on_redisplay_word(self):
        """
        Re-display selected word regardless of skip.
        Show ★ before count ONLY when the word's skip flag == 1.
        """
        sel_raw = (self.word_combo.get() or "").strip()
        if not sel_raw:
            messagebox.showinfo("Re-display", "Please choose a saved word to re-display.")
            return
        # sel_raw may be "word (count)" or just "word"
        if " (" in sel_raw and sel_raw.endswith(")"):
            word = sel_raw.rsplit(" (", 1)[0].strip()
        else:
            word = sel_raw
        if not word:
            messagebox.showinfo("Re-display", "Please choose a valid saved word.")
            return

        db = load_word_db()
        info = db.get(word)
        if not info:
            messagebox.showerror("Re-display", f"No data for '{word}'.")
            return

        en = info.get("en", "")
        ja = info.get("ja", "")
        # safety: ensure count is int
        try:
            cnt = int(info.get("count", 0))
        except Exception:
            cnt = 0

        # safety: ensure skip_flag is int 0 or 1 (fallback 0)
        try:
            skip_flag = int(info.get("skip", 0))
        except Exception:
            skip_flag = 0
        if skip_flag not in (0, 1):
            skip_flag = 0

        # ★は skip==1 のときだけ
        starred = (skip_flag == 1)

        # buffer: if last wordflash <2s ago, schedule delayed display in background
        now_t = time.time()
        elapsed = now_t - self._last_wordflash_time
        if elapsed < 2.0:
            def _delayed_show():
                time.sleep(2.0 - elapsed)
                self._last_wordflash_time = time.time()
                # starred determined by skip flag
                self.after(0, lambda c=cnt, de=word, e=en, j=ja, s=starred: self.wordhud.show_word(c, de, e, j, starred=s))
            threading.Thread(target=_delayed_show, daemon=True).start()
        else:
            self._last_wordflash_time = time.time()
            self.wordhud.show_word(cnt, word, en, ja, starred=starred)

    # --- toggle skip flag for selected word ---
    def on_toggle_skip(self):
        sel_raw = (self.word_combo.get() or "").strip()
        if not sel_raw:
            messagebox.showinfo("Skip Flash", "Please choose a saved word to toggle skip.")
            return
        if " (" in sel_raw and sel_raw.endswith(")"):
            word = sel_raw.rsplit(" (", 1)[0].strip()
        else:
            word = sel_raw
        if not word:
            messagebox.showinfo("Skip Flash", "Please choose a valid saved word.")
            return
        db = load_word_db()
        info = db.get(word)
        if not info:
            messagebox.showerror("Skip Flash", f"No data for '{word}'.")
            return
        cur = int(info.get("skip",0))
        new = 0 if cur==1 else 1
        info["skip"] = new
        db[word] = info
        save_word_db(db)
        self.refresh_word_list()
        messagebox.showinfo("Skip Flash", f"'{word}' skip set to {new}.")

    # --- save PNG for selected word (manual) ---
    def on_save_png_for_selected(self):
        sel_raw = (self.word_combo.get() or "").strip()
        if not sel_raw:
            messagebox.showinfo("Save PNG", "Please choose a saved word to save PNG.")
            return
        if " (" in sel_raw and sel_raw.endswith(")"):
            word = sel_raw.rsplit(" (", 1)[0].strip()
        else:
            word = sel_raw
        if not word:
            messagebox.showinfo("Save PNG", "Please choose a valid saved word.")
            return
        try:
            self._save_wordflash_png(de=word)
            messagebox.showinfo("Save PNG", f"Saved WordFlash PNG for '{word}'.")
        except Exception as e:
            messagebox.showerror("Save PNG", f"Failed to save PNG: {e}")

    def _save_wordflash_png(self, de: str):
        """
        Capture WordFlash window and save PNG to log/FlashPNG/yyyyMMdd/(sanitized_word).png
        """
        now = tz_now_jst()
        ymd = now.strftime("%Y%m%d")
        out_dir = os.path.join(os.path.dirname(__file__), "log", "FlashPNG", ymd)
        ensure_dir(out_dir)
        filename = sanitize_filename(de)[:120] + ".png"
        path = os.path.join(out_dir, filename)
        # attempt to capture via PIL.ImageGrab using window geometry
        try:
            if ImageGrab is None:
                raise RuntimeError("Pillow ImageGrab not available")
            # ensure window is visible and updated
            self.wordhud.update_idletasks()
            x = self.wordhud.winfo_rootx()
            y = self.wordhud.winfo_rooty()
            w = self.wordhud.winfo_width()
            h = self.wordhud.winfo_height()
            if w <=0 or h <=0:
                raise RuntimeError("WordFlash window size invalid")
            bbox = (x, y, x + w, y + h)
            img = ImageGrab.grab(bbox=bbox)
            img.save(path)
            return path
        except Exception as e:
            # fallback: try to capture the two canvases as postscript and convert if PIL available
            try:
                # create an image by rendering the canvases to postscript and converting
                ps_path = os.path.join(out_dir, sanitize_filename(de)[:120] + ".ps")
                # attempt to grab en_canvas postscript and combine... (best-effort)
                self.wordhud.update_idletasks()
                en_ps = self.wordhud.en_canvas.postscript(colormode='color')
                # try to convert en_ps to image if PIL present
                if Image is None:
                    raise RuntimeError("PIL not available to convert postscript")
                from io import BytesIO
                # NOTE: PIL cannot directly open a Postscript string without external tools;
                # this fallback is best-effort and may fail on many systems.
                img_en = Image.open(BytesIO(en_ps.encode('utf-8')))
                img_en.save(path)
                return path
            except Exception as e2:
                raise RuntimeError(f"PNG save failed: {e}; fallback failed: {e2}")

    # --- Q&A ---
    def on_send_ask(self):
        if not self.client_ok or lms is None:
            messagebox.showerror("Client", "lmstudio client not available."); return
        q = self.txt_ask.get("1.0", tk.END).strip()
        if not q:
            messagebox.showwarning("Input", "Please enter a question."); return

        now = tz_now_jst()
        ts_hm  = now.strftime("%H:%M")
        ts12   = now.strftime("%y%m%d%H%M%S")
        ts14   = now.strftime("%Y%m%d%H%M%S")
        ymd    = now.strftime("%Y%m%d")

        self.display.append_line(f"{ts_hm} : {q}", tag="qa_in")

        day_dir = os.path.join(os.path.dirname(__file__), "log", ymd)
        ensure_dir(day_dir)
        f_q = os.path.join(day_dir, f"{ts12}_006_Question.log")
        f_a = os.path.join(day_dir, f"{ts12}_007_Answer.log")
        write_text(f_q, q)

        self.display.start_status("ja")

        cfg = {"temperature": float(self.var_temp.get())}
        mt_val = int(self.var_maxtok.get())
        if mt_val > 0: cfg["maxTokens"] = mt_val
        model_name = (self.var_model.get().strip() or self.config_dict.get("MODEL_NAME"))
        threading.Thread(target=self._ask_worker, args=(model_name, q, cfg, f_a), daemon=True).start()

    def _ask_worker(self, model_name: str, prompt: str, cfg: dict, log_path: str):
        try:
            model = lms.llm(model_name) if model_name else lms.llm()
            res = model.respond(prompt, config=cfg)
            raw = res if isinstance(res, str) else getattr(res, "content", str(res))
        except Exception as e:
            raw = f"(error) {e}"
        write_text(log_path, raw)
        final_text = extract_final_message_only(raw) or raw.strip()
        self.after(0, lambda: self._apply_ask_result(final_text))

    def _apply_ask_result(self, text: str):
        self.display.done_status("ja")
        self.display.append_line(text, tag="qa_out")

# ---------- entry ----------
if __name__ == "__main__":
    app = App()
    app.mainloop()
