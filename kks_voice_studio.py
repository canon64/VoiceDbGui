"""
KKS Voice Studio
-----------------
Tab 1: 抽出   - UnityPy で KKS の AssetBundle から WAV を抽出
Tab 2: DB構築 - 抽出済み WAV から SQLite DB を構築
Tab 3: ブラウズ - DB を閲覧・絞り込み・エクスポート
"""

import csv
import datetime as dt
import json
import os
import queue
import re
import shutil
import sqlite3
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

try:
    import UnityPy
    UNITYPY_OK = True
except ImportError:
    UNITYPY_OK = False

# ── Constants ─────────────────────────────────────────────────────────────────

APP_STATE_PATH = Path(__file__).resolve().with_name("kks_voice_studio_state.json")
HISTORY_MAX    = 200
INVALID_FS_CHARS = '<>:"/\\|?*'

ALL_CHARS = [f"c{i:02d}" for i in range(44)] + ["c-13", "c-100"]

# h_{type}_{char}_{level}_{seq}.wav
FILENAME_RE = re.compile(
    r"^h_([a-z0-9]+)_(-?\d+)_(\d{2})_(\d+)\.wav$", re.IGNORECASE)

TYPE_INFO = {
    "ai":   ("喘ぎ",    "aibu"),
    "fe":   ("前戯",    "fe"),
    "hh":   ("奉仕",    "houshi"),
    "hh3p": ("奉仕3P",  "houshi_3p"),
    "ka":   ("愛撫",    "aibu"),
    "ka3p": ("愛撫3P",  "aibu_3p"),
    "ko":   ("行為中",  "ko"),
    "on":   ("オナニー","masturbation"),
    "so":   ("挿入",    "sonyu"),
    "so3p": ("挿入3P",  "sonyu_3p"),
}

LEVEL_NAME = {"00": "控えめ", "01": "通常", "02": "興奮", "03": "絶頂"}

VISIBLE_COLS = {
    "voices":      ["id","chara","mode_name","voice_id","level_name","filename",
                    "file_type","insert_type","houshi_type","aibu_type",
                    "situation_type","wav_path","serif"],
    "breaths":     ["id","chara","mode_name","voice_id","level_name","group_id",
                    "filename","breath_type","wav_path","serif"],
    "shortbreaths":["id","chara","voice_id","level_name","filename",
                    "face","not_overwrite","wav_path","serif"],
}

COMBO_FILTERS  = ["chara","mode_name","level_name","file_type",
                  "insert_type","houshi_type","aibu_type","situation_type","breath_type"]
LIKE_FILTERS   = ["filename","serif","wav_path"]

# ── Helpers ───────────────────────────────────────────────────────────────────

def sanitize(value: str, max_len: int = 120) -> str:
    if not value:
        return "_"
    for c in INVALID_FS_CHARS:
        value = value.replace(c, "_")
    return value.strip()[:max_len] or "_"

def parse_voice_filename(filename: str):
    m = FILENAME_RE.match(filename)
    if not m:
        return None
    type_code, char_num, level_code, seq = m.groups()
    info = TYPE_INFO.get(type_code.lower())
    if not info:
        return None
    cn = int(char_num)
    chara = f"c{cn:02d}" if cn >= 0 else f"c{cn}"
    return {
        "chara":      chara,
        "mode_name":  info[1],
        "voice_id":   int(seq),
        "level":      int(level_code),
        "level_name": LEVEL_NAME.get(level_code, f"level_{level_code}"),
        "filename":   filename,
        "file_type":  type_code.lower(),
    }

# ── Extract Tab ───────────────────────────────────────────────────────────────

class ExtractTab(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self._log_queue  = queue.Queue()
        self._running    = False
        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 6, "pady": 3}

        # KKS root
        fr = tk.Frame(self)
        fr.pack(fill="x", **pad)
        tk.Label(fr, text="KKSフォルダ:", width=14, anchor="w").pack(side="left")
        self._kks_var = tk.StringVar()
        tk.Entry(fr, textvariable=self._kks_var).pack(side="left", fill="x", expand=True)
        tk.Button(fr, text="参照", command=self._browse_kks).pack(side="left", padx=2)

        # Output dir
        fr2 = tk.Frame(self)
        fr2.pack(fill="x", **pad)
        tk.Label(fr2, text="WAV出力先:", width=14, anchor="w").pack(side="left")
        self._out_var = tk.StringVar(value=str(Path.home() / "kks_wav"))
        tk.Entry(fr2, textvariable=self._out_var).pack(side="left", fill="x", expand=True)
        tk.Button(fr2, text="参照", command=self._browse_out).pack(side="left", padx=2)

        # Char select
        lf = tk.LabelFrame(self, text="キャラクター選択")
        lf.pack(fill="x", padx=6, pady=3)

        btn_fr = tk.Frame(lf)
        btn_fr.pack(fill="x")
        tk.Button(btn_fr, text="全選択",  command=self._select_all).pack(side="left", padx=2)
        tk.Button(btn_fr, text="全解除", command=self._deselect_all).pack(side="left", padx=2)

        cb_fr = tk.Frame(lf)
        cb_fr.pack(fill="x")
        self._char_vars = {}
        for i, ch in enumerate(ALL_CHARS):
            var = tk.BooleanVar(value=False)
            self._char_vars[ch] = var
            tk.Checkbutton(cb_fr, text=ch, variable=var, width=7).grid(
                row=i // 12, column=i % 12, sticky="w")

        # Start / Stop
        ctrl = tk.Frame(self)
        ctrl.pack(fill="x", **pad)
        self._start_btn = tk.Button(ctrl, text="▶ 抽出開始", command=self._start,
                                    bg="#4CAF50", fg="white", width=16)
        self._start_btn.pack(side="left", padx=2)
        self._stop_btn  = tk.Button(ctrl, text="■ 停止", command=self._stop,
                                    state=tk.DISABLED, width=10)
        self._stop_btn.pack(side="left", padx=2)
        self._status_var = tk.StringVar(value="待機中")
        tk.Label(ctrl, textvariable=self._status_var).pack(side="left", padx=8)

        # Log
        self._log = tk.Text(self, height=18, state=tk.DISABLED,
                            font=("Consolas", 9), wrap="word")
        sb = tk.Scrollbar(self, command=self._log.yview)
        self._log.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self._log.pack(fill="both", expand=True, padx=6, pady=3)

    def _browse_kks(self):
        d = filedialog.askdirectory(title="KKSインストールフォルダを選択")
        if d:
            self._kks_var.set(d)

    def _browse_out(self):
        d = filedialog.askdirectory(title="WAV出力先を選択")
        if d:
            self._out_var.set(d)

    def _select_all(self):
        for v in self._char_vars.values():
            v.set(True)

    def _deselect_all(self):
        for v in self._char_vars.values():
            v.set(False)

    def get_settings(self):
        return {
            "kks_dir": self._kks_var.get(),
            "out_dir": self._out_var.get(),
            "chars": {k: v.get() for k, v in self._char_vars.items()},
        }

    def apply_settings(self, d):
        if not d:
            return
        if d.get("kks_dir"):
            self._kks_var.set(d["kks_dir"])
        if d.get("out_dir"):
            self._out_var.set(d["out_dir"])
        for k, v in d.get("chars", {}).items():
            if k in self._char_vars:
                self._char_vars[k].set(bool(v))

    def _append_log(self, text: str):
        self._log.config(state=tk.NORMAL)
        self._log.insert("end", text)
        self._log.see("end")
        self._log.config(state=tk.DISABLED)

    def _drain(self):
        try:
            while True:
                item = self._log_queue.get_nowait()
                if item == "__done__":
                    self._running = False
                    self._start_btn.config(state=tk.NORMAL)
                    self._stop_btn.config(state=tk.DISABLED)
                    self._status_var.set("完了")
                    return
                self._append_log(item)
        except queue.Empty:
            pass
        if self._running:
            self.after(100, self._drain)

    def _start(self):
        if not UNITYPY_OK:
            messagebox.showerror("エラー", "UnityPy が見つかりません。\npip install UnityPy")
            return
        kks = self._kks_var.get().strip()
        out = self._out_var.get().strip()
        chars = [c for c, v in self._char_vars.items() if v.get()]
        if not kks:
            messagebox.showerror("エラー", "KKSフォルダを指定してください。")
            return
        if not out:
            messagebox.showerror("エラー", "WAV出力先を指定してください。")
            return
        if not chars:
            messagebox.showerror("エラー", "キャラクターを1つ以上選択してください。")
            return
        self._running = True
        self._start_btn.config(state=tk.DISABLED)
        self._stop_btn.config(state=tk.NORMAL)
        self._status_var.set("抽出中...")
        threading.Thread(target=self._worker, args=(kks, out, chars),
                         daemon=True).start()
        self.after(100, self._drain)

    def _stop(self):
        self._running = False
        self._log_queue.put("[停止要求]\n")

    def _worker(self, kks_root: str, out_dir: str, chars: list):
        total = 0
        for char in chars:
            if not self._running:
                self._log_queue.put("[停止しました]\n")
                break
            bundle_dir = Path(kks_root) / "abdata" / "sound" / "data" / "pcm" / char / "h"
            if not bundle_dir.exists():
                self._log_queue.put(f"[skip] {char}: フォルダなし\n")
                continue
            char_out = Path(out_dir) / char
            char_out.mkdir(parents=True, exist_ok=True)
            count = 0
            for bp in sorted(bundle_dir.glob("*.unity3d")):
                if not self._running:
                    break
                self._log_queue.put(f"  [{char}] {bp.name}\n")
                try:
                    env = UnityPy.load(str(bp))
                    for obj in env.objects:
                        if obj.type.name != "AudioClip":
                            continue
                        clip = obj.read()
                        wav_name = clip.m_Name + ".wav"
                        out_path = char_out / wav_name
                        if out_path.exists():
                            continue
                        for audio_data in clip.samples.values():
                            out_path.write_bytes(audio_data)
                            count += 1
                            break
                except Exception as e:
                    self._log_queue.put(f"  [error] {bp.name}: {e}\n")
            self._log_queue.put(f"[完了] {char}: {count} ファイル\n")
            total += count
        self._log_queue.put(f"\n── 合計 {total} ファイル抽出 ──\n")
        self._log_queue.put("__done__")


# ── Build DB Tab ──────────────────────────────────────────────────────────────

DB_DDL = """
CREATE TABLE IF NOT EXISTS voices (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    chara TEXT, mode_name TEXT, voice_id INTEGER,
    level INTEGER, level_name TEXT,
    filename TEXT, file_type TEXT,
    insert_type TEXT DEFAULT '', houshi_type TEXT DEFAULT '',
    aibu_type TEXT DEFAULT '', situation_type TEXT DEFAULT '',
    wav_path TEXT, serif TEXT DEFAULT ''
);
CREATE TABLE IF NOT EXISTS breaths (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    chara TEXT, mode_name TEXT, voice_id INTEGER,
    level INTEGER, level_name TEXT,
    group_id TEXT DEFAULT '', filename TEXT,
    breath_type TEXT DEFAULT '', wav_path TEXT, serif TEXT DEFAULT ''
);
CREATE TABLE IF NOT EXISTS shortbreaths (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    chara TEXT, voice_id INTEGER,
    level INTEGER, level_name TEXT,
    filename TEXT, face INTEGER DEFAULT -1,
    not_overwrite INTEGER DEFAULT 0,
    wav_path TEXT, serif TEXT DEFAULT ''
);
CREATE INDEX IF NOT EXISTS idx_voices_chara     ON voices(chara);
CREATE INDEX IF NOT EXISTS idx_voices_mode      ON voices(mode_name);
CREATE INDEX IF NOT EXISTS idx_voices_level     ON voices(level);
CREATE INDEX IF NOT EXISTS idx_voices_file_type ON voices(file_type);
"""

class BuildDbTab(tk.Frame):
    def __init__(self, parent, on_build_done=None):
        super().__init__(parent)
        self._log_queue   = queue.Queue()
        self._running     = False
        self._on_done     = on_build_done  # callback(db_path: str)
        self._last_db     = None
        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 6, "pady": 3}

        # WAV source dir
        fr = tk.Frame(self)
        fr.pack(fill="x", **pad)
        tk.Label(fr, text="WAVフォルダ:", width=16, anchor="w").pack(side="left")
        self._wav_var = tk.StringVar(value=str(Path.home() / "kks_wav"))
        tk.Entry(fr, textvariable=self._wav_var).pack(side="left", fill="x", expand=True)
        tk.Button(fr, text="参照", command=self._browse_wav).pack(side="left", padx=2)

        # DB output
        fr2 = tk.Frame(self)
        fr2.pack(fill="x", **pad)
        tk.Label(fr2, text="DB出力先:", width=16, anchor="w").pack(side="left")
        self._db_var = tk.StringVar(
            value=str(Path.home() / "kks_voices.db"))
        tk.Entry(fr2, textvariable=self._db_var).pack(side="left", fill="x", expand=True)
        tk.Button(fr2, text="参照", command=self._browse_db).pack(side="left", padx=2)

        # Button
        ctrl = tk.Frame(self)
        ctrl.pack(fill="x", **pad)
        self._build_btn = tk.Button(ctrl, text="▶ DB構築", command=self._start,
                                    bg="#2196F3", fg="white", width=16)
        self._build_btn.pack(side="left", padx=2)
        self._status_var = tk.StringVar(value="待機中")
        tk.Label(ctrl, textvariable=self._status_var).pack(side="left", padx=8)

        # Log
        self._log = tk.Text(self, height=22, state=tk.DISABLED,
                            font=("Consolas", 9), wrap="word")
        sb = tk.Scrollbar(self, command=self._log.yview)
        self._log.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self._log.pack(fill="both", expand=True, padx=6, pady=3)

    def _browse_wav(self):
        d = filedialog.askdirectory(title="WAVフォルダを選択")
        if d:
            self._wav_var.set(d)

    def _browse_db(self):
        p = filedialog.asksaveasfilename(
            title="DB出力先を選択",
            defaultextension=".db",
            filetypes=[("SQLite DB", "*.db"), ("All", "*.*")])
        if p:
            self._db_var.set(p)

    def get_settings(self):
        return {
            "wav_dir": self._wav_var.get(),
            "db_path": self._db_var.get(),
        }

    def apply_settings(self, d):
        if not d:
            return
        if d.get("wav_dir"):
            self._wav_var.set(d["wav_dir"])
        if d.get("db_path"):
            self._db_var.set(d["db_path"])

    def _append_log(self, text: str):
        self._log.config(state=tk.NORMAL)
        self._log.insert("end", text)
        self._log.see("end")
        self._log.config(state=tk.DISABLED)

    def _drain(self):
        try:
            while True:
                item = self._log_queue.get_nowait()
                if item == "__done__":
                    self._running = False
                    self._build_btn.config(state=tk.NORMAL)
                    self._status_var.set("完了")
                    if self._on_done and self._last_db:
                        self._on_done(self._last_db)
                    return
                self._append_log(item)
        except queue.Empty:
            pass
        if self._running:
            self.after(100, self._drain)

    def _start(self):
        wav = self._wav_var.get().strip()
        db  = self._db_var.get().strip()
        if not wav:
            messagebox.showerror("エラー", "WAVフォルダを指定してください。")
            return
        if not db:
            messagebox.showerror("エラー", "DB出力先を指定してください。")
            return
        self._running = True
        self._build_btn.config(state=tk.DISABLED)
        self._status_var.set("構築中...")
        threading.Thread(target=self._worker,
                         args=(wav, db),
                         daemon=True).start()
        self.after(100, self._drain)

    def _worker(self, wav_dir: str, db_path: str):
        try:
            # DB出力先がディレクトリならファイル名を補完
            p = Path(db_path)
            if p.is_dir() or not p.suffix:
                p = p / "kks_voices.db"
                db_path = str(p)
            self._last_db = db_path
            self._log_queue.put(f"[DB] 出力先: {db_path}\n")
            p.parent.mkdir(parents=True, exist_ok=True)
            conn = sqlite3.connect(db_path)
            conn.executescript(DB_DDL)
            conn.execute("DELETE FROM voices")
            conn.execute("DELETE FROM breaths")
            conn.execute("DELETE FROM shortbreaths")
            conn.commit()

            voices_rows = []
            total_skip  = 0

            wav_root = Path(wav_dir)
            char_dirs = sorted(wav_root.glob("c*"))
            for char_dir in char_dirs:
                if not char_dir.is_dir():
                    continue
                char = char_dir.name
                wavs = sorted(char_dir.rglob("*.wav"))
                self._log_queue.put(f"[{char}] {len(wavs)} ファイル処理中...\n")
                for wav_path in wavs:
                    fn = wav_path.name
                    parsed = parse_voice_filename(fn)
                    if parsed is None:
                        total_skip += 1
                        continue
                    voices_rows.append((
                        parsed["chara"],
                        parsed["mode_name"],
                        parsed["voice_id"],
                        parsed["level"],
                        parsed["level_name"],
                        fn,
                        parsed["file_type"],
                        "", "", "", "",     # insert_type, houshi_type, aibu_type, situation_type
                        str(wav_path),
                        "",  # serif は後からブラウズタブのエクスポートCSVで設定
                    ))

            self._log_queue.put(f"[DB] {len(voices_rows)} 件 INSERT 中...\n")
            conn.executemany("""
                INSERT INTO voices
                    (chara, mode_name, voice_id, level, level_name, filename,
                     file_type, insert_type, houshi_type, aibu_type, situation_type,
                     wav_path, serif)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, voices_rows)
            conn.commit()
            conn.close()

            self._log_queue.put(
                f"\n── 完了 ──\n"
                f"  voices : {len(voices_rows)} 件\n"
                f"  スキップ: {total_skip} 件（名前が不一致）\n"
                f"  DB出力 : {db_path}\n"
            )
        except Exception as e:
            self._log_queue.put(f"[ERROR] {e}\n")
        finally:
            self._log_queue.put("__done__")


# ── Browse Tab ────────────────────────────────────────────────────────────────

class BrowseTab(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.conn             = None
        self.table_columns    = {}
        self.current_rows     = []
        self.current_visible  = []
        self.current_where    = ""
        self.current_params   = []
        self.app_state        = {"last": None, "history": []}
        self.history_win      = None
        self.history_list     = None
        self._load_state()
        self._build_ui()
        self._apply_last()

    # ── UI ──
    def _build_ui(self):
        top = tk.Frame(self)
        top.pack(fill="x", padx=6, pady=3)

        # DB path
        tk.Label(top, text="DB:", width=4, anchor="w").pack(side="left")
        self._db_var = tk.StringVar()
        tk.Entry(top, textvariable=self._db_var, width=50).pack(side="left")
        tk.Button(top, text="参照", command=self._choose_db).pack(side="left", padx=2)
        tk.Button(top, text="接続", command=self._connect).pack(side="left", padx=2)

        # Export dir
        tk.Label(top, text="  保存先:", anchor="w").pack(side="left")
        self._exp_var = tk.StringVar(value=str(Path.home() / "kks_voice_export"))
        tk.Entry(top, textvariable=self._exp_var, width=30).pack(side="left")
        tk.Button(top, text="参照", command=self._choose_exp).pack(side="left", padx=2)

        # Table selector
        mid = tk.Frame(self)
        mid.pack(fill="x", padx=6, pady=2)
        tk.Label(mid, text="テーブル:").pack(side="left")
        self._tbl_var = tk.StringVar(value="voices")
        self._tbl_combo = ttk.Combobox(mid, textvariable=self._tbl_var,
                                       state="readonly", width=14)
        self._tbl_combo.pack(side="left", padx=4)
        self._tbl_combo.bind("<<ComboboxSelected>>", lambda e: self._on_table_changed())
        tk.Label(mid, text="1ページ:").pack(side="left")
        self._psize_var = tk.IntVar(value=500)
        ttk.Spinbox(mid, textvariable=self._psize_var,
                    from_=10, to=5000, increment=100, width=6).pack(side="left")
        self._page_var = tk.IntVar(value=1)
        tk.Button(mid, text="◀", command=self._prev_page).pack(side="left", padx=2)
        tk.Label(mid, text="ページ:").pack(side="left")
        tk.Label(mid, textvariable=self._page_var).pack(side="left")
        tk.Button(mid, text="▶", command=self._next_page).pack(side="left", padx=2)
        self._total_var = tk.StringVar(value="0件")
        tk.Label(mid, textvariable=self._total_var).pack(side="left", padx=8)

        # Filters
        filt_lf = tk.LabelFrame(self, text="フィルタ")
        filt_lf.pack(fill="x", padx=6, pady=2)
        self._combo_vars = {k: tk.StringVar() for k in COMBO_FILTERS}
        self._like_vars  = {k: tk.StringVar() for k in LIKE_FILTERS}
        self._combo_widgets = {}
        self._like_widgets  = {}

        row1 = tk.Frame(filt_lf)
        row1.pack(fill="x")
        for k in COMBO_FILTERS:
            fr = tk.Frame(row1)
            fr.pack(side="left", padx=3)
            tk.Label(fr, text=k, font=("", 8)).pack()
            cb = ttk.Combobox(fr, textvariable=self._combo_vars[k],
                              state="readonly", width=14)
            cb.pack()
            self._combo_widgets[k] = cb

        row2 = tk.Frame(filt_lf)
        row2.pack(fill="x", pady=2)
        for k in LIKE_FILTERS:
            fr = tk.Frame(row2)
            fr.pack(side="left", padx=3)
            tk.Label(fr, text=f"{k}含む", font=("", 8)).pack()
            e = tk.Entry(fr, textvariable=self._like_vars[k], width=20)
            e.pack()
            self._like_widgets[k] = e

        btns = tk.Frame(filt_lf)
        btns.pack(fill="x", pady=2)
        tk.Button(btns, text="検索", command=self._search,
                  bg="#4CAF50", fg="white", width=10).pack(side="left", padx=4)
        tk.Button(btns, text="クリア", command=self._clear_filters,
                  width=8).pack(side="left", padx=2)
        tk.Button(btns, text="履歴", command=self._open_history,
                  width=8).pack(side="left", padx=2)

        # Tree + Detail
        pane = tk.PanedWindow(self, orient="vertical", sashwidth=6)
        pane.pack(fill="both", expand=True, padx=6, pady=3)

        tree_fr = tk.Frame(pane)
        pane.add(tree_fr, height=320)
        self._tree = ttk.Treeview(tree_fr, selectmode="extended")
        xsb = ttk.Scrollbar(tree_fr, orient="horizontal",
                             command=self._tree.xview)
        ysb = ttk.Scrollbar(tree_fr, orient="vertical",
                             command=self._tree.yview)
        self._tree.configure(xscrollcommand=xsb.set, yscrollcommand=ysb.set)
        xsb.pack(side="bottom", fill="x")
        ysb.pack(side="right",  fill="y")
        self._tree.pack(fill="both", expand=True)
        self._tree.bind("<<TreeviewSelect>>", self._on_select)

        det_fr = tk.Frame(pane)
        pane.add(det_fr, height=120)
        self._detail = tk.Text(det_fr, height=6, state=tk.DISABLED,
                               font=("Consolas", 9), wrap="word")
        det_sb = tk.Scrollbar(det_fr, command=self._detail.yview)
        self._detail.configure(yscrollcommand=det_sb.set)
        det_sb.pack(side="right", fill="y")
        self._detail.pack(fill="both", expand=True)

        # Export buttons
        exp_fr = tk.Frame(self)
        exp_fr.pack(fill="x", padx=6, pady=3)
        tk.Button(exp_fr, text="表示中を保存",
                  command=lambda: self._export(all_displayed=True),
                  width=16).pack(side="left", padx=2)
        tk.Button(exp_fr, text="選択行を保存",
                  command=lambda: self._export(all_displayed=False),
                  width=16).pack(side="left", padx=2)
        self._status_var = tk.StringVar(value="Ready")
        tk.Label(exp_fr, textvariable=self._status_var).pack(side="left", padx=8)

    # ── DB ──
    def _choose_db(self):
        p = filedialog.askopenfilename(
            title="DB を選択",
            filetypes=[("SQLite", "*.db"), ("All", "*.*")])
        if p:
            self._db_var.set(p)

    def _choose_exp(self):
        d = filedialog.askdirectory(title="保存先を選択")
        if d:
            self._exp_var.set(d)

    def _connect(self):
        db_path = self._db_var.get().strip()
        if not db_path or not Path(db_path).is_file():
            messagebox.showerror("Error", f"DBが見つかりません:\n{db_path}")
            return
        try:
            if self.conn:
                self.conn.close()
            self.conn = sqlite3.connect(db_path)
            self.conn.row_factory = sqlite3.Row
            self.table_columns = {}
            cur = self.conn.execute(
                "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
            tables = [r[0] for r in cur]
            for t in tables:
                cols = [r[1] for r in
                        self.conn.execute(f"PRAGMA table_info({t})")]
                self.table_columns[t] = cols
            self._tbl_combo["values"] = tables
            if "voices" in tables:
                self._tbl_var.set("voices")
            elif tables:
                self._tbl_var.set(tables[0])
            self._on_table_changed()
            self._status_var.set(f"接続: {Path(db_path).name}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _on_table_changed(self):
        self._page_var.set(1)
        self._refresh_filter_state()
        self._load_distinct_values()

    def _refresh_filter_state(self):
        tbl = self._tbl_var.get()
        cols = self.table_columns.get(tbl, [])
        for k, w in self._combo_widgets.items():
            w.config(state="readonly" if k in cols else tk.DISABLED)
            if k not in cols:
                self._combo_vars[k].set("")
        for k, w in self._like_widgets.items():
            w.config(state=tk.NORMAL if k in cols else tk.DISABLED)
            if k not in cols:
                self._like_vars[k].set("")

    def _load_distinct_values(self):
        if not self.conn:
            return
        tbl  = self._tbl_var.get()
        cols = self.table_columns.get(tbl, [])
        for k in COMBO_FILTERS:
            if k not in cols:
                continue
            cur = self.conn.execute(
                f"SELECT DISTINCT TRIM({k}) FROM {tbl} WHERE {k} IS NOT NULL "
                f"ORDER BY TRIM({k})")
            vals = [""] + list({r[0] for r in cur if r[0]})
            vals.sort()
            self._combo_widgets[k]["values"] = vals

    def _build_where(self):
        tbl  = self._tbl_var.get()
        cols = self.table_columns.get(tbl, [])
        clauses, params = [], []
        for k in COMBO_FILTERS:
            v = self._combo_vars[k].get().strip()
            if v and k in cols:
                clauses.append(f"TRIM({k}) = ?")
                params.append(v)
        for k in LIKE_FILTERS:
            v = self._like_vars[k].get().strip()
            if v and k in cols:
                clauses.append(f"{k} LIKE ?")
                params.append(f"%{v}%")
        where = ("WHERE " + " AND ".join(clauses)) if clauses else ""
        return where, params

    def _search(self):
        self._page_var.set(1)
        self._run_query()
        self._push_history()
        self._save_last()

    def _run_query(self):
        if not self.conn:
            return
        tbl   = self._tbl_var.get()
        cols  = self.table_columns.get(tbl, [])
        where, params = self._build_where()
        self.current_where  = where
        self.current_params = params

        # Count
        cnt = self.conn.execute(
            f"SELECT COUNT(*) FROM {tbl} {where}", params).fetchone()[0]
        self._total_var.set(f"{cnt:,}件")

        # Order column
        order = next((c for c in ["id","idx","voice_id","filename","rowid"]
                      if c in cols), "rowid")

        # Paging
        size   = max(1, self._psize_var.get())
        page   = max(1, self._page_var.get())
        offset = (page - 1) * size

        cur = self.conn.execute(
            f"SELECT * FROM {tbl} {where} ORDER BY {order} "
            f"LIMIT {size} OFFSET {offset}", params)
        self.current_rows    = [dict(r) for r in cur]
        self.current_visible = [c for c in VISIBLE_COLS.get(tbl, []) if c in cols]

        self._populate_tree()

    def _populate_tree(self):
        self._tree.delete(*self._tree.get_children())
        self._tree["columns"] = self.current_visible
        self._tree["show"]    = "headings"
        widths = {"id":50,"chara":60,"mode_name":90,"voice_id":70,
                  "level_name":70,"filename":200,"file_type":70,
                  "wav_path":300,"serif":300}
        for c in self.current_visible:
            w = widths.get(c, 100)
            self._tree.heading(c, text=c)
            self._tree.column(c, width=w, minwidth=40, stretch=False)
        for row in self.current_rows:
            vals = [str(row.get(c, "")) for c in self.current_visible]
            self._tree.insert("", "end", values=vals)

    def _on_select(self, _event=None):
        sel = self._tree.selection()
        if not sel:
            return
        idx = self._tree.index(sel[0])
        if idx >= len(self.current_rows):
            return
        row = self.current_rows[idx]
        text = "\n".join(f"{k}: {v}" for k, v in row.items())
        self._detail.config(state=tk.NORMAL)
        self._detail.delete("1.0", "end")
        self._detail.insert("end", text)
        self._detail.config(state=tk.DISABLED)

    def _prev_page(self):
        p = self._page_var.get()
        if p > 1:
            self._page_var.set(p - 1)
            self._run_query()

    def _next_page(self):
        self._page_var.set(self._page_var.get() + 1)
        self._run_query()

    def _clear_filters(self):
        for v in self._combo_vars.values():
            v.set("")
        for v in self._like_vars.values():
            v.set("")

    # ── Export ──
    def _get_rows_for_export(self, all_displayed: bool):
        if all_displayed:
            return self.current_rows
        sel = self._tree.selection()
        idxs = [self._tree.index(s) for s in sel]
        return [self.current_rows[i] for i in idxs
                if i < len(self.current_rows)]

    def _export_path(self, row: dict, tbl: str) -> str:
        chara    = sanitize(str(row.get("chara",    "")))
        mode     = sanitize(str(row.get("mode_name","") or
                                f"mode_{row.get('mode','')}"))
        level    = sanitize(str(row.get("level_name","") or
                                f"level_{row.get('level','')}"))
        category = sanitize(str(
            row.get("file_type","") or
            row.get("breath_type","") or ""))
        filename = row.get("filename","") or "unknown.wav"
        ext      = Path(filename).suffix or ".wav"
        stem     = sanitize(Path(filename).stem)
        return str(Path(tbl) / chara / mode / level / category /
                   (stem + ext))

    def _export(self, all_displayed: bool):
        rows    = self._get_rows_for_export(all_displayed)
        exp_dir = self._exp_var.get().strip()
        tbl     = self._tbl_var.get()
        if not rows:
            messagebox.showinfo("Info", "エクスポート対象がありません。")
            return
        if not exp_dir:
            messagebox.showerror("Error", "保存先を指定してください。")
            return

        ts      = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        manifest_path = Path(exp_dir) / f"export_manifest_{tbl}_{ts}.csv"
        vtext_path    = Path(exp_dir) / f"export_voice_text_{tbl}_{ts}.csv"
        seen    = set()
        copied  = skipped = missing = failed = 0

        with open(manifest_path, "w", newline="", encoding="utf-8") as mf, \
             open(vtext_path,    "w", newline="", encoding="utf-8") as vf:
            mw = csv.writer(mf)
            vw = csv.writer(vf, delimiter="|")
            mw.writerow(["table","id","chara","mode_name","level_name",
                          "filename","source_wav_path","exported_path"])
            for row in rows:
                src = row.get("wav_path","")
                if not src or not Path(src).is_file():
                    missing += 1
                    continue
                if src in seen:
                    skipped += 1
                    continue
                seen.add(src)
                rel  = self._export_path(row, tbl)
                dest = Path(exp_dir) / rel
                dest.parent.mkdir(parents=True, exist_ok=True)
                try:
                    shutil.copy2(src, dest)
                    copied += 1
                    mw.writerow([tbl, row.get("id",""),
                                  row.get("chara",""), row.get("mode_name",""),
                                  row.get("level_name",""), row.get("filename",""),
                                  src, str(dest)])
                    serif = row.get("serif","") or ""
                    vw.writerow([row.get("filename",""),
                                  row.get("chara",""), "JP", serif])
                except Exception:
                    failed += 1

        msg = (f"完了\nコピー: {copied}\n"
               f"スキップ（重複）: {skipped}\n"
               f"ファイルなし: {missing}\n"
               f"失敗: {failed}")
        self._status_var.set(
            f"コピー:{copied} スキップ:{skipped} なし:{missing}")
        messagebox.showinfo("エクスポート完了", msg)

    # ── History ──
    def _snapshot(self):
        return {
            "saved_at": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "db_path":  self._db_var.get(),
            "export_dir": self._exp_var.get(),
            "table":    self._tbl_var.get(),
            "page_size": self._psize_var.get(),
            "combo_filters": {k: v.get() for k, v in self._combo_vars.items()},
            "like_filters":  {k: v.get() for k, v in self._like_vars.items()},
        }

    def _apply_snapshot(self, snap):
        if not snap:
            return
        if snap.get("db_path"):
            self._db_var.set(snap["db_path"])
        if snap.get("export_dir"):
            self._exp_var.set(snap["export_dir"])
        if snap.get("table"):
            self._tbl_var.set(snap["table"])
        if snap.get("page_size"):
            self._psize_var.set(snap["page_size"])
        for k, v in snap.get("combo_filters", {}).items():
            if k in self._combo_vars:
                self._combo_vars[k].set(v)
        for k, v in snap.get("like_filters", {}).items():
            if k in self._like_vars:
                self._like_vars[k].set(v)

    def _save_last(self):
        self.app_state["last"] = self._snapshot()
        self._write_state()

    def _push_history(self):
        snap = self._snapshot()
        hist = self.app_state.setdefault("history", [])
        hist.insert(0, snap)
        self.app_state["history"] = hist[:HISTORY_MAX]
        self._write_state()

    def _load_state(self):
        if APP_STATE_PATH.exists():
            try:
                self.app_state = json.loads(APP_STATE_PATH.read_text("utf-8"))
            except Exception:
                pass

    def _write_state(self):
        try:
            # 既存ファイルのキー（extract/build など）を保持して上書き
            existing = {}
            if APP_STATE_PATH.exists():
                try:
                    existing = json.loads(APP_STATE_PATH.read_text("utf-8"))
                except Exception:
                    pass
            existing["last"]    = self.app_state.get("last")
            existing["history"] = self.app_state.get("history", [])
            tmp = APP_STATE_PATH.with_suffix(".tmp")
            tmp.write_text(json.dumps(existing, ensure_ascii=False, indent=2),
                           encoding="utf-8")
            tmp.replace(APP_STATE_PATH)
        except Exception:
            pass

    def _apply_last(self):
        self._apply_snapshot(self.app_state.get("last"))

    def _open_history(self):
        if self.history_win and self.history_win.winfo_exists():
            self.history_win.lift()
            return
        self.history_win = tk.Toplevel(self)
        self.history_win.title("検索履歴")
        self.history_win.geometry("520x400")
        lb_fr = tk.Frame(self.history_win)
        lb_fr.pack(fill="both", expand=True, padx=6, pady=6)
        self.history_list = tk.Listbox(lb_fr, width=70)
        sb = tk.Scrollbar(lb_fr, command=self.history_list.yview)
        self.history_list.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.history_list.pack(fill="both", expand=True)
        self._refresh_history()
        btns = tk.Frame(self.history_win)
        btns.pack(fill="x", padx=6, pady=4)
        tk.Button(btns, text="適用", command=self._apply_history,
                  width=10).pack(side="left", padx=2)
        tk.Button(btns, text="削除", command=self._delete_history,
                  width=10).pack(side="left", padx=2)
        tk.Button(btns, text="全削除",
                  command=self._clear_history, width=10).pack(side="left", padx=2)

    def _refresh_history(self):
        if not self.history_list:
            return
        self.history_list.delete(0, "end")
        for h in self.app_state.get("history", []):
            ts  = h.get("saved_at","")
            tbl = h.get("table","")
            combo = {k: v for k, v in
                     h.get("combo_filters",{}).items() if v}
            like  = {k: v for k, v in
                     h.get("like_filters", {}).items() if v}
            label = f"{ts}  [{tbl}]"
            if combo:
                label += "  " + " ".join(f"{k}={v}" for k,v in combo.items())
            if like:
                label += "  " + " ".join(f"{k}~{v}" for k,v in like.items())
            self.history_list.insert("end", label)

    def _apply_history(self):
        sel = self.history_list.curselection()
        if not sel:
            return
        snap = self.app_state["history"][sel[0]]
        self._apply_snapshot(snap)
        if self.conn:
            self._on_table_changed()
            self._run_query()

    def _delete_history(self):
        sel = self.history_list.curselection()
        if not sel:
            return
        del self.app_state["history"][sel[0]]
        self._write_state()
        self._refresh_history()

    def _clear_history(self):
        if not messagebox.askyesno("確認", "履歴を全削除しますか？"):
            return
        self.app_state["history"] = []
        self._write_state()
        self._refresh_history()


# ── Main App ──────────────────────────────────────────────────────────────────

class KksVoiceStudio(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("KKS Voice Studio")
        self.geometry("1300x900")

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True)

        self._tab_extract = ExtractTab(nb)
        self._tab_browse  = BrowseTab(nb)
        self._tab_build   = BuildDbTab(nb, on_build_done=self._on_build_done)

        nb.add(self._tab_extract, text="  抽出  ")
        nb.add(self._tab_build,   text="  DB構築  ")
        nb.add(self._tab_browse,  text="  ブラウズ  ")

        self._load_settings()

    def _on_build_done(self, db_path: str):
        self._tab_browse._db_var.set(db_path)
        self._tab_browse._connect()

    def _load_settings(self):
        if not APP_STATE_PATH.exists():
            return
        try:
            state = json.loads(APP_STATE_PATH.read_text("utf-8"))
            self._tab_extract.apply_settings(state.get("extract"))
            self._tab_build.apply_settings(state.get("build"))
        except Exception:
            pass

    def _save_settings(self):
        try:
            existing = {}
            if APP_STATE_PATH.exists():
                try:
                    existing = json.loads(APP_STATE_PATH.read_text("utf-8"))
                except Exception:
                    pass
            existing["extract"] = self._tab_extract.get_settings()
            existing["build"]   = self._tab_build.get_settings()
            tmp = APP_STATE_PATH.with_suffix(".tmp")
            tmp.write_text(json.dumps(existing, ensure_ascii=False, indent=2),
                           encoding="utf-8")
            tmp.replace(APP_STATE_PATH)
        except Exception:
            pass

    def destroy(self):
        self._save_settings()
        try:
            if self._tab_browse.conn:
                self._tab_browse.conn.close()
        except Exception:
            pass
        super().destroy()


def main():
    app = KksVoiceStudio()
    app.mainloop()

if __name__ == "__main__":
    main()
