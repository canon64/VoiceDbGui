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
from collections import defaultdict
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
    "fe":   ("前戯",    "foreplay"),
    "hh":   ("奉仕",    "houshi"),
    "hh3p": ("奉仕3P",  "houshi_3p"),
    "ka":   ("愛撫",    "aibu_touch"),
    "ka3p": ("愛撫3P",  "aibu_3p"),
    "ko":   ("行為中",  "act_common"),
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

def _load_char_display_map(kks_dir: str) -> dict:
    """ハードコードされたcharacter_mapを返す。"""
    return {
        "c00": "c00 セクシー系お姉さま",
        "c01": "c01 お嬢様",
        "c02": "c02 タカビー",
        "c03": "c03 小悪魔",
        "c04": "c04 ミステリアス",
        "c05": "c05 電波",
        "c06": "c06 大和撫子",
        "c07": "c07 ボーイッシュ",
        "c08": "c08 純真無垢",
        "c09": "c09 アホの子",
        "c10": "c10 邪気眼",
        "c11": "c11 母性的",
        "c12": "c12 姉御肌",
        "c13": "c13 ギャル",
        "c14": "c14 不良少女",
        "c15": "c15 野性的",
        "c16": "c16 意識高クール",
        "c17": "c17 ひねくれ",
        "c18": "c18 不幸少女",
        "c19": "c19 文学少女",
        "c20": "c20 モジモジ",
        "c21": "c21 正統派ヒロイン",
        "c22": "c22 ミーハー",
        "c23": "c23 オタク女子",
        "c24": "c24 ヤンデレ",
        "c25": "c25 ダル",
        "c26": "c26 無口",
        "c27": "c27 意地っ張り",
        "c28": "c28 ロリばばぁ",
        "c29": "c29 素直クール",
        "c30": "c30 気さく",
        "c31": "c31 勝ち気",
        "c32": "c32 誠実",
        "c33": "c33 艶やか",
        "c34": "c34 帰国子女",
        "c35": "c35 方言娘",
        "c36": "c36 Ｓッ気",
        "c37": "c37 無感情",
        "c38": "c38 几帳面",
        "c39": "c39 島っ娘",
        "c40": "c40 高潔",
        "c41": "c41 ボクっ娘",
        "c42": "c42 天真爛漫",
        "c43": "c43 ノリノリ",
    }

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
    tc = type_code.lower()
    info = TYPE_INFO.get(tc)
    if not info:
        return None
    cn = int(char_num)
    chara = f"c{cn:02d}" if cn >= 0 else f"c{cn}"
    ft = info[1]  # 元DBのfile_type値 ("sonyu", "act_common", "aibu_touch" 等)
    return {
        "chara":          chara,
        "mode_name":      ft,
        "voice_id":       int(seq),
        "level":          int(level_code),
        "level_name":     LEVEL_NAME.get(level_code, f"level_{level_code}"),
        "filename":       Path(filename).stem,  # 拡張子なし（元DBに合わせる）
        "file_type":      ft,
        "type_code":      tc,   # 生の型コード (so/hh/ka/ai/on/ko/so3p等)
        "insert_type":    "",
        "houshi_type":    "",
        "aibu_type":      "",
        "situation_type": "",
    }


# ── VoicePatternData 型マップ構築 ─────────────────────────────────────────────

_TEKOKI_CONDS  = {1, 2}
_PAIZURI_CONDS = {9, 10, 11, 12, 13, 14, 15, 16, 17}
_WP_IS_CONDS   = {1, 3, 5, 7, 9, 11}
_WP_NOT_CONDS  = {2, 4, 6, 8, 10}

_START_TAGS = {
    0: "not_insert_ok",   1: "dangerous_day",    2: "safe_day",
    3: "girlfriend_virgin", 4: "aibu_pos0",       5: "aibu_pos1",
    6: "houshi_fera",     7: "houshi_tekoki",     8: "houshi_paizuri",
    9: "sonyu_female",   10: "sonyu_pos0",       11: "sonyu_pos1",
    12: "sonyu_pos2",    13: "insert_ok",        14: "aibu_pos0",
    15: "aibu_pos1",     16: "sonyu_female",     17: "role_main",
    18: "role_sub",      19: "sonyu_normal",
}
_MAST_TAGS = {0: "found"}
_LES_TAGS  = {0: "role_main", 1: "role_sub"}
_TAG_ORDER = [
    "not_insert_ok", "insert_ok", "dangerous_day", "safe_day", "girlfriend_virgin",
    "aibu_pos0", "aibu_pos1", "houshi_fera", "houshi_tekoki", "houshi_paizuri",
    "sonyu_female", "sonyu_normal", "sonyu_pos0", "sonyu_pos1", "sonyu_pos2",
    "role_main", "role_sub", "found",
]


def _load_pattern_trees(kks_dir: str, log_fn) -> dict:
    """h/list/*.unity3d から全 VoicePatternData を読み込む。mode → [tree] 辞書を返す。"""
    h_list = Path(kks_dir) / "abdata" / "h" / "list"
    if not h_list.is_dir():
        log_fn(f"[WARN] {h_list} が見つかりません\n")
        return {}
    suffixes = {
        0: "voice_00_00", 1: "voice_01_00", 2: "voice_02_00",
        3: "voice_03_00", 4: "voice_04_00", 6: "voice_06_00",
    }
    result = {m: [] for m in suffixes}
    for bp in sorted(h_list.glob("*.unity3d")):
        try:
            env = UnityPy.load(str(bp))
            for key in env.container:
                for mode, suf in suffixes.items():
                    if suf in key:
                        try:
                            result[mode].append(env.container[key].read_typetree())
                        except Exception:
                            pass
        except Exception as e:
            log_fn(f"[WARN] {bp.name}: {e}\n")
    return result


def _iter_lst_info(trees):
    for tree in trees:
        for param in tree.get("param", []):
            for li in param.get("lstInfo", []):
                yield param, li


def _build_insert_map(trees: list) -> dict:
    """voice_03_00 → voice_id → insert_type"""
    tags = defaultdict(set)
    for _, li in _iter_lst_info(trees):
        conds  = set(li.get("lstConditions", []))
        voices = li.get("lstVoice", []) + li.get("lstSecondVoice", [])
        for vid in voices:
            if 0 in conds: tags[vid].add("kokan")
            if 1 in conds: tags[vid].add("anal")
    result = {}
    for vid, t in tags.items():
        if "kokan" in t and "anal" in t: result[vid] = "both"
        elif "kokan" in t: result[vid] = "kokan"
        elif "anal"  in t: result[vid] = "anal"
    return result


def _build_houshi_map(trees: list) -> dict:
    """voice_02_00 → voice_id → houshi_type"""
    tags = defaultdict(set)
    for _, li in _iter_lst_info(trees):
        conds  = set(li.get("lstConditions", []))
        voices = li.get("lstVoice", []) + li.get("lstSecondVoice", [])
        for vid in voices:
            if 0 in conds:                  tags[vid].add("fera")
            if conds & _TEKOKI_CONDS:       tags[vid].add("tekoki")
            if conds & _PAIZURI_CONDS:      tags[vid].add("paizuri")
    return {vid: "/".join(sorted(t)) for vid, t in tags.items() if t}


def _build_aibu_map(trees: list) -> dict:
    """voice_01_00 → voice_id → aibu_type"""
    vid_tags  = defaultdict(set)
    kiss_vids = set()
    for param, li in _iter_lst_info(trees):
        pid    = param.get("id", -1)
        conds  = set(li.get("lstConditions", []))
        voices = li.get("lstVoice", []) + li.get("lstSecondVoice", [])
        for vid in voices:
            if pid == 41:                       kiss_vids.add(vid)
            if conds & _WP_IS_CONDS:            vid_tags[vid].add("weakpoint")
            if 1 in conds:                      vid_tags[vid].add("no_weakpoint")
            if 0 in conds:                      vid_tags[vid].add("has_weakpoint")
            if conds & _WP_NOT_CONDS and not (conds & _WP_IS_CONDS):
                                                vid_tags[vid].add("nonspot")
    order = ["kiss", "weakpoint", "has_weakpoint", "no_weakpoint", "nonspot"]
    result = {}
    for vid in set(list(vid_tags.keys()) + list(kiss_vids)):
        t = {"kiss"} if vid in kiss_vids else vid_tags.get(vid, set())
        if t:
            result[vid] = "/".join(sorted(t, key=lambda x: order.index(x) if x in order else 99))
    return result


def _build_situation_map(trees: list, tag_dict: dict) -> dict:
    """voice_00_00 / 04 / 06 → voice_id → situation_type"""
    vid_tags = defaultdict(set)
    for _, li in _iter_lst_info(trees):
        conds  = set(li.get("lstConditions", []))
        voices = li.get("lstVoice", []) + li.get("lstSecondVoice", [])
        for c, tag in tag_dict.items():
            if c in conds:
                for vid in voices:
                    vid_tags[vid].add(tag)
    result = {}
    for vid, tags in vid_tags.items():
        result[vid] = "/".join(sorted(tags, key=lambda t: _TAG_ORDER.index(t) if t in _TAG_ORDER else 99))
    return result

# ── Extract Tab ───────────────────────────────────────────────────────────────

class ExtractTab(tk.Frame):
    def __init__(self, parent, on_kks_change=None):
        super().__init__(parent)
        self._log_queue    = queue.Queue()
        self._running      = False
        self._on_kks_change = on_kks_change
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
        self._out_var = tk.StringVar(value="")
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
        self._char_cbs  = {}
        for i, ch in enumerate(ALL_CHARS):
            var = tk.BooleanVar(value=False)
            self._char_vars[ch] = var
            w = tk.Checkbutton(cb_fr, text=ch, variable=var, width=14, anchor="w")
            w.grid(row=i // 8, column=i % 8, sticky="w")
            self._char_cbs[ch] = w

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

    def update_char_labels(self, char_map: dict):
        for ch, widget in self._char_cbs.items():
            widget.config(text=char_map.get(ch, ch))

    def _browse_kks(self):
        d = filedialog.askdirectory(title="KKSインストールフォルダを選択")
        if d:
            self._kks_var.set(d)
            self._out_var.set(str(Path(d) / "wave"))
            self.update_char_labels(_load_char_display_map(d))
            if self._on_kks_change:
                self._on_kks_change(d)

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
            self.update_char_labels(_load_char_display_map(d["kks_dir"]))
        if d.get("out_dir"):
            self._out_var.set(d["out_dir"])
        elif d.get("kks_dir"):
            self._out_var.set(str(Path(d["kks_dir"]) / "wave"))
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
    insert_type TEXT, houshi_type TEXT,
    aibu_type TEXT, situation_type TEXT,
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
    def __init__(self, parent, on_build_done=None, get_kks_dir=None):
        super().__init__(parent)
        self._log_queue   = queue.Queue()
        self._running     = False
        self._on_done     = on_build_done  # callback(db_path: str)
        self._get_kks_dir = get_kks_dir or (lambda: "")
        self._last_db     = None
        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 6, "pady": 3}

        # WAV source dir
        fr = tk.Frame(self)
        fr.pack(fill="x", **pad)
        tk.Label(fr, text="WAVフォルダ:", width=18, anchor="w").pack(side="left")
        self._wav_var = tk.StringVar(value="")
        tk.Entry(fr, textvariable=self._wav_var).pack(side="left", fill="x", expand=True)
        tk.Button(fr, text="参照", command=self._browse_wav).pack(side="left", padx=2)

        # DB output
        fr2 = tk.Frame(self)
        fr2.pack(fill="x", **pad)
        tk.Label(fr2, text="DB出力先:", width=18, anchor="w").pack(side="left")
        self._db_var = tk.StringVar(value="")
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
        kks = self._get_kks_dir().strip()
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
                         args=(wav, db, kks),
                         daemon=True).start()
        self.after(100, self._drain)

    def _worker(self, wav_dir: str, db_path: str, kks_dir: str):
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

            # ── VoicePatternData から型マップを構築 ──
            type_maps = {}
            if kks_dir and UNITYPY_OK and (Path(kks_dir) / "abdata" / "h" / "list").is_dir():
                self._log_queue.put("[DB] AssetBundle から型データ読み込み中...\n")
                try:
                    ptrees = _load_pattern_trees(kks_dir, self._log_queue.put)
                    type_maps = {
                        "insert":   _build_insert_map(ptrees.get(3, [])),
                        "houshi":   _build_houshi_map(ptrees.get(2, [])),
                        "aibu":     _build_aibu_map(ptrees.get(1, [])),
                        "sit_0":    _build_situation_map(ptrees.get(0, []), _START_TAGS),
                        "sit_4":    _build_situation_map(ptrees.get(4, []), _MAST_TAGS),
                        "sit_6":    _build_situation_map(ptrees.get(6, []), _LES_TAGS),
                    }
                    self._log_queue.put(
                        f"[DB] 型マップ: insert={len(type_maps['insert'])}, "
                        f"houshi={len(type_maps['houshi'])}, "
                        f"aibu={len(type_maps['aibu'])}, "
                        f"situation={len(type_maps['sit_0'])}\n"
                    )
                except Exception as e:
                    self._log_queue.put(f"[WARN] 型データ読み込みエラー: {e}\n")
            elif kks_dir:
                self._log_queue.put("[WARN] KKSフォルダ内に abdata/h/list が見つかりません\n")
            else:
                self._log_queue.put("[DB] KKSフォルダ未指定 → 型列は空のまま構築\n")

            # ── CSVからセリフ辞書を構築 (kks_dir/voice_extract/voice_csv/ を自動検索) ──
            # キー: WAVファイル名(拡張子あり), 値: セリフ文字列
            serif_map = {}
            csv_dir = Path(kks_dir) / "voice_extract" / "voice_csv" if kks_dir else None
            if csv_dir and csv_dir.is_dir():
                csv_files = sorted(csv_dir.glob("c*.csv"))
                for cp in csv_files:
                    try:
                        with cp.open(encoding="utf-8-sig") as f:
                            for line in f:
                                parts = line.rstrip("\n").split("|")
                                if len(parts) >= 4:
                                    serif_map[parts[0]] = parts[3]
                    except Exception as e:
                        self._log_queue.put(f"[WARN] CSV読み込みエラー {cp.name}: {e}\n")
                self._log_queue.put(f"[DB] セリフ辞書: {len(serif_map)}件 ({len(csv_files)}ファイル)\n")
            else:
                self._log_queue.put("[DB] voice_extract/voice_csv が見つからないため serif は空\n")

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

                    vid = parsed["voice_id"]
                    tc  = parsed["type_code"]

                    # 型マップから各列を解決 (未定義は None → DB NULL)
                    if type_maps:
                        if tc in ("so", "so3p"):
                            insert_type    = type_maps["insert"].get(vid)
                            houshi_type    = None
                            aibu_type      = None
                            situation_type = None
                        elif tc in ("hh", "hh3p"):
                            insert_type    = None
                            houshi_type    = type_maps["houshi"].get(vid)
                            aibu_type      = None
                            situation_type = None
                        elif tc == "ai":
                            insert_type    = None
                            houshi_type    = None
                            aibu_type      = type_maps["aibu"].get(vid)
                            situation_type = None
                        elif tc in ("ka", "ka3p"):
                            insert_type    = None
                            houshi_type    = None
                            aibu_type      = None
                            situation_type = type_maps["sit_0"].get(vid)
                        elif tc == "on":
                            insert_type    = None
                            houshi_type    = None
                            aibu_type      = None
                            situation_type = type_maps["sit_4"].get(vid)
                        elif tc == "ko":
                            insert_type    = None
                            houshi_type    = None
                            aibu_type      = None
                            situation_type = type_maps["sit_6"].get(vid)
                        else:
                            insert_type = houshi_type = aibu_type = situation_type = None
                    else:
                        insert_type = houshi_type = aibu_type = situation_type = None

                    voices_rows.append((
                        parsed["chara"],
                        parsed["mode_name"],
                        parsed["voice_id"],
                        parsed["level"],
                        parsed["level_name"],
                        parsed["filename"],
                        parsed["file_type"],
                        insert_type, houshi_type,
                        aibu_type,   situation_type,
                        str(wav_path),
                        serif_map.get(fn, ""),
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
        self._char_display_map = {}   # {code: "c13 ギャル"}
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
        self._exp_var = tk.StringVar(value="")
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
        tk.Button(exp_fr, text="全選択",
                  command=self._select_all_rows,
                  width=8).pack(side="left", padx=2)
        tk.Button(exp_fr, text="表示中を保存",
                  command=lambda: self._export(all_displayed=True),
                  width=16).pack(side="left", padx=2)
        tk.Button(exp_fr, text="選択行を保存",
                  command=lambda: self._export(all_displayed=False),
                  width=16).pack(side="left", padx=2)
        self._save_csv_var = tk.BooleanVar(value=True)
        tk.Checkbutton(exp_fr, text="CSVも保存",
                       variable=self._save_csv_var).pack(side="left", padx=6)
        self._flat_var = tk.BooleanVar(value=False)
        tk.Checkbutton(exp_fr, text="フラット(1フォルダ)",
                       variable=self._flat_var).pack(side="left", padx=6)
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
            if k == "chara":
                vals = [""] + sorted({
                    self._char_display_map.get(r[0], r[0])
                    for r in cur if r[0]
                })
            else:
                vals = [""] + sorted({r[0] for r in cur if r[0]})
            self._combo_widgets[k]["values"] = vals

    def _build_where(self):
        tbl  = self._tbl_var.get()
        cols = self.table_columns.get(tbl, [])
        clauses, params = [], []
        for k in COMBO_FILTERS:
            v = self._combo_vars[k].get().strip()
            if v and k in cols:
                if k == "chara":
                    v = v.split()[0]  # "c13 ギャル" → "c13"
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

        # すべての結果を取得（ページング無し）
        cur = self.conn.execute(
            f"SELECT * FROM {tbl} {where} ORDER BY {order}", params)
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

    def _select_all_rows(self):
        self._tree.selection_set(self._tree.get_children())

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

    def _build_relative_export_path(self, row: dict) -> Path:
        tbl   = self._tbl_var.get()
        chara = sanitize(str(row.get("chara") or ""))
        mode_name = row.get("mode_name")
        mode_seg  = sanitize(str(mode_name)) if mode_name \
                    else f"mode_{row.get('mode','unknown')}"
        level_name = row.get("level_name")
        level_seg  = sanitize(str(level_name)) if level_name \
                     else f"level_{row.get('level','unknown')}"
        category = (row.get("file_type") or row.get("breath_type") or
                    row.get("houshi_type") or row.get("aibu_type") or
                    row.get("situation_type") or "voice")
        cat_seg  = sanitize(str(category))
        src = str(row.get("wav_path") or "")
        ext = Path(src).suffix if Path(src).suffix else ".wav"
        fn  = sanitize(str(row.get("filename") or f"id_{row.get('id','unknown')}"))
        return Path(tbl) / chara / mode_seg / level_seg / cat_seg / f"{fn}{ext}"


    def _voice_text_row(self, row: dict) -> list:
        fn    = str(row.get("filename") or "").strip()
        if not fn:
            src = str(row.get("wav_path") or "")
            fn  = Path(src).name if src else f"id_{row.get('id','unknown')}.wav"
        if not Path(fn).suffix:
            fn += ".wav"
        chara = str(row.get("chara") or "").strip() or "unknown"
        serif = str(row.get("serif") or "")
        serif = serif.replace("\r\n", "\n").replace("\r", "\n").replace("\n", " ")
        return [fn, chara, "JP", serif]

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

        filter_parts = [
            sanitize(self._combo_vars[k].get())
            for k in COMBO_FILTERS
            if self._combo_vars[k].get().strip()
        ]
        filter_tag = "_".join(filter_parts) if filter_parts else tbl

        dest_root = Path(exp_dir)
        if self._flat_var.get():
            dest_root = dest_root / filter_tag
        dest_root.mkdir(parents=True, exist_ok=True)

        copied = missing = failed = duplicate_skipped = 0
        voice_text_rows = []
        seen_sources    = set()
        seen_dest_paths = set()

        for row in rows:
            src = row.get("wav_path")
            if not src:
                missing += 1
                continue
            src_norm = os.path.normcase(os.path.normpath(str(src)))
            if src_norm in seen_sources:
                duplicate_skipped += 1
                continue
            rel_norm = os.path.normcase(str(self._build_relative_export_path(row)))
            if rel_norm in seen_dest_paths:
                duplicate_skipped += 1
                continue
            if not os.path.isfile(src):
                missing += 1
                continue
            if self._flat_var.get():
                src_path = str(row.get("wav_path") or "")
                ext = Path(src_path).suffix if Path(src_path).suffix else ".wav"
                fn  = sanitize(str(row.get("filename") or f"id_{row.get('id','unknown')}"))
                dst = dest_root / f"{fn}{ext}"
            else:
                dst = dest_root / self._build_relative_export_path(row)
            dst.parent.mkdir(parents=True, exist_ok=True)
            try:
                shutil.copy2(src, dst)
                copied += 1
                seen_sources.add(src_norm)
                seen_dest_paths.add(rel_norm)
                voice_text_rows.append(self._voice_text_row(row))
            except Exception:
                failed += 1

        if self._save_csv_var.get() and voice_text_rows:
            stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
            vtext_path = dest_root / f"voice_text_{filter_tag}_{stamp}.csv"
            with vtext_path.open("w", newline="", encoding="utf-8-sig") as f:
                csv.writer(f, delimiter="|", lineterminator="\n").writerows(voice_text_rows)

        msg = (f"保存完了\n対象行: {len(rows)}\n保存成功: {copied}\n"
               f"重複スキップ: {duplicate_skipped}\nファイルなし: {missing}\n失敗: {failed}")
        self._status_var.set(msg.replace("\n", " | "))
        messagebox.showinfo("エクスポート完了", msg)
        os.startfile(dest_root)

    # ── History ──
    def _snapshot(self):
        return {
            "saved_at": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "db_path":  self._db_var.get(),
            "export_dir": self._exp_var.get(),
            "table":    self._tbl_var.get(),
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
        for k, v in snap.get("combo_filters", {}).items():
            if k in self._combo_vars:
                if k == "chara" and v:
                    v = self._char_display_map.get(v, v)
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
        kks_dir = self.app_state.get("extract", {}).get("kks_dir", "")
        self._char_display_map = _load_char_display_map(kks_dir)

    def set_char_map(self, char_map: dict):
        self._char_display_map = char_map
        self._load_distinct_values()

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
        if not self._exp_var.get():
            kks_dir = self.app_state.get("extract", {}).get("kks_dir", "")
            if kks_dir:
                self._exp_var.set(str(Path(kks_dir) / "extract_wave"))
        db_path = self._db_var.get().strip()
        if db_path and Path(db_path).is_file():
            self._connect()

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

        self._tab_extract = ExtractTab(nb, on_kks_change=self._on_kks_change)
        self._tab_browse  = BrowseTab(nb)
        self._tab_build   = BuildDbTab(nb, on_build_done=self._on_build_done,
                                       get_kks_dir=lambda: self._tab_extract._kks_var.get())

        nb.add(self._tab_extract, text="  抽出  ")
        nb.add(self._tab_build,   text="  DB構築  ")
        nb.add(self._tab_browse,  text="  ブラウズ  ")

        self._load_settings()

    def _on_build_done(self, db_path: str):
        self._tab_browse._db_var.set(db_path)
        self._tab_browse._connect()

    def _on_kks_change(self, kks_dir: str):
        """ExtractTab で KKS フォルダが変更されたとき BuildDbTab/BrowseTab を更新する。"""
        wave_dir = str(Path(kks_dir) / "wave")
        if not self._tab_build._wav_var.get():
            self._tab_build._wav_var.set(wave_dir)
        if not self._tab_build._db_var.get():
            self._tab_build._db_var.set(str(Path(kks_dir) / "wave" / "kks_voices.db"))
        self._tab_browse.set_char_map(_load_char_display_map(kks_dir))
        if not self._tab_browse._exp_var.get():
            self._tab_browse._exp_var.set(str(Path(kks_dir) / "extract_wave"))

    def _load_settings(self):
        if not APP_STATE_PATH.exists():
            return
        try:
            state = json.loads(APP_STATE_PATH.read_text("utf-8"))
            self._tab_extract.apply_settings(state.get("extract"))
            build_s = state.get("build", {})
            kks_dir = state.get("extract", {}).get("kks_dir", "")
            # kks_dir からデフォルトを導出（保存値が無い場合のみ）
            if kks_dir and not build_s.get("wav_dir"):
                build_s = dict(build_s, wav_dir=str(Path(kks_dir) / "wave"))
            if kks_dir and not build_s.get("db_path"):
                build_s = dict(build_s, db_path=str(Path(kks_dir) / "wave" / "kks_voices.db"))
            self._tab_build.apply_settings(build_s)
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
            self._tab_browse._save_last()
        except Exception:
            pass
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
