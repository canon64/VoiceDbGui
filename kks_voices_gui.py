import csv
import datetime as dt
import json
import os
import shutil
import sqlite3
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk


DEFAULT_DB_PATH = ""
DEFAULT_EXPORT_DIR = str(Path.home() / "kks_voice_export")
APP_STATE_PATH = Path(__file__).resolve().with_name("kks_voices_gui_state.json")
HISTORY_MAX = 200

INVALID_FS_CHARS = '<>:"/\\|?*'

VISIBLE_COLUMN_CANDIDATES = {
    "voices": [
        "id",
        "chara",
        "mode_name",
        "voice_id",
        "level_name",
        "filename",
        "file_type",
        "insert_type",
        "houshi_type",
        "aibu_type",
        "situation_type",
        "wav_path",
        "serif",
    ],
    "breaths": [
        "id",
        "chara",
        "mode_name",
        "voice_id",
        "level_name",
        "group_id",
        "filename",
        "breath_type",
        "wav_path",
        "serif",
    ],
    "shortbreaths": [
        "id",
        "chara",
        "voice_id",
        "level_name",
        "filename",
        "face",
        "not_overwrite",
        "wav_path",
        "serif",
    ],
}

FILTER_COMBO_COLUMNS = [
    "chara",
    "mode_name",
    "level_name",
    "file_type",
    "insert_type",
    "houshi_type",
    "aibu_type",
    "situation_type",
    "breath_type",
]
FILTER_LIKE_COLUMNS = ["filename", "serif", "wav_path"]


def sanitize_segment(value):
    text = "" if value is None else str(value)
    text = text.strip()
    if not text:
        return "unknown"
    text = text.replace("\r", " ").replace("\n", " ")
    for ch in INVALID_FS_CHARS:
        text = text.replace(ch, "_")
    text = "".join(ch for ch in text if ord(ch) >= 32)
    text = text.strip(" .")
    if not text:
        text = "unknown"
    return text[:120]


class KksVoiceDbGui(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("KKS Voices DB Browser")
        self.geometry("1640x980")

        self.db_path_var = tk.StringVar(value=DEFAULT_DB_PATH)
        self.export_dir_var = tk.StringVar(value=DEFAULT_EXPORT_DIR)
        self.table_var = tk.StringVar(value="voices")
        self.page_size_var = tk.IntVar(value=500)
        self.page_var = tk.IntVar(value=1)
        self.total_rows_var = tk.IntVar(value=0)
        self.status_var = tk.StringVar(value="Ready")

        self.combo_filter_vars = {col: tk.StringVar(value="") for col in FILTER_COMBO_COLUMNS}
        self.like_filter_vars = {col: tk.StringVar(value="") for col in FILTER_LIKE_COLUMNS}

        self.conn = None
        self.table_columns = {}
        self.current_rows = []
        self.current_visible_columns = []
        self.current_where_sql = ""
        self.current_where_params = []
        self.history_window = None
        self.history_listbox = None
        self.app_state = {"last": None, "history": []}

        self._load_app_state()
        self._apply_last_state_to_vars()

        self._build_ui()
        self._connect_and_load()

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(3, weight=1)

        frame_db = ttk.Frame(self, padding=8)
        frame_db.grid(row=0, column=0, sticky="ew")
        frame_db.columnconfigure(1, weight=1)
        ttk.Label(frame_db, text="DB").grid(row=0, column=0, sticky="w")
        ttk.Entry(frame_db, textvariable=self.db_path_var).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(frame_db, text="DB選択", command=self._choose_db).grid(row=0, column=2, padx=2)
        ttk.Button(frame_db, text="再接続", command=self._connect_and_load).grid(row=0, column=3, padx=2)

        frame_ctrl = ttk.Frame(self, padding=(8, 0, 8, 6))
        frame_ctrl.grid(row=1, column=0, sticky="ew")
        frame_ctrl.columnconfigure(9, weight=1)

        ttk.Label(frame_ctrl, text="テーブル").grid(row=0, column=0, sticky="w")
        self.table_combo = ttk.Combobox(frame_ctrl, textvariable=self.table_var, state="readonly", width=20)
        self.table_combo.grid(row=0, column=1, sticky="w", padx=(6, 12))
        self.table_combo.bind("<<ComboboxSelected>>", lambda _e: self._on_table_changed())

        ttk.Label(frame_ctrl, text="1ページ件数").grid(row=0, column=2, sticky="w")
        ttk.Spinbox(frame_ctrl, from_=50, to=5000, increment=50, textvariable=self.page_size_var, width=8).grid(
            row=0, column=3, sticky="w", padx=(6, 12)
        )

        ttk.Button(frame_ctrl, text="検索", command=self._on_search_clicked).grid(row=0, column=4, padx=2)
        ttk.Button(frame_ctrl, text="履歴", command=self._open_history_window).grid(row=0, column=5, padx=2)
        ttk.Button(frame_ctrl, text="前へ", command=self._prev_page).grid(row=0, column=6, padx=2)
        ttk.Button(frame_ctrl, text="次へ", command=self._next_page).grid(row=0, column=7, padx=2)
        self.page_label = ttk.Label(frame_ctrl, text="Page 1 / 1")
        self.page_label.grid(row=0, column=8, sticky="w", padx=(8, 0))

        self.total_label = ttk.Label(frame_ctrl, text="Total: 0")
        self.total_label.grid(row=0, column=9, sticky="e")

        frame_filter = ttk.LabelFrame(self, text="絞り込み", padding=8)
        frame_filter.grid(row=2, column=0, sticky="ew", padx=8, pady=(0, 8))
        for i in range(8):
            frame_filter.columnconfigure(i, weight=1 if i % 2 == 1 else 0)

        self.filter_widgets = {}
        row = 0
        col = 0
        for label, key in [
            ("chara", "chara"),
            ("mode_name", "mode_name"),
            ("level_name", "level_name"),
            ("file_type", "file_type"),
            ("insert_type", "insert_type"),
            ("houshi_type", "houshi_type"),
            ("aibu_type", "aibu_type"),
            ("situation_type", "situation_type"),
            ("breath_type", "breath_type"),
        ]:
            ttk.Label(frame_filter, text=label).grid(row=row, column=col, sticky="w")
            combo = ttk.Combobox(frame_filter, textvariable=self.combo_filter_vars[key], width=20, state="readonly")
            combo.grid(row=row, column=col + 1, sticky="ew", padx=(6, 12), pady=2)
            combo.bind("<Return>", lambda _e: self._on_search_clicked())
            self.filter_widgets[key] = combo
            col += 2
            if col >= 8:
                col = 0
                row += 1

        for label, key in [("filename含む", "filename"), ("serif含む", "serif"), ("wav_path含む", "wav_path")]:
            ttk.Label(frame_filter, text=label).grid(row=row, column=col, sticky="w")
            entry = ttk.Entry(frame_filter, textvariable=self.like_filter_vars[key], width=30)
            entry.grid(row=row, column=col + 1, sticky="ew", padx=(6, 12), pady=2)
            entry.bind("<Return>", lambda _e: self._on_search_clicked())
            self.filter_widgets[key] = entry
            col += 2
            if col >= 8:
                col = 0
                row += 1

        button_row = row + 1
        ttk.Button(frame_filter, text="検索", command=self._on_search_clicked).grid(
            row=button_row, column=6, sticky="e", padx=2, pady=(4, 0)
        )
        ttk.Button(frame_filter, text="クリア", command=self._clear_filters).grid(
            row=button_row, column=7, sticky="e", padx=2, pady=(4, 0)
        )

        frame_result = ttk.Frame(self, padding=(8, 0, 8, 0))
        frame_result.grid(row=3, column=0, sticky="nsew")
        frame_result.rowconfigure(0, weight=3)
        frame_result.rowconfigure(1, weight=1)
        frame_result.columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(frame_result, show="headings", selectmode="extended")
        self.tree.grid(row=0, column=0, sticky="nsew")
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

        yscroll = ttk.Scrollbar(frame_result, orient="vertical", command=self.tree.yview)
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll = ttk.Scrollbar(frame_result, orient="horizontal", command=self.tree.xview)
        xscroll.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        frame_detail = ttk.LabelFrame(frame_result, text="選択行の詳細", padding=6)
        frame_detail.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(8, 0))
        frame_detail.rowconfigure(0, weight=1)
        frame_detail.columnconfigure(0, weight=1)
        self.detail_text = tk.Text(frame_detail, height=8, wrap="none")
        self.detail_text.grid(row=0, column=0, sticky="nsew")
        detail_scroll = ttk.Scrollbar(frame_detail, orient="vertical", command=self.detail_text.yview)
        detail_scroll.grid(row=0, column=1, sticky="ns")
        self.detail_text.configure(yscrollcommand=detail_scroll.set)

        frame_export = ttk.LabelFrame(self, text="エクスポート", padding=8)
        frame_export.grid(row=4, column=0, sticky="ew", padx=8, pady=8)
        frame_export.columnconfigure(1, weight=1)

        ttk.Label(frame_export, text="保存先").grid(row=0, column=0, sticky="w")
        ttk.Entry(frame_export, textvariable=self.export_dir_var).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(frame_export, text="保存先選択", command=self._choose_export_dir).grid(row=0, column=2, padx=2)
        ttk.Button(frame_export, text="表示中を保存", command=lambda: self._export_rows("displayed")).grid(row=0, column=3, padx=2)
        ttk.Button(frame_export, text="選択行を保存", command=lambda: self._export_rows("selected")).grid(row=0, column=4, padx=2)

        status = ttk.Label(self, textvariable=self.status_var, relief="sunken", anchor="w")
        status.grid(row=5, column=0, sticky="ew")

    def _choose_db(self):
        path = filedialog.askopenfilename(
            title="kks_voices.db を選択",
            filetypes=[("SQLite DB", "*.db *.sqlite *.sqlite3"), ("All files", "*.*")],
        )
        if path:
            self.db_path_var.set(path)
            self._save_last_state()

    def _choose_export_dir(self):
        path = filedialog.askdirectory(title="保存先フォルダを選択")
        if path:
            self.export_dir_var.set(path)
            self._save_last_state()

    def _connect_and_load(self):
        db_path = self.db_path_var.get().strip()
        if not db_path:
            messagebox.showerror("Error", "DBパスが空です。")
            return
        if not os.path.isfile(db_path):
            messagebox.showerror("Error", f"DBが見つかりません:\n{db_path}")
            return

        try:
            if self.conn is not None:
                self.conn.close()
            self.conn = sqlite3.connect(db_path)
            self.conn.row_factory = sqlite3.Row
            self._load_table_columns()
            self._setup_table_list()
            self._on_table_changed()
            self._save_last_state()
            self.status_var.set(f"Connected: {db_path}")
        except Exception as exc:
            messagebox.showerror("Error", f"DB接続に失敗:\n{exc}")

    def _load_table_columns(self):
        self.table_columns = {}
        table_rows = self.conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name"
        ).fetchall()
        for row in table_rows:
            table = row["name"]
            cols = [r["name"] for r in self.conn.execute(f"PRAGMA table_info({table})").fetchall()]
            self.table_columns[table] = cols

    def _setup_table_list(self):
        tables = [t for t in self.table_columns.keys() if t != "sqlite_sequence"]
        if not tables:
            raise RuntimeError("テーブルが見つかりません。")
        self.table_combo["values"] = tables
        if self.table_var.get() not in tables:
            self.table_var.set(tables[0])

    def _on_table_changed(self):
        table = self.table_var.get()
        if table not in self.table_columns:
            return
        self._refresh_filter_ui_state()
        self._load_distinct_filter_values()
        self._run_query(reset_page=True)
        self._save_last_state()

    def _on_search_clicked(self):
        self._run_query(reset_page=True)
        snapshot = self._snapshot_current_query()
        self.app_state["last"] = snapshot
        self._append_history(snapshot)
        self._write_app_state()

    def _snapshot_current_query(self):
        try:
            page_size = max(1, int(self.page_size_var.get()))
        except Exception:
            page_size = 500
            self.page_size_var.set(page_size)

        snapshot = {
            "db_path": self.db_path_var.get().strip(),
            "export_dir": self.export_dir_var.get().strip(),
            "table": self.table_var.get().strip(),
            "page_size": page_size,
            "combo_filters": {k: self.combo_filter_vars[k].get().strip() for k in FILTER_COMBO_COLUMNS},
            "like_filters": {k: self.like_filter_vars[k].get().strip() for k in FILTER_LIKE_COLUMNS},
        }
        return snapshot

    def _apply_snapshot_to_vars(self, snapshot):
        if not isinstance(snapshot, dict):
            return

        db_path = str(snapshot.get("db_path", "")).strip()
        export_dir = str(snapshot.get("export_dir", "")).strip()
        table = str(snapshot.get("table", "")).strip()
        page_size = snapshot.get("page_size", 500)

        if db_path:
            self.db_path_var.set(db_path)
        if export_dir:
            self.export_dir_var.set(export_dir)
        if table:
            self.table_var.set(table)
        try:
            self.page_size_var.set(max(1, int(page_size)))
        except Exception:
            self.page_size_var.set(500)

        combo_filters = snapshot.get("combo_filters", {}) or {}
        like_filters = snapshot.get("like_filters", {}) or {}

        for k in FILTER_COMBO_COLUMNS:
            self.combo_filter_vars[k].set(str(combo_filters.get(k, "")).strip())
        for k in FILTER_LIKE_COLUMNS:
            self.like_filter_vars[k].set(str(like_filters.get(k, "")).strip())

    def _load_app_state(self):
        self.app_state = {"last": None, "history": []}
        if not APP_STATE_PATH.exists():
            return

        try:
            data = json.loads(APP_STATE_PATH.read_text(encoding="utf-8"))
            if not isinstance(data, dict):
                return
            if not isinstance(data.get("history", []), list):
                data["history"] = []
            self.app_state = data
        except Exception:
            self.app_state = {"last": None, "history": []}

    def _write_app_state(self):
        try:
            APP_STATE_PATH.parent.mkdir(parents=True, exist_ok=True)
            tmp_path = APP_STATE_PATH.with_suffix(".tmp")
            tmp_path.write_text(
                json.dumps(self.app_state, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            tmp_path.replace(APP_STATE_PATH)
        except Exception:
            pass

    def _apply_last_state_to_vars(self):
        last = self.app_state.get("last")
        if isinstance(last, dict):
            self._apply_snapshot_to_vars(last)

    def _save_last_state(self):
        self.app_state["last"] = self._snapshot_current_query()
        self._write_app_state()

    def _query_signature(self, snapshot):
        if not isinstance(snapshot, dict):
            return ""
        payload = {
            "db_path": str(snapshot.get("db_path", "")).strip(),
            "table": str(snapshot.get("table", "")).strip(),
            "page_size": int(snapshot.get("page_size", 500) or 500),
            "combo_filters": snapshot.get("combo_filters", {}) or {},
            "like_filters": snapshot.get("like_filters", {}) or {},
        }
        return json.dumps(payload, sort_keys=True, ensure_ascii=False)

    def _append_history(self, snapshot):
        history = self.app_state.get("history", [])
        if not isinstance(history, list):
            history = []

        item = {
            "saved_at": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "query": snapshot,
        }

        if history and self._query_signature(history[0].get("query")) == self._query_signature(snapshot):
            history[0] = item
        else:
            history.insert(0, item)

        if len(history) > HISTORY_MAX:
            history = history[:HISTORY_MAX]

        self.app_state["history"] = history
        self._refresh_history_listbox()

    def _open_history_window(self):
        if self.history_window is not None and self.history_window.winfo_exists():
            self.history_window.lift()
            self.history_window.focus_set()
            self._refresh_history_listbox()
            return

        win = tk.Toplevel(self)
        win.title("検索履歴")
        win.geometry("980x420")
        win.transient(self)
        win.columnconfigure(0, weight=1)
        win.rowconfigure(0, weight=1)

        listbox = tk.Listbox(win)
        listbox.grid(row=0, column=0, sticky="nsew", padx=(8, 0), pady=8)
        listbox.bind("<Double-Button-1>", lambda _e: self._apply_selected_history())

        yscroll = ttk.Scrollbar(win, orient="vertical", command=listbox.yview)
        yscroll.grid(row=0, column=1, sticky="ns", pady=8, padx=(4, 8))
        listbox.configure(yscrollcommand=yscroll.set)

        btn_frame = ttk.Frame(win)
        btn_frame.grid(row=1, column=0, columnspan=2, sticky="ew", padx=8, pady=(0, 8))
        btn_frame.columnconfigure(4, weight=1)
        ttk.Button(btn_frame, text="適用", command=self._apply_selected_history).grid(row=0, column=0, padx=2)
        ttk.Button(btn_frame, text="削除", command=self._delete_selected_history).grid(row=0, column=1, padx=2)
        ttk.Button(btn_frame, text="全削除", command=self._clear_history).grid(row=0, column=2, padx=2)
        ttk.Button(btn_frame, text="閉じる", command=win.destroy).grid(row=0, column=3, padx=2)

        win.protocol("WM_DELETE_WINDOW", win.destroy)
        self.history_window = win
        self.history_listbox = listbox
        self._refresh_history_listbox()

    def _refresh_history_listbox(self):
        if self.history_listbox is None or not self.history_listbox.winfo_exists():
            return

        self.history_listbox.delete(0, "end")
        history = self.app_state.get("history", [])
        if not isinstance(history, list):
            history = []

        for item in history:
            self.history_listbox.insert("end", self._history_item_label(item))

    def _history_item_label(self, item):
        saved_at = str(item.get("saved_at", "")).strip() or "unknown-time"
        snapshot = item.get("query", {}) or {}
        table = str(snapshot.get("table", "")).strip() or "table?"
        combo = snapshot.get("combo_filters", {}) or {}
        like = snapshot.get("like_filters", {}) or {}

        parts = []
        for k in FILTER_COMBO_COLUMNS:
            v = str(combo.get(k, "")).strip()
            if v:
                parts.append(f"{k}={v}")
        for k in FILTER_LIKE_COLUMNS:
            v = str(like.get(k, "")).strip()
            if v:
                parts.append(f"{k}~{v}")

        summary = " / ".join(parts[:4])
        if len(parts) > 4:
            summary += f" / ... ({len(parts)} filters)"
        if not summary:
            summary = "(filterなし)"

        return f"{saved_at} | {table} | {summary}"

    def _apply_selected_history(self):
        if self.history_listbox is None or not self.history_listbox.winfo_exists():
            return
        sel = self.history_listbox.curselection()
        if not sel:
            return

        idx = int(sel[0])
        history = self.app_state.get("history", [])
        if not isinstance(history, list) or idx < 0 or idx >= len(history):
            return

        snapshot = history[idx].get("query", {})
        self._apply_snapshot_to_vars(snapshot)
        self._on_table_changed()
        self._save_last_state()
        self.status_var.set("履歴を適用しました。")

    def _delete_selected_history(self):
        if self.history_listbox is None or not self.history_listbox.winfo_exists():
            return
        sel = self.history_listbox.curselection()
        if not sel:
            return

        idx = int(sel[0])
        history = self.app_state.get("history", [])
        if not isinstance(history, list) or idx < 0 or idx >= len(history):
            return

        del history[idx]
        self.app_state["history"] = history
        self._write_app_state()
        self._refresh_history_listbox()

    def _clear_history(self):
        if not messagebox.askyesno("確認", "履歴を全削除しますか？"):
            return
        self.app_state["history"] = []
        self._write_app_state()
        self._refresh_history_listbox()

    def _refresh_filter_ui_state(self):
        table = self.table_var.get()
        columns = set(self.table_columns.get(table, []))

        for col in FILTER_COMBO_COLUMNS:
            w = self.filter_widgets.get(col)
            if w is None:
                continue
            if col in columns:
                w.configure(state="readonly")
            else:
                self.combo_filter_vars[col].set("")
                w.configure(state="disabled")

        for col in FILTER_LIKE_COLUMNS:
            w = self.filter_widgets.get(col)
            if w is None:
                continue
            if col in columns:
                w.configure(state="normal")
            else:
                self.like_filter_vars[col].set("")
                w.configure(state="disabled")

    def _load_distinct_filter_values(self):
        table = self.table_var.get()
        columns = set(self.table_columns.get(table, []))
        for col in FILTER_COMBO_COLUMNS:
            widget = self.filter_widgets.get(col)
            if widget is None:
                continue
            if col not in columns:
                widget["values"] = []
                continue
            sql = (
                f"SELECT DISTINCT {col} AS value "
                f"FROM {table} "
                f"WHERE {col} IS NOT NULL AND {col} <> '' "
                f"ORDER BY {col} LIMIT 5000"
            )
            raw_values = [r["value"] for r in self.conn.execute(sql).fetchall()]
            values = []
            seen = set()
            for raw in raw_values:
                text = str(raw).strip()
                if not text:
                    continue
                key = text.casefold()
                if key in seen:
                    continue
                seen.add(key)
                values.append(text)
            widget["values"] = [""] + values

    def _build_where(self):
        table = self.table_var.get()
        columns = set(self.table_columns.get(table, []))

        conditions = []
        params = []

        for col in FILTER_COMBO_COLUMNS:
            val = self.combo_filter_vars[col].get().strip()
            if val and col in columns:
                conditions.append(f"TRIM(COALESCE({col}, '')) = ?")
                params.append(val)

        for col in FILTER_LIKE_COLUMNS:
            val = self.like_filter_vars[col].get().strip()
            if val and col in columns:
                conditions.append(f"COALESCE({col}, '') LIKE ?")
                params.append(f"%{val}%")

        where_sql = ""
        if conditions:
            where_sql = " WHERE " + " AND ".join(conditions)
        return where_sql, params

    def _resolve_order_column(self, table):
        cols = self.table_columns.get(table, [])
        for c in ["id", "idx", "voice_id", "filename"]:
            if c in cols:
                return c
        return "rowid"

    def _resolve_visible_columns(self, table):
        cols = self.table_columns.get(table, [])
        preferred = VISIBLE_COLUMN_CANDIDATES.get(table, [])
        selected = [c for c in preferred if c in cols]
        if not selected:
            selected = cols[:12]
        return selected

    def _run_query(self, reset_page=False):
        if self.conn is None:
            return

        table = self.table_var.get()
        if table not in self.table_columns:
            return

        try:
            page_size = max(1, int(self.page_size_var.get()))
        except Exception:
            page_size = 500
            self.page_size_var.set(page_size)

        if reset_page:
            self.page_var.set(1)

        where_sql, params = self._build_where()
        self.current_where_sql = where_sql
        self.current_where_params = list(params)

        count_sql = f"SELECT COUNT(*) AS c FROM {table}{where_sql}"
        total = self.conn.execute(count_sql, params).fetchone()["c"]
        self.total_rows_var.set(total)

        max_page = max(1, (total + page_size - 1) // page_size)
        current_page = min(max(1, int(self.page_var.get())), max_page)
        self.page_var.set(current_page)

        offset = (current_page - 1) * page_size
        order_col = self._resolve_order_column(table)
        sql = f"SELECT * FROM {table}{where_sql} ORDER BY {order_col} LIMIT ? OFFSET ?"
        query_params = list(params) + [page_size, offset]

        rows = self.conn.execute(sql, query_params).fetchall()
        self.current_rows = [dict(r) for r in rows]
        self.current_visible_columns = self._resolve_visible_columns(table)

        self._populate_tree()

        self.page_label.configure(text=f"Page {current_page} / {max_page}")
        self.total_label.configure(text=f"Total: {total}")
        self.status_var.set(f"Loaded {len(rows)} rows from {table} (matched: {total})")

    def _populate_tree(self):
        self.tree.delete(*self.tree.get_children())

        cols = self.current_visible_columns
        self.tree["columns"] = cols

        for c in cols:
            self.tree.heading(c, text=c)
            width = 120
            if c in ("serif", "wav_path"):
                width = 480
            elif c in ("filename", "mode_name", "file_type", "insert_type", "houshi_type", "aibu_type", "situation_type"):
                width = 170
            self.tree.column(c, width=width, minwidth=80, stretch=False)

        for idx, row in enumerate(self.current_rows):
            values = []
            for c in cols:
                v = row.get(c, "")
                if v is None:
                    v = ""
                s = str(v).replace("\r", " ").replace("\n", " ")
                values.append(s)
            self.tree.insert("", "end", iid=str(idx), values=values)

        self.detail_text.delete("1.0", "end")

    def _on_tree_select(self, _event):
        selected = self.tree.selection()
        if not selected:
            self.detail_text.delete("1.0", "end")
            return

        idx = int(selected[0])
        if idx < 0 or idx >= len(self.current_rows):
            return

        row = self.current_rows[idx]
        lines = [f"{k}: {row.get(k, '')}" for k in self.table_columns.get(self.table_var.get(), [])]
        self.detail_text.delete("1.0", "end")
        self.detail_text.insert("1.0", "\n".join(lines))

    def _prev_page(self):
        cur = max(1, int(self.page_var.get()))
        if cur > 1:
            self.page_var.set(cur - 1)
            self._run_query(reset_page=False)

    def _next_page(self):
        cur = max(1, int(self.page_var.get()))
        self.page_var.set(cur + 1)
        self._run_query(reset_page=False)

    def _clear_filters(self):
        for var in self.combo_filter_vars.values():
            var.set("")
        for var in self.like_filter_vars.values():
            var.set("")
        self._run_query(reset_page=True)

    def _rows_for_export(self, mode):
        if mode == "selected":
            selected = self.tree.selection()
            rows = []
            for iid in selected:
                idx = int(iid)
                if 0 <= idx < len(self.current_rows):
                    rows.append(self.current_rows[idx])
            return rows
        return list(self.current_rows)

    def _build_relative_export_path(self, row):
        table = self.table_var.get()
        chara = sanitize_segment(row.get("chara"))

        mode_name = row.get("mode_name")
        mode_num = row.get("mode")
        if mode_name:
            mode_segment = sanitize_segment(mode_name)
        elif mode_num is not None:
            mode_segment = f"mode_{mode_num}"
        else:
            mode_segment = "mode_unknown"

        level_name = row.get("level_name")
        level_num = row.get("level")
        if level_name:
            level_segment = sanitize_segment(level_name)
        elif level_num is not None:
            level_segment = f"level_{level_num}"
        else:
            level_segment = "level_unknown"

        category = (
            row.get("file_type")
            or row.get("breath_type")
            or row.get("houshi_type")
            or row.get("aibu_type")
            or row.get("situation_type")
            or "voice"
        )
        category_segment = sanitize_segment(category)

        source = str(row.get("wav_path") or "")
        ext = Path(source).suffix if Path(source).suffix else ".wav"
        filename = sanitize_segment(row.get("filename") or f"id_{row.get('id', 'unknown')}")

        return Path(table) / chara / mode_segment / level_segment / category_segment / f"{filename}{ext}"

    def _unique_destination_path(self, dest_root, row):
        rel = self._build_relative_export_path(row)
        dst = dest_root / rel
        if not dst.exists():
            return dst

        stem = dst.stem
        suffix = dst.suffix
        row_id = row.get("id") if row.get("id") is not None else row.get("voice_id")
        row_id = sanitize_segment(row_id)
        for n in range(1, 1000):
            alt = dst.with_name(f"{stem}_id{row_id}_{n}{suffix}")
            if not alt.exists():
                return alt
        return dst.with_name(f"{stem}_{dt.datetime.now().strftime('%H%M%S%f')}{suffix}")

    def _build_voice_text_row(self, row):
        filename = str(row.get("filename") or "").strip()
        if not filename:
            src = str(row.get("wav_path") or "").strip()
            filename = Path(src).name if src else f"id_{row.get('id', 'unknown')}.wav"
        if not Path(filename).suffix:
            filename = f"{filename}.wav"

        chara = str(row.get("chara") or "").strip() or "unknown"
        serif = str(row.get("serif") or "")
        serif = serif.replace("\r\n", "\n").replace("\r", "\n").replace("\n", " ")
        return [filename, chara, "JP", serif]

    def _export_rows(self, mode):
        rows = self._rows_for_export(mode)
        if not rows:
            messagebox.showinfo("Info", "保存対象がありません。")
            return

        dest = self.export_dir_var.get().strip()
        if not dest:
            messagebox.showerror("Error", "保存先が空です。")
            return

        dest_root = Path(dest)
        dest_root.mkdir(parents=True, exist_ok=True)

        copied = 0
        missing = 0
        failed = 0
        duplicate_skipped = 0
        manifest_rows = []
        voice_text_rows = []
        seen_sources = set()

        for row in rows:
            src = row.get("wav_path")
            if not src:
                missing += 1
                continue
            src_norm = os.path.normcase(os.path.normpath(str(src)))
            if src_norm in seen_sources:
                duplicate_skipped += 1
                continue
            if not os.path.isfile(src):
                missing += 1
                continue

            dst = self._unique_destination_path(dest_root, row)
            dst.parent.mkdir(parents=True, exist_ok=True)
            try:
                shutil.copy2(src, dst)
                copied += 1
                seen_sources.add(src_norm)
                manifest_rows.append(
                    {
                        "table": self.table_var.get(),
                        "id": row.get("id"),
                        "chara": row.get("chara"),
                        "mode_name": row.get("mode_name"),
                        "level_name": row.get("level_name"),
                        "filename": row.get("filename"),
                        "source_wav_path": src,
                        "exported_path": str(dst),
                    }
                )
                voice_text_rows.append(self._build_voice_text_row(row))
            except Exception:
                failed += 1

        stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        manifest_path = dest_root / f"export_manifest_{self.table_var.get()}_{stamp}.csv"
        with manifest_path.open("w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(
                f,
                fieldnames=[
                    "table",
                    "id",
                    "chara",
                    "mode_name",
                    "level_name",
                    "filename",
                    "source_wav_path",
                    "exported_path",
                ],
            )
            writer.writeheader()
            writer.writerows(manifest_rows)

        voice_text_path = dest_root / f"export_voice_text_{self.table_var.get()}_{stamp}.csv"
        with voice_text_path.open("w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f, delimiter="|", lineterminator="\n")
            writer.writerows(voice_text_rows)

        summary = (
            f"保存完了\n"
            f"- 対象行: {len(rows)}\n"
            f"- 保存成功: {copied}\n"
            f"- 重複スキップ: {duplicate_skipped}\n"
            f"- ソース不足/未発見: {missing}\n"
            f"- 失敗: {failed}\n"
            f"- マニフェスト: {manifest_path}"
        )
        self.status_var.set(summary.replace("\n", " | "))
        messagebox.showinfo("Export", summary)

    def destroy(self):
        try:
            if self.conn is not None:
                self.conn.close()
                self.conn = None
        finally:
            super().destroy()


def main():
    app = KksVoiceDbGui()
    app.mainloop()


if __name__ == "__main__":
    main()
