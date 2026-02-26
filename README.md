# KKS Voices DB GUI

`kks_voices.db` をGUIで閲覧・絞り込み・エクスポートするツール。

## 対象ファイル
- DB: `F:\kks\work\kks_voices.db`
- スクリプト: `F:\kks\work\VoiceDbGui\kks_voices_gui.py`
- 起動バッチ: `F:\kks\work\VoiceDbGui\run_kks_voices_gui.bat`

## 起動
1. `run_kks_voices_gui.bat` を実行
2. もしくは PowerShell で:
   - `python F:\kks\work\VoiceDbGui\kks_voices_gui.py`

## 機能
1. テーブル選択 (`voices`, `breaths`, `shortbreaths`, `cond_desc`)
2. フィルタ絞り込み
   - `chara`, `mode_name`, `level_name`, `file_type`, `insert_type`, `houshi_type`, `aibu_type`, `situation_type`, `breath_type` (一致)
   - `filename`, `serif`, `wav_path` (部分一致)
3. ページング表示
   - 1ページ件数を変更可能
4. 条件保存と履歴
   - 前回検索条件を自動保存し、次回起動時に復元
   - 検索ボタン押下時に履歴へ保存
   - `履歴` ボタンから過去条件を適用/削除/全削除
5. エクスポート
   - `表示中を保存`: 現在表示されているページの行を保存
   - `選択行を保存`: 表で選択した行だけ保存
   - 同一 `wav_path` の行は重複として自動スキップ（同じ音声の重複保存を防止）

## エクスポート先フォルダ構成
保存先配下に以下規則で出力:

`<table>/<chara>/<mode_name or mode_x>/<level_name or level_x>/<category>/<filename>.wav`

- `category` は `file_type` などを優先して自動決定
- 同名衝突時は `_id..._n` を付与
- 同時に `export_manifest_<table>_<timestamp>.csv` を出力

## 注意
1. `wav_path` が存在しない行はスキップされる
2. `cond_desc` は音声パス列がないため、実質閲覧用途
