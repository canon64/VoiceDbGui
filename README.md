# KKS Voice Studio

KKS (Koikatsu Sunshine) の音声ファイルを抽出・DB化・閲覧・エクスポートする統合ツール。

## 機能

### タブ1: 抽出
- AssetBundle から WAV ファイルを抽出
- キャラクター単位で選択可能
- UnityPy を使用

### タブ2: DB構築
- 抽出済み WAV から SQLite DB を構築
- VoicePatternData から挿入位置・奉仕種別・愛撫種別・シチュエーション種別を自動取得
- セリフ CSV があれば字幕を付与

### タブ3: ブラウズ
- DB を絞り込み・ページング表示
- フィルタ: キャラ・モード・レベル・種別など
- キャラ名を日本語表示（`voice_extract/character_map.json` 参照）
- 表示中 or 選択行を WAV エクスポート
  - フォルダ階層モード / フラット（1フォルダ）モード
- voice_text CSV 同時出力（チェックボックスで切り替え）
- エクスポート後に出力先フォルダを自動で開く
- 検索条件を履歴として保存・復元

## 必要環境

- Python 3.8+
- [UnityPy](https://github.com/K0lb3/UnityPy) (`pip install UnityPy`) ※抽出タブのみ必要

```bash
pip install UnityPy
```

## 起動

```bash
# CMD ウィンドウなし
kks_voice_studio.vbs

# CMD あり（デバッグ用）
kks_voice_studio.bat
```

## 設定

初回起動時は「抽出」タブで KKS フォルダを指定するだけで、WAV 出力先・DB パス・エクスポート先が自動設定されます。

| パス | デフォルト |
|------|-----------|
| WAV 出力先 | `{KKSフォルダ}/wave` |
| DB | `{KKSフォルダ}/wave/kks_voices.db` |
| エクスポート先 | `{KKSフォルダ}/extract_wave` |

## キャラクター名マッピング

`{KKSフォルダ}/voice_extract/character_map.json` に以下の形式で配置すると、
GUI のドロップダウンやチェックボックスにキャラ名が表示されます。

```json
{
  "c00": "00 セクシー系お姉さま",
  "c13": "13 ギャル"
}
```

## ライセンス

MIT
