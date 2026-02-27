# KKS Voice Studio

An all-in-one tool for extracting, indexing, browsing, and exporting voice files from Koikatsu Sunshine (KKS).

## Features

### Tab 1: Extract
- Extract WAV files from AssetBundles using UnityPy
- Select characters individually

### Tab 2: Build DB
- Build a SQLite database from extracted WAV files
- Automatically resolves insert / service / caress / situation types from VoicePatternData
- Attaches subtitles if a voice CSV is present

### Tab 3: Browse
- Filter, paginate, and inspect the database
- Filters: character, mode, level, type, etc.
- Japanese character names shown in UI (reads `voice_extract/character_map.json`)
- Export displayed or selected rows as WAV files
  - Structured folder mode or flat (single folder) mode
- Optional voice_text CSV export alongside WAV files
- Automatically opens the export folder on completion
- Search history saved and restored across sessions

## Requirements

- Python 3.8+
- [UnityPy](https://github.com/K0lb3/UnityPy) — only required for the Extract tab

```bash
pip install UnityPy
```

## Launch

```bash
# No console window
kks_voice_studio.vbs

# With console (for debugging)
kks_voice_studio.bat
```

## Configuration

Set the KKS root folder in the Extract tab on first launch. All other paths are derived automatically:

| Path | Default |
|------|---------|
| WAV output | `{KKS folder}/wave` |
| Database | `{KKS folder}/wave/kks_voices.db` |
| Export destination | `{KKS folder}/extract_wave` |

## Character Name Mapping

Place a `character_map.json` file at `{KKS folder}/voice_extract/character_map.json` to display character names in the UI:

```json
{
  "c00": "00 Sexy Older Sister",
  "c13": "13 Gyaru"
}
```

## License

MIT
