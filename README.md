# CDSA Lecture (`cdsappt`)

CDSA Lecture is a Codex plugin for producing editable PowerPoint lecture decks with one CDSA master and 13 tested layouts. It searches a PowerPoint corpus that you provide, reuses selected text, tables, diagrams and images as editable slide objects, and checks both package structure and rendered slides.

## What is included

- `cdsappt` skill and PowerPoint authoring harness
- CDSA master with 13 layouts and preview images
- BM25 corpus indexer for Korean and English PowerPoint content
- editable object reuse, cover, logo, structure-check and PowerPoint-render tools
- public Codex plugin marketplace metadata

The Git repository excludes the private 78-deck corpus and the CDSA lecture knowledge base. They are not required to run the engine. A full original archive can be created locally when needed, but it is not published with this repository.

## Install

Clone or download the repository. The installable plugin is in `plugins/cdsa-lecture`; the repository also contains `.agents/plugins/marketplace.json` for Codex plugin discovery.

```powershell
git clone https://github.com/kyj2294/cdsa-lecture.git
cd cdsa-lecture
```

Register this repository as a plugin source in Codex, then install **CDSA Lecture**. For a manual local setup, keep the whole `plugins/cdsa-lecture` directory together because the skill calls scripts at the plugin root.

## Build your corpus

Point the setup script at a folder containing `.pptx` or `.pptm` files that you may reuse:

```powershell
powershell -ExecutionPolicy Bypass -File .\plugins\cdsa-lecture\scripts\setup-corpus.ps1 -Source "C:\path\to\your-ppt-folder"
$env:CDSA_LECTURE_DB = "$env:LOCALAPPDATA\cdsa-lecture\lecture-library.sqlite"
```

Search it directly to confirm the setup:

```powershell
python .\plugins\cdsa-lecture\skills\cdsappt\scripts\lecture_library.py search "AI 리터러시" --limit 5
```

Use `$cdsappt` in Codex and describe the audience, duration, topic, institution and any source files you want included.

## Validate

```powershell
python -m unittest discover -s .\plugins\cdsa-lecture\scripts -p "test_*.py"
```

The full release gate also requires Microsoft PowerPoint on Windows because final visual review uses actual PowerPoint renders.

Read [KNOWN_LIMITATIONS.md](KNOWN_LIMITATIONS.md) before distributing generated decks.
