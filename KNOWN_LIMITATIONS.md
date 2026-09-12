# Known limitations and recorded gaps

## Distribution

- The Git repository does not bundle the private 78-deck corpus or the CDSA knowledge base. Corpus search and source-slide reuse require a user-built database. Any full original archive is kept locally and must be reviewed independently before redistribution or reuse.
- No open-source license has been granted in this release. The repository is publicly downloadable, but copyright remains with the respective owners. A formal license should be added before accepting external redistribution or contributions.
- The bundled CDSA master has embedded font binaries removed. PowerPoint may substitute fonts that are not installed on the user's computer, so spacing can change. Always render and inspect the final deck on the delivery machine.

## Runtime

- `.ppt` files are not indexed directly. Convert legacy files to `.pptx` first.
- Final render verification requires desktop Microsoft PowerPoint on Windows. Structural checks alone do not detect every clipping, overlap or font-substitution problem.
- Corpus databases contain extracted slide text, local source paths and copied package parts. Store them outside the repository and do not publish them without reviewing source rights and confidential data.
- Search ranking proposes candidates. A human or agent still has to compare the slide purpose, wording, visual structure and source rights before reuse.

## Content and visual quality

- Reused slides may retain old organization names, dates, footers or course labels inside body objects. The render-review gate must remove or update them.
- Very dense source slides can require splitting or a different layout. The automated scale and object-count warnings are safeguards, not a guarantee of legibility.
- Generated covers and externally sourced logos need their own provenance and usage-rights records in `assets.json`.

## Validation findings fixed in this release

- Renamed the corpus setup parameter from `-Db` to `-Database` because PowerShell interpreted `-Db` as a conflicting abbreviation of the common `-Debug` parameter.
- Corrected visual export so its JSON manifest is printed to stdout instead of overwriting the exported `.pptx` or image file.
- Corrected custom corpus paths so the post-ingest summary is written beside the selected database.
