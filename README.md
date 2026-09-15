# Palimpsest

**Inspect what a Word document still carries: revision groups, metadata, recoverable deletions, embedded files, and pixels outside image crops. Then choose what to remove.**

[Open the live app](https://parttimediarybehavior.github.io/palimpsest/) · [Download the offline release](https://github.com/parttimediarybehavior/palimpsest/releases/latest)

The application is one self-contained HTML file with eight embedded examples. Open `dist/palimpsest.html` in a modern browser, choose a sample, or drop a `.docx` onto the page. Analysis and cleaning run locally. No account, upload, API key, or server is required. Loading the hosted page downloads the app; processing a document makes no network request.

## What you can do

- Inspect authorship properties, recorded editing time, comments, tracked deletions, hidden-text candidates, language tags, and embedded objects.
- Color the text by recorded revision ID and inspect groups that are absent from the document's session list.
- Compare two versions using shared revision IDs and paragraph IDs, with an inferred relationship and text differences.
- Reveal the full stored image behind a Word crop and inspect supported image metadata.
- Select cleaning categories, download a rewritten `.docx`, and inspect the tool's before/after analysis. Embedded-object and macro removal are opt-in because they remove content. Image crop baking is available in the browser; the CLI reports that it cannot perform that step.

## What the evidence means

Revision-save IDs (`w:rsid`) are useful grouping clues, not a complete edit log. Their values do not provide a reliable clock. IDs may be inherited, omitted, disabled, removed, or rewritten. An ID absent from `settings.xml` can suggest imported content but does not identify a unique source document.

Recorded editing minutes are a document property, not a reliable measurement of typing or labor. A low value does not prove pasting, AI generation, or plagiarism. Sparse metadata does not prove intentional cleaning. Dated revisions and author fields are recorded claims in a mutable file, not authenticated facts. [Technical sources and sample provenance](SOURCES.md).

Cleaning changes only the supported, selected categories. Some selections remove content. The tool re-analyzes its output and the tests check container consistency and extracted-text preservation; **Microsoft Word or LibreOffice has not independently verified rendering or layout preservation**. Inspect the cleaned copy before relying on it. Encrypted files are rejected.

## Run from the command line

Node.js 18 or later; runtime code has no npm dependencies.

```sh
node cli.js file.docx
node cli.js file.docx --json
node cli.js file.docx --scrub out.docx
node scan.js corpus/
node test.js
node validate-lineage.js
node scrub-test.js corpus/
node scrub-test.js fixtures/
node build.js
```

The build combines the local source and embedded samples into `dist/palimpsest.html`. It does not fetch documents or libraries. `corpus/harvest.py` is an optional network-based research utility, never invoked automatically.

To run the browser checks, use Node.js 20 or later:

```sh
npm ci
npx playwright install chromium
npm run test:browser
```

Those two installation steps require internet access. The browser test itself opens the local HTML offline and checks analysis, selective cleaning, crop baking, comparison, and downloads.

## Verification

The original synthetic suite has **88 passing assertions**. The lineage test recovered the expected direction for **28 of 28 directed pairs** of dated NIH and NIST files. The scrub harness produced outputs passing its container and extracted-text checks for **81/81 corpus files and 115/115 parseable POI fixtures**. The other 15 POI inputs are outside the accepted document format or malformed; they are not counted as successful scrubs.

These figures were rechecked for publication. Fresh outputs and browser evidence are in [docs/verification/](docs/verification/). The original broader measurements remain in [docs/corpus-report.txt](docs/corpus-report.txt) and [docs/scrub-report.txt](docs/scrub-report.txt); their earlier inference wording is superseded by the qualified interpretation above. [Publication changes](docs/PUBLICATION_REVIEW.md).

## Source map

| File | Purpose |
|---|---|
| `dist/palimpsest.html` | Complete offline app |
| `src/zipxml.js`, `src/zipwrite.js` | ZIP and XML handling |
| `src/media.js` | Image geometry and supported metadata |
| `src/engine.js` | Analysis, revision groups, comparison, reports |
| `src/scrub.js` | Selective rewrites and output checks |
| `src/ui.js`, `src/index.html`, `src/app.css` | Browser interface |
| `test.js`, `scrub-test.js`, `validate-lineage.js` | Synthetic and corpus checks |
| `samples/`, `fixtures/`, `corpus/` | Supplied examples and test data; see `SOURCES.md` |

## Valuable next work

Render original and cleaned files in a real Office application, compare their pages, and investigate differences before claiming layout preservation. Improve handling of individual passages and expand the independently verified cleaning cases. The present implementation focuses on Word OOXML documents.
