# Publication verification

These checks were run on 2026-09-15 against the publication copy of the supplied project.

- `synthetic.txt`: all 88 assertions pass, including selective cleaning and container relationships.
- `lineage.txt`: all 28 directed comparisons agree with the supplied dated NIH/NIST examples.
- `corpus.txt`: 81/81 outputs pass the scrub harness; four findings remain under default CLI options.
- `fixtures.txt`: 115/115 parseable fixtures pass the scrub harness; six findings remain under default CLI options. The directory includes 15 other inputs that are rejected or outside the supported format, and are not counted as successful scrubs.
- `browser.json`: offline Chromium checks, including actual downloaded DOCX re-analysis, image crop baking, version comparison, malformed input, and a 390px report viewport.
- `browser-analysis.png`, `browser-mobile.png`: actual browser screenshots.
- `browser-cleaned-synthetic.docx`: the document downloaded through the browser's cleaning flow.

The corpus outputs were captured during the initial publication check; subsequent changes qualify report wording and repair browser layout, without changing extraction or scrub algorithms. The synthetic suite and browser checks cover the final publication build. The source archive retains the original broader corpus reports separately under `docs/`.

The container checker is a set of XML/relationship/content-type checks, not a complete OOXML conformance validator. Text checks compare extracted content rather than rendered pages. No Word/LibreOffice rendering check was performed. See the main README for interpretation limits and reproduction commands.
