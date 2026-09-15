# Publication review — 2026-09-15

This release was prepared from the supplied Palimpsest project in a separate copy. The original directory was preserved. Runtime algorithms for parsing, metadata extraction, comparison, and document rewriting were retained.

Changes for publication:

- Qualified claims that editing-time ratios or revision IDs prove pasting, authorship, a complete writing history, or deliberate scrubbing. The numerical signals remain available; explanatory wording now distinguishes observations from inference.
- Removed assertions that cleaning guarantees identical rendered layout or removes every possible trace. The existing harness checks container relationships and extracted text; it does not run Microsoft Word or LibreOffice.
- Removed the optional query-string and `data-url` remote sample loaders. Built-in examples are embedded; opening a crafted URL should not silently fetch a remote document.
- Fixed narrow-screen report overflow found by the browser check: constrained grid columns and wrapped report actions and long metadata values.
- Added public entry links, test commands, source attribution, upstream POI notices, browser verification, and release packaging. No new license was assigned to the original source.

The two synthetic assertions that matched old headline text were updated to check the qualified wording. No test was removed. The original corpus reports are retained; fresh publication checks are in `docs/verification/`.

This is an inspection and selective-cleaning tool. Its own re-analysis is useful evidence about supported fields, not an independent Office-rendering or comprehensive sanitization guarantee.
