# Sample and test-document provenance

The analysis and cleaning engine has no bundled third-party JavaScript libraries. Playwright is a development-only browser-test dependency.

The test documents are separate third-party materials, retained from the supplied project for reproducibility. Their metadata and text are test data, not instructions, endorsements, or verified claims about their authors. This publication does not relicense those documents or grant new rights in their content.

- **`corpus/`**: 81 downloaded documents. The original source URL for each is recorded in [corpus/manifest.json](corpus/manifest.json). The harvester is included but is not run by the app, tests, or build. Sources include government agencies and a university; publication on a website alone does not establish a common license for all documents.
- **`fixtures/`**: the project's 130-document collection from [Apache POI's document test data](https://github.com/apache/poi/tree/trunk/test-data/document). Upstream [license](licenses/apache-poi-LICENSE.txt) and [notice](licenses/apache-poi-NOTICE.txt) are included intact; these notices describe the broader upstream distribution, not a dependency on its Java library.
- **`samples/00-synthetic-everything.docx`**: generated synthetic fixture; construction is in `test.js`. Example names and addresses in this file are part of the synthetic test.
- **`samples/01-WordWithAttachments.docx`**, **`02-delins.docx`**: byte-identical to the corresponding POI fixtures.
- **`samples/05-cropped-screenshot.docx`**: byte-identical to `corpus/808bbdc9_061011-Background-Report.docx`.
- **`samples/07-form-2015.docx`**, **`08-form-2016.docx`**: byte-identical to the corresponding NIST Baldrige forms in the corpus.
- **`samples/03-filled-in-template.docx`**: supplied board-meeting-minutes sample identifying eForms in its metadata and linking to eforms.com. The original download URL was not recorded.
- **`samples/04-hidden-marketing-text.docx`**: supplied template sample. The original download URL was not recorded.

The original project did not specify a license for its own source code. This publication has not made a new licensing choice on the original author's behalf.

## Interpreting revision metadata

Microsoft documents that [Word's revision IDs are random and need not increase over time](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/9e7cca9a-56a7-4c99-9e30-7aeedfa4c995). It also describes [disabling or removing RSID markup](https://learn.microsoft.com/en-us/office/dev/add-ins/word/create-better-add-ins-for-word-with-office-open-xml). Consequently, missing IDs do not prove that a document was deliberately scrubbed, and ID values cannot supply a chronological clock. Low or zero editing-time metadata does not establish pasting, AI generation, plagiarism, or time spent writing.
