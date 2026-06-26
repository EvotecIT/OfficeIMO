# OfficeIMO Legacy XLS Diagnostic Corpus

This folder is for small `.xls` fixtures that are expected to produce import
errors or hard file-format blockers. These files are intentionally separate from
`LegacyXlsCorpus`, whose fixtures should import without errors and participate in
desktop Excel open/import/open validation.

For each `sample.xls`, keep an approved `sample.import-report.md` generated from
`LegacyXlsImportReport.ToMarkdown()`. Refresh baselines with the same
`OFFICEIMO_UPDATE_LEGACY_XLS_CORPUS_BASELINES=1` workflow used by the normal
legacy XLS corpus.
