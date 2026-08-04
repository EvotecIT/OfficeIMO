# OfficeIMO RTF interoperability corpus

This folder contains small RTF fixtures with checked provenance and executable conversion expectations. `corpus-manifest.json` is the source of truth for fixture ids, producers, versions, origins, licenses, SHA-256 hashes, required controls/text, adapter coverage, producer scorecard status, and reopen evidence.

The current corpus includes:

- a Microsoft Word 16 document generated locally and reopened in Word after OfficeIMO normalization;
- a Microsoft Outlook 16 message saved as RTF;
- four LibreOffice upstream regression fixtures pinned to an exact source commit;
- a synthetic Outlook HTML-encapsulation grammar fixture, labeled synthetic;
- focused synthetic files for core syntax, formatting, lists, tables, images, notes/fields, layout, code pages, and pathological input.

Every `.rtf` file must appear in the manifest with a stable hash and redistribution permission. `RtfGoldenCorpusTests` verifies exact source bytes, normalized reparse, required semantic text/control words, executable adapter claims, producer scorecard honesty, and reopen evidence.

Do not relabel synthetic grammar coverage as producer evidence. `corpus-manifest.json` records five separate external paths: genuine Google Docs and macOS TextEdit/RTFD output, a redacted Epic EHI export, a Salesforce configuration-report workflow, and a helpdesk RichEdit artifact. The CRM and helpdesk artifacts are workflow evidence, not vendor-native exports. The checked GemBox.Document fixture is reproducibly generated from the exact package version under `Build/ProducerCorpus`.

`Build/Test-RtfExternalProducerEvidence.ps1` downloads each pinned artifact, verifies its byte/hash or semantic provenance, and then exercises the bounded RTF reader, web-safe HTML, Markdown, and diagnostic-preserving Word bridge. Third-party bytes remain external and are not redistributed.
