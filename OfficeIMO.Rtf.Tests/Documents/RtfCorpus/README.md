# OfficeIMO RTF interoperability corpus

This folder contains small RTF fixtures with checked provenance and executable conversion expectations. `corpus-manifest.json` is the source of truth for fixture ids, producers, versions, origins, licenses, SHA-256 hashes, required controls/text, adapter coverage, producer scorecard status, and reopen evidence.

The current corpus includes:

- a Microsoft Word 16 document generated locally and reopened in Word after OfficeIMO normalization;
- a Microsoft Outlook 16 message saved as RTF;
- four LibreOffice upstream regression fixtures pinned to an exact source commit;
- a synthetic Outlook HTML-encapsulation grammar fixture, labeled synthetic;
- focused synthetic files for core syntax, formatting, lists, tables, images, notes/fields, layout, code pages, and pathological input.

Every `.rtf` file must appear in the manifest with a stable hash and redistribution permission. `RtfGoldenCorpusTests` verifies exact source bytes, normalized reparse, required semantic text/control words, executable adapter claims, producer scorecard honesty, and reopen evidence.

Do not relabel synthetic grammar coverage as producer evidence. Google Docs, macOS TextEdit/RTFD, EHR/CRM/helpdesk generators, and commercial-library output remain `unverified` until redistributable source files and reproducible reopen evidence are added.
