---
title: "Word, Excel, and PowerPoint Format Compatibility"
description: "Check OfficeIMO support for DOC, DOCX, XLS, XLSX, XLSB, PPT, PPTX, templates, slide shows, macros, conversion directions, preservation, and fidelity."
layout: compatibility
slug: compatibility
meta.social_card_badge: "Format compatibility"
meta.social_card_metrics: "28|Office variants|package;137|Tracked behaviors|check;3|Format families|code"
---

## Choose the result your workflow needs

The safest conversion is not always the one that produces a file at any cost. OfficeIMO lets an application choose whether it requires native representation, accepts an editable approximation, prefers visual fidelity, permits a documented best-effort result, or retains the complete source for recovery.

Use `AnalyzeConversion(...)` before writing when a service needs an approval gate. The returned report identifies the affected feature, source location when available, chosen fallback, and fidelity dimensions at risk. `RequireNoLoss()` turns any known loss or blocked feature into a failed operation.

Evidence is grouped by concrete format contract. Excel XLS and XLSB have separate rows and counts because BIFF8 and binary Open XML do not share the same implementation or fidelity boundary.

## Important boundaries

- A listed extension means OfficeIMO can classify and route the format. It does not mean every format is a writable destination.
- Legacy Excel `.xlt`, `.xla`, `.xlm`, and `.xlw` authoring and PowerPoint `.ppa` authoring are currently blocked by structured preflight.
- Visual fallbacks preserve appearance but are not presented as editable formulas, charts, SmartArt, or document structure.
- Embedded source payloads may contain macros, linked content, or other active material. Treat recovered bytes as untrusted input.
- External LibreOffice and optional Microsoft Office checks are validation oracles, not runtime dependencies.

Read the [compatibility contract and validation guide](https://github.com/EvotecIT/OfficeIMO/tree/master/Docs/Compatibility) for the complete generated rows, corpus policy, and interoperability gates.
