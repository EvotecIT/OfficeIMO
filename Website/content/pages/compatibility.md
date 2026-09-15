---
title: "OfficeIMO Format and Operation Compatibility"
description: "Check package-owned document operations and detailed Word, Excel, and PowerPoint format fidelity in one generated compatibility view."
layout: compatibility
slug: compatibility
meta.social_card_badge: "Format compatibility"
meta.social_card_metrics: "63|Package owners|package;1310|Operation outcomes|check;3|Detailed legacy families|code"
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
