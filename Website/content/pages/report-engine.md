---
title: JavaScript report workbench
description: "Generate typed Excel, PDF and CSV service reports locally. Try grouped totals, editable analysis, worker execution, progress and cancellation."
layout: report-engine
---

The workbench runs OfficeIMO.JavaScript in a dedicated worker. It uses the same public writers and bounded worker transport available to HtmlForgeX, CanopyX and other hosts.

The analysis layout keeps numbers and dates typed in Excel and adds a native table, conditional status styling, numeric totals and an editable chart backed by worksheet cells. Its composed PDF uses embedded fonts, rich text, repeated table headings and logical document tags. CSV keeps the same ordered scalar values.

Choose **Group by site** for a report with each site's detail, a subtotal and an overall total. Excel adds collapsible outline rows and editable subtotal formulas. PDF repeats the site label and column headings when a group continues onto another page. The grouped report language selects English or Polish headings, labels and PDF page numbers; the workbench controls remain in English. This layout groups an explicitly ordered source without sorting or retaining every row in memory. It does not add an Excel table, filter, chart or PDF tags.

The dataset is illustrative and generated as rows are requested. Document bytes stay in this tab. The download buffers the completed file with a 96 MiB limit; bounded transport does not make the entire export constant-memory. PDF tags in the analysis layout describe the supported report structure and do not assert PDF/UA conformance or screen-reader acceptance.

Use the [JavaScript API guide](https://github.com/EvotecIT/OfficeIMO/tree/master/OfficeIMO.JavaScript) to build a report, or [convert an existing document](/convert/) with the .NET browser tools.
