---
title: JavaScript report workbench
description: "Generate typed Excel, readable wide PDF and CSV reports locally. Try hierarchical analysis, worker exports, progress and cancellation."
layout: report-engine
---

The workbench runs OfficeIMO.JavaScript in a dedicated worker. It uses the same public writers and bounded worker transport available to HtmlForgeX, CanopyX and other hosts.

The analysis layout keeps numbers and dates typed in Excel and adds a native table, conditional status styling, numeric totals and an editable chart backed by worksheet cells. Its composed PDF uses embedded fonts, rich text, repeated table headings and logical document tags. CSV keeps the same ordered scalar values.

Choose **Group by site** for a report with each site's detail, a subtotal and an overall total. Excel adds collapsible outline rows and editable subtotal formulas. PDF gives each group a heading and navigation entry, repeats context on continuation pages and associates summary values with their column and row headings. The grouped report language selects English or Polish headings, labels and PDF page numbers; the workbench controls remain in English. This layout groups an explicitly ordered source without sorting or retaining every row in memory. It does not add an Excel table, filter or chart.

Choose **Hierarchical analysis** for a selected-view report grouped by portfolio, environment and owner. Excel includes report context, three collapsible grouping levels, parent totals and an editable overview chart. Status rules and allowed-value lists apply to detail rows. PDF places the grouped appendix between its opening context and recommendations. Nested sections and navigation entries preserve the group hierarchy; each subtotal has its own headed table under the group it describes. This layout supports English or Polish headings, report labels and page numbers. CSV remains a flat export of the visible values.

Choose **Wide analysis** to add request volume, error rate, resource usage, certificate expiry, version and incident fields. Excel keeps all 16 typed detail columns together. PDF keeps readable column widths by splitting the same row window into consecutive column bands, repeating service IDs and group context in each band. Its logical table reads each source cell once in row order, including cells continued on another page. It consumes the ordered source once; the report does not sort or replay records for each band. English and Polish report resources are available. Selecting CSV restores the ordinary flat six-field export.

The dataset is illustrative and generated as rows are requested. Document bytes stay in this tab. The download buffers the completed file with a 96 MiB limit; bounded transport does not make the entire export constant-memory. PDF tags describe the report's logical structure. Reader and assistive-technology behavior still needs qualification for each named workflow; tags alone do not establish accessible PDF or PDF/UA conformance.

Use the [JavaScript API guide](https://github.com/EvotecIT/OfficeIMO/tree/master/OfficeIMO.JavaScript) to build a report, or [convert an existing document](/convert/) with the .NET browser tools.
