---
title: Report workbench
description: "Generate typed Excel, PDF and CSV service reports locally. Try editable charts, rich PDF tables, worker execution, progress and cancellation on the same dataset."
layout: report-engine
---

The workbench runs OfficeIMO.JavaScript in a dedicated worker. It uses the same public writers and bounded worker transport available to HtmlForgeX, CanopyX and other hosts.

Excel keeps numbers and dates typed and adds a native table, conditional status styling, numeric totals and an editable chart backed by worksheet cells. PDF uses embedded fonts, rich text, repeated table headings and logical document tags. CSV keeps the same ordered scalar values.

The dataset is illustrative and generated as rows are requested. Document bytes stay in this tab. The download buffers the completed file; bounded transport does not make the entire export constant-memory. PDF tags describe the supported report structure and do not assert PDF/UA conformance or screen-reader acceptance.

Use the [JavaScript API guide](https://github.com/EvotecIT/OfficeIMO/tree/master/OfficeIMO.JavaScript) to build a report, or [convert an existing document](/convert/) with the .NET browser tools.
