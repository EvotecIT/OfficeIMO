---
title: "Reports across PowerShell modules"
description: "Inspect generated supplier reviews, incident briefs and email packages built with PowerShell modules. Download the files and adapt the complete scripts."
layout: docs
---

Start with the input you already have. These examples pass ordinary objects and files between PSParseHTML, PSEventViewer, PSWriteOffice and Mailozaurr.

| Input | Worked example | Output |
| --- | --- | --- |
| An HTML supplier table | [Supplier offer review](./supplier-review/) | Word, PDF and numeric Excel |
| A bounded Windows event snapshot | [Event incident brief](./event-brief/) | Word/PDF handoff and full Excel register |
| Existing report attachments | [Local email review package](./report-email/) | EML with embedded attachments |

Each page includes real generated files and the complete source that produced them. The downloads contain synthetic data. The recipes run with published modules and add no required dependencies to PSWriteOffice.

Use the [integration guide](https://github.com/EvotecIT/PSWriteOffice/tree/76b1994c7a3a2fe23d5e57b180c94aa1c8e43d58/Examples/Integrations) for module requirements and input contracts. The [download manifest](/downloads/integrations/manifest.json) records the source revision, module versions and file hashes.
