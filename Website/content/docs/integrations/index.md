---
title: "Reports across PowerShell modules"
description: "Inspect generated supplier reviews, incident briefs and email packages built with PowerShell modules. Download the files and adapt the complete scripts."
layout: docs
---

Start with the input you already have. These examples pass ordinary objects and files between PSParseHTML, PSEventViewer, PSWriteOffice and Mailozaurr.

| Input | Worked example | Output |
| --- | --- | --- |
| An HTML supplier table | [Supplier offer review](/docs/integrations/supplier-review/) | Word, PDF and numeric Excel |
| A bounded Windows event snapshot | [Event incident brief](/docs/integrations/event-brief/) | Word/PDF handoff and full Excel register |
| Existing report attachments | [Local email review package](/docs/integrations/report-email/) | EML with embedded attachments |

Each page shows the complete PowerShell script alongside its generated documents. Download an example ZIP, extract it to any folder, install the listed modules and run the script. The ZIP contains the script and any sample input it needs. You can also copy the code from the page or download the script separately.

The samples use synthetic data and run with published modules. Start with the offline examples, then substitute your HTML, event records or report attachments. If you have an older module installed, run `Update-Module` for that module and start a fresh PowerShell session before using the examples.

The [download manifest](/downloads/integrations/manifest.json) records the generating source revision, module versions and file hashes for reproducibility.
