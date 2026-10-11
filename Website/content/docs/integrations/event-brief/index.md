---
title: "Windows event incident brief"
description: "Turn a bounded event snapshot into a Word/PDF investigation handoff and an Excel register with full messages."
layout: docs
---

Give an investigator a short brief and the underlying event records. The default example uses four synthetic records; live collection is a separate option.

![Generated incident brief with provider counts and investigation handoff](/downloads/integrations/events/preview.png)

[Word brief](/downloads/integrations/events/EventIncidentBrief.docx) · [PDF brief](/downloads/integrations/events/EventIncidentBrief.pdf) · [Excel register](/downloads/integrations/events/EventRegister.xlsx)

## Start with the offline example

```powershell
Install-Module PSWriteOffice -Scope CurrentUser
./Examples/Integrations/Recipe-PSEventViewer-OfficeReport.ps1 -OutputDirectory ./Output/EventBrief
```

Run from a PSWriteOffice checkout, or download the [complete script](/downloads/integrations/source/Recipe-PSEventViewer-OfficeReport.ps1). Its [source revision](https://github.com/EvotecIT/PSWriteOffice/blob/db029a8ee3fcb96227777894daaf8bbd7f94dab1/Examples/Integrations/Recipe-PSEventViewer-OfficeReport.ps1) matches these downloads.

## Use an existing snapshot or collect once

```powershell
Install-Module PSEventViewer -Scope CurrentUser
$events = @(Get-EVXEvent -LogName Application -Level 1,2,3 -TimePeriod Last24Hours -MaxEvents 100 -ReadMode Message)
./Examples/Integrations/Recipe-PSEventViewer-OfficeReport.ps1 -Event $events -MaxEvents 100 -OutputDirectory ./Output/LiveBrief
```

Alternatively, use `-CollectLive -LogName Application -MaxEvents 100`. Live collection requires Windows, PSEventViewer 4.0.1 or later and access to the channel. PSEventViewer is maintained in [EventViewerX](https://github.com/EvotecIT/EventViewerX).

![Native Excel preview of retained timestamps, providers and full event messages](/downloads/integrations/events/workbook.png)

The brief shows the retained count and top eight providers with critical, error and warning counts. Excel keeps every retained record up to `MaxEvents`, including the full message and typed timestamp. `-Event @()` produces an explicit empty report. Counts describe this bounded snapshot; they do not establish root cause, service impact or the total size of a log. Review messages before sharing a live register.

Use EventViewerX's `Show-EVXEvent` for its native reports. Choose this integration when the handoff also needs an editable Word brief and PDF, then prepare a [local email package](../report-email/).
