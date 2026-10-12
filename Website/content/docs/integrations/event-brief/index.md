---
title: "Windows event incident brief"
description: "Turn a bounded Windows event snapshot into a Word/PDF investigation handoff and an Excel register retaining source computers, logs and full messages."
layout: docs
---

Give an investigator a short brief and the underlying event records. The default example uses four synthetic records; live collection is a separate option.

![Generated incident brief with provider counts and investigation handoff](/downloads/integrations/events/preview.png)

[Word brief](/downloads/integrations/events/EventIncidentBrief.docx) · [PDF brief](/downloads/integrations/events/EventIncidentBrief.pdf) · [Excel register](/downloads/integrations/events/EventRegister.xlsx)

## Start with the offline example

Download the [event brief example ZIP](/downloads/integrations/event-brief.zip) to your working folder. It contains one standalone script. Open PowerShell in that folder and run:

```powershell
Install-Module PSWriteOffice -Scope CurrentUser
Expand-Archive -LiteralPath ./event-brief.zip -DestinationPath ./EventBriefExample
./EventBriefExample/Recipe-PSEventViewer-OfficeReport.ps1 -OutputDirectory ./Output/EventBrief
```

Open `Output/EventBrief` to find the DOCX, PDF and XLSX. This offline example requires PSWriteOffice 3.0.8 or later. If an older copy is installed, run `Update-Module PSWriteOffice` and start a fresh PowerShell session. You can also save the [standalone script](/downloads/integrations/source/Recipe-PSEventViewer-OfficeReport.ps1) anywhere and run it from that location.

If Windows blocks the downloaded script, read it first, then unblock that file with `Unblock-File -LiteralPath ./EventBriefExample/Recipe-PSEventViewer-OfficeReport.ps1`. Keep your normal execution policy.

## Use an existing snapshot or collect once

```powershell
Install-Module PSEventViewer -Scope CurrentUser
$events = @(Get-EVXEvent -LogName Application -Level 1,2,3 -TimePeriod Last24Hours -MaxEvents 100 -ReadMode Message)
./EventBriefExample/Recipe-PSEventViewer-OfficeReport.ps1 -Event $events -MaxEvents 100 -OutputDirectory ./Output/LiveBrief
```

Alternatively, use `-CollectLive -LogName Application -MaxEvents 100`. Live collection requires Windows, PSEventViewer 4.0.1 or later and access to the channel. PSEventViewer is maintained in [EventViewerX](https://github.com/EvotecIT/EventViewerX).

![Native Excel preview of retained timestamps, providers, source computers, logs and full event messages](/downloads/integrations/events/workbook.png)

The brief shows the retained count and top eight providers with critical, error and warning counts. Excel keeps every retained record up to `MaxEvents`, including source computer and log where supplied, the full message and typed timestamp. The XLSX timestamp format includes seconds and milliseconds so events within the same minute remain distinguishable. Raw and typed `Get-EVXEvent` results are accepted; typed results use their original `SourceEvent` snapshot. `-Event @()` produces an explicit empty report. Counts describe this bounded snapshot; they do not establish root cause, service impact or the total size of a log. Review messages before sharing a live register.

Use EventViewerX's `Show-EVXEvent` for its native reports. Choose this integration when the handoff also needs an editable Word brief and PDF, then prepare a [local email package](/docs/integrations/report-email/).

## Complete PowerShell script

The code below is the same script included in the ZIP and standalone download. Change the sample records or pass your own snapshot with `-Event`.

{{< integration-example id="event-brief" >}}
