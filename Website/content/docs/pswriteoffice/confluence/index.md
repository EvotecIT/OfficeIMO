---
title: "Publish to Confluence Cloud"
description: "Plan and publish Confluence pages, preserve managed sections, and transfer attachments. Includes PowerShell examples, safety notes, and cmdlet links."
layout: docs
---

The Confluence family exports seven commands over OfficeIMO.Adf and OfficeIMO.Confluence. Scripts can read pages, plan or publish create/update requests, remove pages, preserve generated sections inside owner-authored content, and list, upload, or download attachments.

## Review a page request before sending it

`Publish-OfficeConfluencePage -PlanOnly` builds the same serializable request plan used by a live write without requiring a session or contacting Atlassian. Use it to review the method, URI, converted body, and payload before supplying credentials.

```powershell
$plan = Publish-OfficeConfluencePage `
    -SpaceId 42 `
    -Title 'Daily status' `
    -Content "# Ready`n`nGenerated from the current report." `
    -PlanOnly

$plan.Payload
```

Create a live session with an Atlassian email/API-token credential, or use an OAuth access token with the tenant cloud ID. Write and delete commands support `ShouldProcess`; destructive purge requests can also be produced as plans first.

## Preserve content people still own

`Set-OfficeConfluenceManagedSection` replaces one marker-delimited section and returns before/after hashes without contacting Confluence. The rest of the storage body stays untouched, so a generated report can be refreshed without taking ownership of the whole page.

## Transfer attachments without buffering files

`Get-OfficeConfluenceAttachment` lists page attachments or streams a download to disk. `Send-OfficeConfluenceAttachment` streams file uploads through the shared OfficeIMO client. Byte-array download output remains available when an output file is not requested.

See the [Confluence reporting example](https://github.com/EvotecIT/PSWriteOffice/tree/main/Examples/Confluence) and search the [command reference](/api/powershell/) for `OfficeConfluence`.
