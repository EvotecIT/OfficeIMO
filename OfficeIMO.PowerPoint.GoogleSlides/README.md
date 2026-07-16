# OfficeIMO.PowerPoint.GoogleSlides

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.PowerPoint.GoogleSlides)](https://www.nuget.org/packages/OfficeIMO.PowerPoint.GoogleSlides)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.PowerPoint.GoogleSlides?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.PowerPoint.GoogleSlides)

`OfficeIMO.PowerPoint.GoogleSlides` provides bidirectional translation between `OfficeIMO.PowerPoint` presentations and Google Slides. It keeps the editable core native and makes fidelity fallbacks visible through a translation report.

## Install

```powershell
dotnet add package OfficeIMO.PowerPoint.GoogleSlides
```

## Create a Google presentation

```csharp
using OfficeIMO.GoogleWorkspace;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.GoogleSlides;

using PowerPointPresentation deck = PowerPointPresentation.Create();
PowerPointSlide slide = deck.AddSlide();
slide.AddTextBoxPoints("Quarterly review", 30, 40, 500, 80);
slide.Notes.Text = "Mention the year-over-year comparison.";

var session = new GoogleWorkspaceSession(
    new StaticAccessTokenCredentialSource("<google-access-token>"),
    new GoogleWorkspaceSessionOptions { DefaultFolderId = "reports-folder-id" });

var options = new GoogleSlidesSaveOptions { Title = "Quarterly review" };
GoogleSlidesTranslationPlan plan = deck.BuildGoogleSlidesPlan(options);
GooglePresentationReference result = await deck.ExportToGoogleSlidesAsync(session, options);

Console.WriteLine(result.WebViewLink);
foreach (TranslationNotice notice in result.Report.Notices) {
    Console.WriteLine($"{notice.Severity}: {notice.Message}");
}
```

Text boxes, core text styling, hyperlinks, tables, pictures, basic shapes, slide size, solid backgrounds, and speaker notes are authored as editable Google Slides content. By default a slide containing charts, SmartArt, media, OLE, or another unsupported complex object is rendered to one PNG so the visual result remains coherent. Select `PreferNativeAndReport` only when skipping unsupported objects is acceptable.

## Replace safely

Read or import the Google presentation first and pass the observed revision when replacing it:

```csharp
GoogleSlidesImportResult imported = await session.ImportGoogleSlidesAsync(
    "presentation-id",
    new GoogleSlidesImportOptions { Mode = GoogleSlidesImportMode.Native });

using (imported.Presentation) {
    var replace = new GoogleSlidesSaveOptions {
        Location = new GoogleDriveFileLocation { ExistingFileId = imported.Source.FileId },
        Replace = new GoogleSlidesReplaceOptions {
            ExpectedRevisionId = imported.Source.RevisionId
        }
    };

    await deck.ExportToGoogleSlidesAsync(session, replace);
}
```

An observed revision is required for an existing presentation. A stale revision fails before mutation. `OverwriteLatest` is an explicit last-writer-wins escape hatch and intentionally omits the Slides write guard.

Set `TemplatePresentationId` to copy a template before authoring. Template copy, target folder placement, shared-drive options, and metadata are handled through `OfficeIMO.GoogleWorkspace.Drive`.

## Import modes

- `DriveExport` exports Google Slides to PPTX and loads it through `OfficeIMO.PowerPoint`. This is the default and broadest-fidelity import.
- `Native` reads Slides API objects directly. It imports slide size, text boxes and core styles, hyperlinks, tables, images, geometry, and speaker-note text, and returns the Slides revision needed for safe replacement.

The Slides API fetches created images from a URL. Export therefore creates short-lived public Drive objects and removes their permissions and files after the batch, including failure paths. Do not grant permanent public access to source assets.

## Authentication and scopes

The core package has no dependency on the Google client SDK. Supply any `IGoogleWorkspaceCredentialSource`, or install `OfficeIMO.GoogleWorkspace.Auth.GoogleApis` when you want adapters for Google service-account or installed-application credentials.

Authoring uses `GoogleWorkspaceScopeCatalog.SlidesAuthoring` (`drive.file` and `presentations`). Native read uses `presentations.readonly`; Drive-export import also needs Drive read access appropriate to the file.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`, plus `net472` on Windows.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)

See `GoogleSlidesFeatureSupportCatalog.Features` for the code-owned support matrix and the [complete OfficeIMO package map](../README.md) for related formats.
