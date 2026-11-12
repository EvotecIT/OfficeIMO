---
title: "Clean Files Before Sharing"
description: "Find and remove comments, author names, speaker notes, image metadata and AI labels from Word, Excel, PowerPoint, PDF and image files with PowerShell, and check what is left."
layout: docs
---

A file that leaves your team carries more than its visible text: author and company names, reviewer comments, tracked changes, speaker notes, custom properties, the location where a photo was taken, and Content Credentials or AI labels on its pictures. PSWriteOffice finds these items, writes a cleaned copy and tells you what happened to each one. The original file is never changed, and neither Office nor any other application needs to be installed.

| Command | What it does |
| --- | --- |
| `Get-OfficeHiddenData` | Lists what would leave with a file. Changes nothing. |
| `Remove-OfficeHiddenData` | Writes a cleaned copy and returns the outcome of every item. |
| `New-OfficeHiddenDataOptions` | Chooses the categories to remove, what to do with tracked changes, and which provenance records to keep. |
| `Get-OfficeDocumentProvenance` | Reports Content Credentials (C2PA) and AI-source labels in detail. |
| `Optimize-OfficePdf -Sanitize` | Removes attachments and scripts from a PDF, with the same options. |

`Get-OfficeHiddenData` and `Remove-OfficeHiddenData` work on Word, Excel, PowerPoint, Visio, OpenDocument, EPUB and PDF files, and on JPEG, PNG, WebP, GIF and TIFF images. They are the PowerShell form of `officeimo hidden-data inspect` and `officeimo hidden-data remove`, and they use the same engine.

The outputs below come from a folder with four files: a Word proposal with a reviewer comment, a custom property and a picture labelled as AI-generated; an Excel budget with a cell comment; a PowerPoint briefing with speaker notes; and a PDF summary with an attached cost model. The [clean-before-sharing recipe](https://github.com/EvotecIT/OfficeIMO/blob/master/PowerShell/PSWriteOffice/Examples/Workflows/Recipe-CleanBeforeSharing.ps1) builds similar files and runs the whole check.

## See what would leave with a file

```powershell
PS> Get-ChildItem -File | Get-OfficeHiddenData

Name          Format     Findings Changed Output
----          ------     -------- ------- ------
Briefing.pptx PowerPoint 4        False
Budget.xlsx   Excel      3        False
Proposal.docx Word       9        False
Summary.pdf   PDF        3        False
```

Each result has the file `Path`, its `Format`, the categories found in `Kinds`, and a `Findings` list:

```powershell
PS> (Get-OfficeHiddenData -Path .\Proposal.docx).Findings | Select-Object Kind, Description, Value

               Kind Description                                  Value
               ---- -----------                                  -----
           Comments Comment by Jane Reviewer                     Check the price before sending.
 PersonalIdentities Author name on 1 comment or revision marks   Jane Reviewer
 DocumentProperties Author                                       Jane Reviewer
   CustomProperties Custom property 'ClientCode'                 CONTOSO-7731
RevisionIdentifiers 1 revision save identifiers
RevisionIdentifiers 6 revision save identifiers
RevisionIdentifiers 17 revision save identifiers
RevisionIdentifiers 1 revision save identifiers
         Provenance Digital source type: TrainedAlgorithmicMedia http://cv.iptc.org/newscodes/digitalsourcety…

PS> (Get-OfficeHiddenData -Path .\Briefing.pptx).Findings | Select-Object Kind, Description, Value, Location

              Kind Description              Value                          Location
              ---- -----------              -----                          --------
      SpeakerNotes Speaker notes            Do not mention the merger yet. Slide 1
DocumentProperties Author                   Jane Reviewer                  /docProps/core.xml
DocumentProperties Last modified by         OfficeIMO                      /docProps/core.xml
         Thumbnail Document thumbnail image                                /docProps/thumbnail.jpeg
```

`Location` says where an item lives: a package part, a slide, a sheet cell or a comment. `Value` is a short preview, such as the author name or the first words of a comment. A file that `Get-OfficeHiddenData` can't read, for example a `.txt` file in the same folder, writes a non-terminating error and the rest of the folder is still checked.

## Write a cleaned copy

```powershell
PS> $report = Remove-OfficeHiddenData -Path .\Proposal.docx
PS> $report

Name          Format Findings Changed Output
----          ------ -------- ------- ------
Proposal.docx Word   8        True    Proposal.cleaned.docx

PS> $report.Findings | Select-Object Outcome, Kind, Description, Reason

 Outcome                Kind Description                                       Reason
 -------                ---- -----------                                       ------
 Removed            Comments Comment by Jane Reviewer                            None
 Removed  DocumentProperties Author                                              None
 Removed    CustomProperties Custom property 'ClientCode'                        None
 Removed RevisionIdentifiers 1 revision save identifiers                         None
 Removed RevisionIdentifiers 6 revision save identifiers                         None
 Removed RevisionIdentifiers 17 revision save identifiers                        None
 Removed RevisionIdentifiers 1 revision save identifiers                         None
Retained          Provenance Digital source type: TrainedAlgorithmicMedia NotSelected

PS> ($report.Findings | Where-Object Outcome -EQ Retained).Note
Kept because this category wasn't chosen for removal.
```

The copy is written as `<name>.cleaned<ext>` beside the input. Use `-OutputPath` to name it, or `-OutputDirectory` to write every copy into one folder under its original name. An existing output is never replaced silently: running the same command again fails with "The output already exists", and `-Force` replaces it. `-WhatIf` shows which files would be written without reading or writing anything.

Every finding has an `Outcome`: `Removed`, `Anonymized`, `Accepted`, `Rejected` or `Retained`. A retained finding has a `Reason` code and a `Note` sentence. Here the AI label on the picture stayed because provenance is not one of the default categories. Inspecting the copy shows what is left:

```powershell
PS> (Get-OfficeHiddenData -Path .\Proposal.cleaned.docx).Findings | Select-Object Outcome, Kind, Description

Outcome       Kind Description
-------       ---- -----------
  Found Provenance Digital source type: TrainedAlgorithmicMedia
```

## Choose what to remove

With no options, `Remove-OfficeHiddenData` removes the categories that never change what readers see: document and custom properties, comments, author names, speaker notes, the thumbnail, custom XML, image metadata such as GPS locations, printer settings, external references, Word document variables and revision identifiers, and scripts and other active content. These stay until you select them:

- **Tracked changes.** Accepting or rejecting them changes the text. Use `-TrackedChanges AcceptAll` or `RejectAll`.
- **Hidden text, rows, columns, sheets and slides.** Readers would see different content, and deleting a hidden sheet can break formulas. Add `HiddenContent`.
- **Attachments and PDF bookmarks.** The reader may need them. Add `Attachments` or `Bookmarks`.
- **Content Credentials and AI labels.** They tell the reader how content was made. Add `Provenance`.

`New-OfficeHiddenDataOptions` builds the choice. `-Kind` takes `Default`, `All` or category names, and combines them:

```powershell
PS> $options = New-OfficeHiddenDataOptions -Kind Default, Provenance
PS> $options | Format-List Kinds, TrackedChanges, AnonymousAuthor

Kinds           : Provenance, Default
TrackedChanges  : Keep
AnonymousAuthor : Author
```

`-KeepProvenance` keeps some provenance records when `Provenance` is selected: `C2pa` keeps Content Credentials, `External` keeps links to them, and `AiSource` keeps AI-generated labels. `New-OfficeHiddenDataOptions -Kind Default, Provenance -KeepProvenance C2pa, External` removes only the AI labels. `-AnonymousAuthor` sets the name that replaces authors on comments and revisions you keep; the default is `Author`.

`Remove-OfficeHiddenData` refuses to save a change that would break an Office package signature. Add `-RemoveInvalidatedSignatures` to remove those signatures and sign the cleaned copy again afterwards. Signed PDFs are always refused.

## Content Credentials and AI labels

`Get-OfficeDocumentProvenance` reports provenance in more detail than the hidden-data list, including signatures and, with a `c2patool` you supply, cryptographic verification:

```powershell
PS> (Get-OfficeDocumentProvenance -Path .\Proposal.docx).Structural | Select-Object Format, HasC2paManifest, HasGenerativeAiDeclaration

    Format HasC2paManifest HasGenerativeAiDeclaration
    ------ --------------- --------------------------
ZipPackage           False                       True
```

Selecting `Provenance` removes the label from the picture inside the proposal:

```powershell
PS> $report = Remove-OfficeHiddenData -Path .\Proposal.docx -Options $options -Force
PS> $report.Findings | Where-Object Kind -EQ Provenance | Select-Object Outcome, Description, Part

Outcome Description                                  Part
------- -----------                                  ----
Removed Digital source type: TrainedAlgorithmicMedia /media/image.png

PS> (Get-OfficeDocumentProvenance -Path .\Proposal.cleaned.docx).Structural.HasGenerativeAiDeclaration
False
```

Content Credentials that cover the whole file, such as those in a PDF or in an Office package's own credential part, are always removed, because any change to the file makes them invalid. Remove provenance only when you mean to: a label that says a picture was generated is often something the reader should see.

## Clean a folder and check the result

```powershell
PS> Get-ChildItem -File | Where-Object Name -NotLike '*.cleaned.*' |
        Remove-OfficeHiddenData -OutputDirectory ..\Share -Options $options

Name          Format     Findings Changed Output
----          ------     -------- ------- ------
Briefing.pptx PowerPoint 4        True    Briefing.pptx
Budget.xlsx   Excel      2        True    Budget.xlsx
Proposal.docx Word       8        True    Proposal.docx
Summary.pdf   PDF        3        True    Summary.pdf

PS> Get-ChildItem ..\Share | Get-OfficeHiddenData

Name          Format     Findings Changed Output
----          ------     -------- ------- ------
Briefing.pptx PowerPoint 0        False
Budget.xlsx   Excel      0        False
Proposal.docx Word       0        False
Summary.pdf   PDF        1        False
```

The finding left in `Summary.pdf` is the attached cost model, because attachments are opt-in. A scheduled job can stop on anything that remains:

```powershell
$left = Get-ChildItem .\Share -File | Get-OfficeHiddenData | Where-Object { $_.Findings.Count -gt 0 }
if ($left) { throw "Hidden data remains in: $($left.Path -join ', ')" }
```

## PDF files

`Remove-OfficeHiddenData` cleans PDFs with the same categories as Office files. For a PDF the defaults remove the author-written document information and XMP metadata, comments, scripts and unsafe actions, and keep attachments, bookmarks and layers:

```powershell
PS> (Get-OfficeHiddenData -Path .\Summary.pdf).Findings | Select-Object Kind, Description, Value

              Kind Description                         Value
              ---- -----------                         -----
DocumentProperties Document information Title          Project summary
DocumentProperties Document information Author         Jane Reviewer
       Attachments Attached file (Names/EmbeddedFiles) cost-model.txt
```

`Optimize-OfficePdf -Sanitize` is the PDF route for removing what can run or travel with the file. Its default is attachments and active content, and it leaves the document information alone:

```powershell
PS> $result = Optimize-OfficePdf -Path .\Summary.pdf -OutputPath .\Summary-sanitized.pdf -Sanitize -PassThru
PS> $result.Report.Findings | Select-Object Outcome, Kind, Description

 Outcome               Kind Description
 -------               ---- -----------
Retained DocumentProperties Document information Title
Retained DocumentProperties Document information Author
 Removed        Attachments Attached file (Names/EmbeddedFiles)
```

Both accept the options from `New-OfficeHiddenDataOptions`; `Optimize-OfficePdf` takes them through `-HiddenDataOptions`. Use [Redact, sanitize, and optimize](/docs/pswriteoffice/pdf-redact-sanitize/) when visible text has to go too: removing hidden data doesn't redact page content.

## Related

- [Prepare a document for sharing](/docs/before-sharing/): the same checks from .NET, the browser and the command line, with accessibility.
- [Remove hidden data and metadata](/docs/before-sharing/hidden-data/): every category and what each format reports.
- [Content Credentials and AI labels](/docs/before-sharing/provenance/)
- [Before sharing from the CLI](/docs/cli/before-sharing/): `officeimo hidden-data` and provenance checks in CI.
- Command reference: [Get-OfficeHiddenData](/api/powershell/get-officehiddendata/), [Remove-OfficeHiddenData](/api/powershell/remove-officehiddendata/), [New-OfficeHiddenDataOptions](/api/powershell/new-officehiddendataoptions/) and [Get-OfficeDocumentProvenance](/api/powershell/get-officedocumentprovenance/).
