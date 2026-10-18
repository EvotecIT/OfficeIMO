---
title: "Automate PowerPoint Presentations"
description: "Compose, inspect, update, theme, import, and render repeatable presentation decks. Includes PowerShell examples, validation notes, and cmdlet links."
layout: docs
---

The PowerPoint family exports 59 commands for slide creation and editing, sections, shapes, images, text, charts, tables, notes, themes, layouts, transitions, import, inspection, designer decks, and semantic deck plans.

## Use a presentation object in normal scripts

Create with `-NoSave`, add slides through an explicit presentation target, then save and close once. This keeps loops and conditional slide generation ordinary PowerShell:

```powershell
$presentation = New-OfficePowerPoint -Path '.\Briefing.pptx' -NoSave
$slide = Add-OfficePowerPointSlide -Presentation $presentation -LayoutType Text -PassThru
Set-OfficePowerPointSlideTitle -Slide $slide -Title 'Actions'
Add-OfficePowerPointTextBox -Slide $slide -Text 'Confirm the production date.' -X 90 -Y 170 -Width 700 -Height 60
$presentation | Close-OfficePowerPoint -Save
```

## Choose direct authoring or a deck plan

Direct authoring with `New-OfficePowerPoint` and `Add-OfficePowerPointSlide` is appropriate when the script owns exact slide composition. Add text boxes, shapes, tables, images, bullets, charts, notes, sections, and transitions inside the presentation context.

Use `New-OfficePowerPointDeckPlan` and the `Add-OfficePowerPointPlan*` commands when the content is semantic and the designer should choose layout. Plan sections, processes, capabilities, case studies, coverage views, card grids, and logo walls can be described before a design alternative is selected.

## Copy complete DSL recipes

- [Quarterly business review](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/PowerPoint/Recipe-PowerPoint-QuarterlyReview.ps1): title slide, performance chart, priorities table, bullets, and speaker notes.
- [Training workshop](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/PowerPoint/Recipe-PowerPoint-TrainingWorkshop.ps1): learning objectives, agenda, call-to-action layout, and presenter notes.
- [Service brief](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Showcase/Showcase-PowerPoint-ServiceBrief.ps1): semantic designer slides combined with direct chart and table slides, sections, transitions, and inspection.

## Other PowerPoint workflows

- [Object composition](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/PowerPoint/Recipe-PowerPoint-ObjectComposition.ps1): build a deck through explicit presentation and slide objects.
- [Sections and speaker notes](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/PowerPoint/Recipe-PowerPoint-SectionsAndNotes.ps1): organize a briefing and keep presenter context with each slide.
- [Copy and remove slides](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/PowerPoint/Recipe-PowerPoint-CopyAndRemoveSlides.ps1): assemble a delivery deck from reusable material.

The [DSL cookbook](/docs/pswriteoffice/dsl-cookbook/) explains when to use direct placement and when to use a semantic deck plan.

## Task guides

- [Read and inspect decks](/docs/pswriteoffice/powerpoint-read-inspect/)
- [Update existing presentations](/docs/pswriteoffice/powerpoint-update-existing/)
- [Reuse and combine slides](/docs/pswriteoffice/powerpoint-reuse-slides/)
- [Export and review presentations](/docs/pswriteoffice/powerpoint-export-review/)

## Inspect and update an existing deck

Inspection commands expose slides, sections, shapes, placeholders, layouts, notes, themes, and slide summaries. Bounded setters update titles, shape text, table cells, slide size and layout, placeholder bounds and text styles, notes, transitions, backgrounds, and theme identity.

Copy or import slides when the workflow assembles a deck from approved sources. Keep theme and layout changes separate from content changes so validation can distinguish a brand update from a data update.

## Review output

Use `ConvertTo-OfficePowerPointHtml` for a browser-review surface and `Export-OfficePowerPointImage` for visual artifacts. For exact parameters and supported chart/layout values, search the [command reference](/api/powershell/) for `OfficePowerPoint`. The [PowerPoint examples](https://github.com/EvotecIT/PSWriteOffice/tree/main/Examples/PowerPoint) demonstrate normal script shapes.
