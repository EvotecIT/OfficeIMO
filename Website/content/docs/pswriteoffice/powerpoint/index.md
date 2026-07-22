---
title: "Automate PowerPoint Presentations"
description: "Compose, inspect, update, theme, import, and render repeatable presentation decks."
layout: docs
---

The PowerPoint family exports 57 commands for slide creation and editing, sections, shapes, images, text, charts, tables, notes, themes, layouts, transitions, import, inspection, designer decks, and semantic deck plans.

## Choose direct authoring or a deck plan

Direct authoring with `New-OfficePowerPoint` and `Add-OfficePowerPointSlide` is appropriate when the script owns exact slide composition. Add text boxes, shapes, tables, images, bullets, charts, notes, sections, and transitions inside the presentation context.

Use `New-OfficePowerPointDeckPlan` and the `Add-OfficePowerPointPlan*` commands when the content is semantic and the designer should choose layout. Plan sections, processes, capabilities, case studies, coverage views, card grids, and logo walls can be described before a design alternative is selected.

## Inspect and update an existing deck

Inspection commands expose slides, sections, shapes, placeholders, layouts, notes, themes, and slide summaries. Bounded setters update titles, shape text, table cells, slide size and layout, placeholder bounds and text styles, notes, transitions, backgrounds, and theme identity.

Copy or import slides when the workflow assembles a deck from approved sources. Keep theme and layout changes separate from content changes so validation can distinguish a brand update from a data update.

## Review output

Use `ConvertTo-OfficePowerPointHtml` for a browser-review surface and `Export-OfficePowerPointImage` for visual artifacts. For exact parameters and supported chart/layout values, search the [command reference](/api/powershell/) for `OfficePowerPoint`. The [PowerPoint examples](https://github.com/EvotecIT/PSWriteOffice/tree/main/Examples/PowerPoint) demonstrate normal script shapes.
