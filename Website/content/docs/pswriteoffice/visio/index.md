---
title: "Automate Visio Diagrams"
description: "Create, inspect, arrange, stencil, and export VSDX diagrams from PowerShell. Includes PowerShell examples, validation notes, and cmdlet links."
layout: docs
---

The Visio family exports 21 commands for VSDX authoring and inspection. Build pages from rectangles, ellipses, diamonds, containers, text boxes, connectors, and stencil shapes; arrange the result; save VSDX; and export SVG or PNG review artifacts.

## Create with primitives or stencils

Use `New-OfficeVisio` and `Add-OfficeVisioPage` to start a diagram. Primitive commands are useful for lightweight flowcharts and architecture maps. For recognizable product and infrastructure shapes, load a built-in or imported catalog and use `Add-OfficeVisioStencilShape`.

Stencil commands can discover catalogs, find shapes, import packages, and export preview galleries. That separates shape selection from document placement and makes diagram scripts repeatable across environments.

## Connect and arrange

`Add-OfficeVisioConnector` establishes graph relationships. Use layout and arrangement commands after content is complete, then apply targeted shape-layout updates only when the output needs manual refinement.

## Inspect and review

`Get-OfficeVisio` and `Get-OfficeVisioInfo` expose document and diagram evidence. `Export-OfficeDocumentImage` writes portable SVG or PNG review output from the editable VSDX source. Use `-Page 1 -OutputPath ./diagram.svg`, or `-OutputDirectory ./pages -Format Png` for every page. `ConvertTo-OfficeDocumentImage` returns the same native image results without writing files.

See the [Visio examples](https://github.com/EvotecIT/OfficeIMO/tree/640a25d8fa79b69a90ed309104f7134cc190cef9/PowerShell/PSWriteOffice/Examples/Visio) and search the [command reference](/api/powershell/) for `OfficeVisio`.
