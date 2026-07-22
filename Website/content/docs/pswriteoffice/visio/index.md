---
title: "Automate Visio Diagrams"
description: "Create, inspect, arrange, stencil, and export VSDX diagrams from PowerShell."
layout: docs
---

The Visio family exports 20 commands for VSDX authoring and inspection. Build pages from rectangles, ellipses, diamonds, containers, text boxes, connectors, and stencil shapes; arrange the result; save VSDX; and export SVG or PNG review artifacts.

## Create with primitives or stencils

Use `New-OfficeVisio` and `Add-OfficeVisioPage` to start a diagram. Primitive commands are useful for lightweight flowcharts and architecture maps. For recognizable product and infrastructure shapes, load a built-in or imported catalog and use `Add-OfficeVisioStencilShape`.

Stencil commands can discover catalogs, find shapes, import packages, and export preview galleries. That separates shape selection from document placement and makes diagram scripts repeatable across environments.

## Connect and arrange

`Add-OfficeVisioConnector` establishes graph relationships. Use layout and arrangement commands after content is complete, then apply targeted shape-layout updates only when the output needs manual refinement.

## Inspect and review

`Get-OfficeVisio` and `Get-OfficeVisioInfo` expose document and diagram evidence. `ConvertTo-OfficeVisioSvg` and `ConvertTo-OfficeVisioPng` provide portable review output without replacing the editable VSDX source.

See the [Visio examples](https://github.com/EvotecIT/PSWriteOffice/tree/main/Examples/Visio) and search the [command reference](/api/powershell/) for `OfficeVisio`.
