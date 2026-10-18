---
title: "Create a Visio diagram"
description: "Use PSWriteOffice to create a Visio diagram with stencils, connectors, and exported previews."
layout: docs
---

This pattern is useful when a script needs to generate an editable `.vsdx` diagram and preview images for review.

It is adapted from `Examples/Visio/Example-Visio-StencilFlow.ps1`.

## Example

```powershell
Import-Module PSWriteOffice

$outputDirectory = Join-Path $PSScriptRoot 'Output'
$null = New-Item -ItemType Directory -Path $outputDirectory -Force
$visioPath = Join-Path $outputDirectory 'OnboardingFlow.vsdx'
$svgPath = Join-Path $outputDirectory 'OnboardingFlow.svg'
$pngPath = Join-Path $outputDirectory 'OnboardingFlow.png'

New-OfficeVisio -Path $visioPath -Title 'Customer onboarding flow' -UseMastersByDefault -RequestRecalcOnOpen {
    Import-OfficeVisioStencil -BuiltIn Flowchart -Name Flow -Default

    VisioStencil -Catalog Flow -Stencil process -Key intake -Text 'Intake' -X 1.5 -Y 4
    VisioStencil -Catalog Flow -Stencil decision -Key review -Text 'Review?' -X 4 -Y 4
    VisioStencil -Catalog Flow -Stencil data -Key archive -Text 'Archive' -X 6.5 -Y 4
    VisioConnector -From intake -To review -Kind RightAngle -EndArrow Triangle -Label 'submit'
    VisioConnector -From review -To archive -Kind RightAngle -EndArrow Triangle -Label 'store'
}

Get-OfficeVisioInfo -Path $visioPath -AsText
ConvertTo-OfficeVisioSvg -Path $visioPath -OutputPath $svgPath
ConvertTo-OfficeVisioPng -Path $visioPath -OutputPath $pngPath
```

## What this demonstrates

- creating editable Visio diagrams without automating desktop Visio
- using built-in stencils and connectors
- exporting SVG and PNG previews from the same `.vsdx`

## Source

- [Example-Visio-StencilFlow.ps1](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Visio/Example-Visio-StencilFlow.ps1)
