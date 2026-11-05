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
    Import-OfficeVisioStencilCatalog -BuiltIn Flowchart -Name Flow -Default

    VisioStencil -Catalog Flow -Stencil process -Key intake -Text 'Intake' -X 1.5in -Y 4in
    VisioStencil -Catalog Flow -Stencil decision -Key review -Text 'Review?' -X 4in -Y 4in
    VisioStencil -Catalog Flow -Stencil data -Key archive -Text 'Archive' -X 6.5in -Y 4in
    VisioConnector -From intake -To review -Kind RightAngle -EndArrow Triangle -Label 'submit'
    VisioConnector -From review -To archive -Kind RightAngle -EndArrow Triangle -Label 'store'
}

(Get-OfficeDocumentInfo -Path $visioPath -Detailed).Details.ToText()
Export-OfficeDocumentImage -Path $visioPath -Page 1 -OutputPath $svgPath | Out-Null
Export-OfficeDocumentImage -Path $visioPath -Page 1 -OutputPath $pngPath | Out-Null
```

## What this demonstrates

- creating editable Visio diagrams without automating desktop Visio
- using built-in stencils and connectors
- exporting SVG and PNG previews from the same `.vsdx`

## Source

- [Example-Visio-StencilFlow.ps1](https://github.com/EvotecIT/OfficeIMO/blob/master/PowerShell/PSWriteOffice/Examples/Visio/Example-Visio-StencilFlow.ps1)
