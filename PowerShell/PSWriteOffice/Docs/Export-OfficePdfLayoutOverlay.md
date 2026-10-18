---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Export-OfficePdfLayoutOverlay
## SYNOPSIS
Exports PDF word, line, region, and reading-order diagnostics as PNG or SVG.

## SYNTAX
### __AllParameterSets
```powershell
Export-OfficePdfLayoutOverlay [-Path] <string> [-OutputPath] <string> [-Page <int>] [-Format <OfficeImageExportFormat>] [-Scale <double>] [-Options <PdfLayoutDebugOverlayOptions>] [-LayoutOptions <PdfTextLayoutOptions>] [-ReadOptions <PdfLoadOptions>] [-Password <string>] [-IgnorePermissionRestrictions] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Exports PDF word, line, region, and reading-order diagnostics as PNG or SVG.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $result = Export-OfficePdfLayoutOverlay -Path .\Report.pdf -OutputPath .\Report-layout.svg -Page 1 -Format Svg -PassThru
```


## PARAMETERS

### -Format
Output image format. Layout overlays support only PNG and SVG.

```yaml
Type: OfficeImageExportFormat
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Png, Svg

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IgnorePermissionRestrictions
After successful password authentication, explicitly ignore owner-imposed extraction restrictions.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LayoutOptions
Optional text layout settings.

```yaml
Type: PdfTextLayoutOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Options
Optional overlay elements, colors, and limits.

```yaml
Type: PdfLayoutDebugOverlayOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputPath
Destination PNG or SVG path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Page
One-based page number.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the structured image export result.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Password used to authenticate an encrypted PDF.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Source PDF path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ReadOptions
Optional bounded PDF parsing settings.

```yaml
Type: PdfLoadOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Scale
Output scale.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `OfficeIMO.Drawing.OfficeImageExportResult`

## RELATED LINKS

- None
