---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Compare-OfficePdfVisual
## SYNOPSIS
Compares rendered PDF pages and returns pixel-level review artifacts.

## SYNTAX
### __AllParameterSets
```powershell
Compare-OfficePdfVisual [-ReferencePath] <string> [-DifferencePath] <string> [-PageRange <string>] [-Options <PdfVisualComparisonOptions>] [-ReferenceReadOptions <PdfLoadOptions>] [-DifferenceReadOptions <PdfLoadOptions>] [-ReferencePassword <string>] [-IgnoreReferencePermissionRestrictions] [-DifferencePassword <string>] [-IgnoreDifferencePermissionRestrictions] [<CommonParameters>]
```

## DESCRIPTION
Compares rendered PDF pages and returns pixel-level review artifacts.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $options = New-OfficePdfVisualComparisonOptions -AllowedDifferenceRatio 0.001 -ChannelTolerance 2
Compare-OfficePdfVisual -ReferencePath .\Expected.pdf -DifferencePath .\Actual.pdf -PageRange '1-3' -Options $options
```

Returns per-page difference ratios, images, and diagnostics.

## PARAMETERS

### -DifferencePassword
Password used to authenticate the actual PDF.

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

### -DifferencePath
Actual PDF path.

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

### -DifferenceReadOptions
Optional bounded read settings for the actual document.

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

### -IgnoreDifferencePermissionRestrictions
After authentication, explicitly ignore restrictions on the actual PDF.

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

### -IgnoreReferencePermissionRestrictions
After authentication, explicitly ignore restrictions on the expected PDF.

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

### -Options
Optional render, tolerance, alignment, background, and ignored regions.

```yaml
Type: PdfVisualComparisonOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PageRange
Optional one-based ranges such as 1-3,5.

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

### -ReferencePassword
Password used to authenticate the expected PDF.

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

### -ReferencePath
Expected PDF path.

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

### -ReferenceReadOptions
Optional bounded read settings for the expected document.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `OfficeIMO.Pdf.PdfVisualComparisonReport`

## RELATED LINKS

- None
