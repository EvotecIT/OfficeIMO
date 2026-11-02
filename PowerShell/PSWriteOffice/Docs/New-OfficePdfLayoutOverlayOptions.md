---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficePdfLayoutOverlayOptions
## SYNOPSIS
Creates native PDF layout-overlay layers, colors and resource limits.

## SYNTAX
### Default (Default)
```powershell
New-OfficePdfLayoutOverlayOptions [-NoWords] [-NoLines] [-NoRegions] [-NoReadingOrder] [-WordColor <OfficeColor>] [-LineColor <OfficeColor>] [-RegionColor <OfficeColor>] [-MaxElements <Int32>] [-MaximumRasterPixelCount <Int64>] [<CommonParameters>]
```

## DESCRIPTION
Creates native PDF layout-overlay layers, colors and resource limits.

## EXAMPLES

### EXAMPLE 1
```powershell
$options = New-OfficePdfLayoutOverlayOptions -NoWords -NoLines -RegionColor '#FF00FF'
Export-OfficePdfLayoutOverlay -Path ./Report.pdf -OutputPath ./Order.svg -Options $options
```


## PARAMETERS

### -LineColor
Line-box color or theme color role.

```yaml
Type: OfficeColor
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxElements
Maximum boxes and labels in one overlay.

```yaml
Type: Int32
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumRasterPixelCount
Maximum allocated pixels in a PNG overlay.

```yaml
Type: Int64
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoLines
Hide line boxes.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoReadingOrder
Hide reading-order labels.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoRegions
Hide paragraph region boxes.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoWords
Hide word boxes.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RegionColor
Paragraph-region color or theme color role.

```yaml
Type: OfficeColor
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WordColor
Word-box color or theme color role.

```yaml
Type: OfficeColor
Parameter Sets: Default
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

- `OfficeIMO.Pdf.PdfLayoutDebugOverlayOptions`

## RELATED LINKS

- None
