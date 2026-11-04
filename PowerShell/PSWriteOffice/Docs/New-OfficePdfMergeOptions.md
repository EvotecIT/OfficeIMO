---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficePdfMergeOptions
## SYNOPSIS
Creates native PDF annotation-flattening and page-normalization settings.

## SYNTAX
### Default (Default)
```powershell
New-OfficePdfMergeOptions [-FlattenVisualAnnotations] [-PageSize <OfficePageSize>] [-Width <OfficeLength>] [-Height <OfficeLength>] [-Orientation <OfficePageOrientation>] [-ResizeMode <PdfPageResizeMode>] [-ResizeMargin <OfficeLength>] [<CommonParameters>]
```

## DESCRIPTION
Creates native PDF annotation-flattening and page-normalization settings.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficePdfMergeOptions -FlattenVisualAnnotations
```


## PARAMETERS

### -FlattenVisualAnnotations
Paint supported visual annotations into page content before merging.

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

### -Height
Custom target page height.

```yaml
Type: OfficeLength
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Orientation
Target page orientation.

```yaml
Type: OfficePageOrientation
Parameter Sets: Default
Aliases: None
Possible values: Portrait, Landscape

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PageSize
Target paper preset or physical dimensions.

```yaml
Type: OfficePageSize
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ResizeMargin
Margin reserved around fitted content.

```yaml
Type: OfficeLength
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ResizeMode
Fit behavior for source page content.

```yaml
Type: PdfPageResizeMode
Parameter Sets: Default
Aliases: None
Possible values: Fit, Fill, Stretch

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Custom target page width.

```yaml
Type: OfficeLength
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

- `OfficeIMO.Pdf.PdfMergeOptions`

## RELATED LINKS

- None
