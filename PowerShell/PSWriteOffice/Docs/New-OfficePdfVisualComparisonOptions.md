---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficePdfVisualComparisonOptions
## SYNOPSIS
Creates native PDF comparison settings and independent bounded policies for both inputs.

## SYNTAX
### Default (Default)
```powershell
New-OfficePdfVisualComparisonOptions [-Scale <Double>] [-ChannelTolerance <Int32>] [-AllowedDifferenceRatio <Double>] [-Alignment <PdfVisualPageAlignment>] [-BackgroundColor <OfficeColor>] [-MaxPages <Int32>] [-MaximumPixelCountPerImage <Int64>] [-MaximumTotalPixelCount <Int64>] [-MaxTotalOutputBytes <Int64>] [-IgnorePermissionRestrictions] [-IgnoreDifferencePermissionRestrictions] [-ReferenceLoadOptions <PdfLoadOptions>] [-DifferenceLoadOptions <PdfLoadOptions>] [<CommonParameters>]
```

## DESCRIPTION
Creates native PDF comparison settings and independent bounded policies for both inputs.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficePdfVisualComparisonOptions -Alignment 'Value'
```


## PARAMETERS

### -Alignment
Native page-image alignment.

```yaml
Type: PdfVisualPageAlignment
Parameter Sets: Default
Aliases: None
Possible values: TopLeft, Center

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AllowedDifferenceRatio
Maximum differing-pixel ratio accepted as equal.

```yaml
Type: Double
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BackgroundColor
Raster background.

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

### -ChannelTolerance
Maximum per-channel difference treated as equal.

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

### -DifferenceLoadOptions
Bounded policy for the difference input.

```yaml
Type: PdfLoadOptions
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IgnoreDifferencePermissionRestrictions
Explicitly ignore authenticated owner-imposed restrictions on the difference PDF.

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

### -IgnorePermissionRestrictions
Explicitly ignore authenticated owner-imposed restrictions on the reference PDF.

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

### -MaximumPixelCountPerImage
Maximum pixels in one expected, actual or difference image.

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

### -MaximumTotalPixelCount
Maximum aggregate raster pixels.

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

### -MaxPages
Maximum pages on either side.

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

### -MaxTotalOutputBytes
Maximum retained PNG bytes.

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

### -ReferenceLoadOptions
Bounded policy for the reference input.

```yaml
Type: PdfLoadOptions
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Scale
Native render scale.

```yaml
Type: Double
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

- `PSWriteOffice.Services.Comparison.OfficePdfVisualComparisonOptions`: Native comparison settings and distinct bounded input policies for the two PDF roles.

## RELATED LINKS

- None
