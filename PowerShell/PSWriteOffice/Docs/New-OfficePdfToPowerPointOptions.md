---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficePdfToPowerPointOptions
## SYNOPSIS
Creates discoverable PDF-to-PowerPoint reconstruction settings.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficePdfToPowerPointOptions [-Mode <PdfPowerPointImportMode>] [-PageRange <string>] [-Dpi <Double>] [-MaxPages <Int32>] [-MaxPixelsPerPage <Int64>] [-MaxOutputBytesPerPage <Int64>] [-MaxTotalOutputBytes <Int64>] [-MaxEditableObjectsPerPage <Int32>] [-MaxRows <Int32>] [-MergePageContinuations] [-SuppressRepeatedBodyHeaderRows] [-MaxRowsPerSlide <Int32>] [-MaxColumnsPerSlide <Int32>] [-TableStyle <PowerPointTableStylePreset>] [-IncludeSourceTitles] [-IncludeColumnHeaderRows] [-BandedRows] [-AlignNumericColumns] [-EmptyPresentationTitle <string>] [-EmptyPresentationMessage <string>] [<CommonParameters>]
```

## DESCRIPTION
Creates discoverable PDF-to-PowerPoint reconstruction settings.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $options = New-OfficePdfToPowerPointOptions -PageRange '1-5' -MaxPages 5 -IncludeSourceTitles
ConvertTo-OfficePdfPowerPoint -Path .\Source.pdf -OutputPath .\Slides.pptx -Options $options
```


## PARAMETERS

### -AlignNumericColumns
Right-align inferred numeric columns.

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

### -BandedRows
Enable banded-row styling.

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

### -Dpi
Raster resolution used by visual import.

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

### -EmptyPresentationMessage
Message used when no supported content is detected.

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

### -EmptyPresentationTitle
Title used when no supported content is detected.

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

### -IncludeColumnHeaderRows
Add inferred column headers.

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

### -IncludeSourceTitles
Add source-page titles.

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

### -MaxColumnsPerSlide
Maximum columns written to one slide; zero means unlimited.

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

### -MaxEditableObjectsPerPage
Maximum editable objects reconstructed per page.

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

### -MaxOutputBytesPerPage
Maximum encoded bytes per rendered page.

```yaml
Type: Int64
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxPages
Maximum pages imported.

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

### -MaxPixelsPerPage
Maximum pixels per rendered page.

```yaml
Type: Int64
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxRows
Maximum body rows imported per table; zero means unlimited.

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

### -MaxRowsPerSlide
Maximum rows written to one slide; zero means unlimited.

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

### -MaxTotalOutputBytes
Maximum aggregate encoded output bytes.

```yaml
Type: Int64
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MergePageContinuations
Merge compatible table segments across pages.

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

### -Mode
Visual, editable-table, hybrid, editable-content, or automatic import mode.

```yaml
Type: PdfPowerPointImportMode
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: VisualPages, EditableTables, HybridVisualAndEditableTables, EditableContent, Auto

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PageRange
Optional one-based page ranges such as 1-3,5.

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

### -SuppressRepeatedBodyHeaderRows
Suppress repeated body header rows.

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

### -TableStyle
PowerPoint table style.

```yaml
Type: PowerPointTableStylePreset
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

- `OfficeIMO.PowerPoint.Pdf.PdfToPowerPointOptions`

## RELATED LINKS

- None
