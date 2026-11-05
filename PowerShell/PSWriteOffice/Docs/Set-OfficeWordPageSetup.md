---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficeWordPageSetup
## SYNOPSIS
Sets page setup options on Word sections.

## SYNTAX
### Context (Default)
```powershell
Set-OfficeWordPageSetup [-PageSize <OfficePageSize>] [-Orientation <OfficePageOrientation>] [-MarginPreset <WordMargin>] [-MarginLeft <OfficeLength>] [-MarginRight <OfficeLength>] [-MarginTop <OfficeLength>] [-MarginBottom <OfficeLength>] [-HeaderDistance <OfficeLength>] [-FooterDistance <OfficeLength>] [-Gutter <OfficeLength>] [-Columns <Int32>] [-ColumnSpacing <OfficeLength>] [-ColumnSeparator] [-PassThru] [-Margin <OfficeLength>] [<CommonParameters>]
```

### Section
```powershell
Set-OfficeWordPageSetup -Section <WordSection> [-PageSize <OfficePageSize>] [-Orientation <OfficePageOrientation>] [-MarginPreset <WordMargin>] [-MarginLeft <OfficeLength>] [-MarginRight <OfficeLength>] [-MarginTop <OfficeLength>] [-MarginBottom <OfficeLength>] [-HeaderDistance <OfficeLength>] [-FooterDistance <OfficeLength>] [-Gutter <OfficeLength>] [-Columns <Int32>] [-ColumnSpacing <OfficeLength>] [-ColumnSeparator] [-PassThru] [-Margin <OfficeLength>] [<CommonParameters>]
```

### Document
```powershell
Set-OfficeWordPageSetup -Document <WordDocument> [-Index <int[]>] [-PageSize <OfficePageSize>] [-Orientation <OfficePageOrientation>] [-MarginPreset <WordMargin>] [-MarginLeft <OfficeLength>] [-MarginRight <OfficeLength>] [-MarginTop <OfficeLength>] [-MarginBottom <OfficeLength>] [-HeaderDistance <OfficeLength>] [-FooterDistance <OfficeLength>] [-Gutter <OfficeLength>] [-Columns <Int32>] [-ColumnSpacing <OfficeLength>] [-ColumnSeparator] [-PassThru] [-Margin <OfficeLength>] [<CommonParameters>]
```

## DESCRIPTION
Updates page size, orientation, margins, and section columns through OfficeIMO.Word.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Add-OfficeWordSection { Set-OfficeWordPageSetup -PageSize A4 -Orientation Landscape -Columns 2 }
```

Updates the current section page setup.

## PARAMETERS

### -Columns
Number of section columns.

```yaml
Type: Int32
Parameter Sets: Context, Section, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ColumnSeparator
Whether to show a separator between columns.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Section, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ColumnSpacing
Space between columns in twips.

```yaml
Type: OfficeLength
Parameter Sets: Context, Section, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Document whose sections should be updated.

```yaml
Type: WordDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -FooterDistance
FooterDistance distance in twips.

```yaml
Type: OfficeLength
Parameter Sets: Context, Section, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Gutter
Gutter size in twips.

```yaml
Type: OfficeLength
Parameter Sets: Context, Section, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HeaderDistance
HeaderDistance distance in twips.

```yaml
Type: OfficeLength
Parameter Sets: Context, Section, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Index
Optional 0-based section indexes when -Document is used.

```yaml
Type: Int32[]
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Margin
Uniform physical margin; numbers mean points.

```yaml
Type: OfficeLength
Parameter Sets: Context, Section, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MarginBottom
MarginBottom margin in twips.

```yaml
Type: OfficeLength
Parameter Sets: Context, Section, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MarginLeft
MarginLeft margin in twips.

```yaml
Type: OfficeLength
Parameter Sets: Context, Section, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MarginPreset
Built-in margin preset.

```yaml
Type: WordMargin
Parameter Sets: Context, Section, Document
Aliases: None
Possible values: Normal, Mirrored, Moderate, Narrow, Wide, Office2003Default, Unknown

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MarginRight
MarginRight margin in twips.

```yaml
Type: OfficeLength
Parameter Sets: Context, Section, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MarginTop
MarginTop margin in twips.

```yaml
Type: OfficeLength
Parameter Sets: Context, Section, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Orientation
Page orientation.

```yaml
Type: OfficePageOrientation
Parameter Sets: Context, Section, Document
Aliases: None
Possible values: Portrait, Landscape

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PageSize
Built-in page size.

```yaml
Type: OfficePageSize
Parameter Sets: Context, Section, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit updated sections.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Section, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Section
Section to update.

```yaml
Type: WordSection
Parameter Sets: Section
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Word.WordSection`
- `OfficeIMO.Word.WordDocument`

## OUTPUTS

- `OfficeIMO.Word.WordSection`

## RELATED LINKS

- None
