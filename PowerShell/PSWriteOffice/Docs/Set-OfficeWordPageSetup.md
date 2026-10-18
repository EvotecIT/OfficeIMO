---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeWordPageSetup
## SYNOPSIS
Sets page setup options on Word sections.

## SYNTAX
### Current (Default)
```powershell
Set-OfficeWordPageSetup [-PageSize <WordPageSize>] [-Orientation <OfficePageOrientation>] [-Margin <WordMargin>] [-Left <Int32>] [-Right <Int32>] [-Top <Int32>] [-Bottom <Int32>] [-Header <Int32>] [-Footer <Int32>] [-Gutter <Int32>] [-Columns <Int32>] [-ColumnSpacing <Int32>] [-ColumnSeparator <Boolean>] [-PassThru] [<CommonParameters>]
```

### Section
```powershell
Set-OfficeWordPageSetup -Section <WordSection> [-PageSize <WordPageSize>] [-Orientation <OfficePageOrientation>] [-Margin <WordMargin>] [-Left <Int32>] [-Right <Int32>] [-Top <Int32>] [-Bottom <Int32>] [-Header <Int32>] [-Footer <Int32>] [-Gutter <Int32>] [-Columns <Int32>] [-ColumnSpacing <Int32>] [-ColumnSeparator <Boolean>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Set-OfficeWordPageSetup -Document <WordDocument> [-Index <int[]>] [-PageSize <WordPageSize>] [-Orientation <OfficePageOrientation>] [-Margin <WordMargin>] [-Left <Int32>] [-Right <Int32>] [-Top <Int32>] [-Bottom <Int32>] [-Header <Int32>] [-Footer <Int32>] [-Gutter <Int32>] [-Columns <Int32>] [-ColumnSpacing <Int32>] [-ColumnSeparator <Boolean>] [-PassThru] [<CommonParameters>]
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

### -Bottom
Bottom margin in twips.

```yaml
Type: Int32
Parameter Sets: Current, Section, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Columns
Number of section columns.

```yaml
Type: Int32
Parameter Sets: Current, Section, Document
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
Type: Boolean
Parameter Sets: Current, Section, Document
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
Type: Int32
Parameter Sets: Current, Section, Document
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

### -Footer
Footer distance in twips.

```yaml
Type: Int32
Parameter Sets: Current, Section, Document
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
Type: Int32
Parameter Sets: Current, Section, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Header
Header distance in twips.

```yaml
Type: Int32
Parameter Sets: Current, Section, Document
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

### -Left
Left margin in twips.

```yaml
Type: Int32
Parameter Sets: Current, Section, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Margin
Built-in margin preset.

```yaml
Type: WordMargin
Parameter Sets: Current, Section, Document
Aliases: None
Possible values: Normal, Mirrored, Moderate, Narrow, Wide, Office2003Default, Unknown

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
Parameter Sets: Current, Section, Document
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
Type: WordPageSize
Parameter Sets: Current, Section, Document
Aliases: None
Possible values: Unknown, Letter, Legal, Statement, Executive, A3, A4, A5, A6, B5, Tabloid, B4Jis, Envelope9, Envelope10, CSheet, EnvelopeDl, EnvelopeC5, EnvelopeC4, EnvelopeB5, EnvelopeMonarch

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
Parameter Sets: Current, Section, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Right
Right margin in twips.

```yaml
Type: Int32
Parameter Sets: Current, Section, Document
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

### -Top
Top margin in twips.

```yaml
Type: Int32
Parameter Sets: Current, Section, Document
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

- `OfficeIMO.Word.WordSection`
- `OfficeIMO.Word.WordDocument`

## OUTPUTS

- `OfficeIMO.Word.WordSection`

## RELATED LINKS

- None
