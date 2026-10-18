---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficeWordPdfOptions
## SYNOPSIS
Creates discoverable Word-to-PDF conversion options for Export-OfficeDocumentPdf.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeWordPdfOptions [-PdfOptions <PdfOptions>] [-FontFamily <string>] [-PageSize <PageSize>] [-Orientation <OfficePageOrientation>] [-DefaultPageSize <WordPageSize>] [-DefaultOrientation <OfficePageOrientation>] [-MarginLeft <Double>] [-MarginTop <Double>] [-MarginRight <Double>] [-MarginBottom <Double>] [-Title <string>] [-Author <string>] [-Subject <string>] [-Keywords <string>] [-IncludePageNumbers] [-PageNumberFormat <string>] [-DefaultTableBorders] [-AllowSystemFontEmbedding] [-AllowDocumentFontEmbedding] [<CommonParameters>]
```

## DESCRIPTION
Creates discoverable Word-to-PDF conversion options for Export-OfficeDocumentPdf.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $options = New-OfficeWordPdfOptions -Title 'Service report' -Author 'Evotec' -IncludePageNumbers -AllowSystemFontEmbedding
Export-OfficeDocumentPdf -InputPath .\Report.docx -Path .\Report.pdf -WordOptions $options
```


## PARAMETERS

### -AllowDocumentFontEmbedding
Allow embedding installed fonts selected by a trusted Word document. System-font embedding must also be enabled.

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

### -AllowSystemFontEmbedding
Allow embedding fonts discovered on the current system.

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

### -Author
PDF author metadata.

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

### -DefaultOrientation
Fallback page orientation for sections without page settings.

```yaml
Type: OfficePageOrientation
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Portrait, Landscape

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DefaultPageSize
Fallback Word page size for sections without page settings.

```yaml
Type: WordPageSize
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Unknown, Letter, Legal, Statement, Executive, A3, A4, A5, A6, B5, Tabloid, B4Jis, Envelope9, Envelope10, CSheet, EnvelopeDl, EnvelopeC5, EnvelopeC4, EnvelopeB5, EnvelopeMonarch

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DefaultTableBorders
Draw default borders for tables that do not specify borders.

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

### -FontFamily
Default font family used when the document does not specify one.

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

### -IncludePageNumbers
Include page numbers in the generated PDF.

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

### -Keywords
PDF keywords metadata.

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

### -MarginBottom
Bottom page margin in PDF points.

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

### -MarginLeft
Left page margin in PDF points.

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

### -MarginRight
Right page margin in PDF points.

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

### -MarginTop
Top page margin in PDF points.

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

### -Orientation
PDF page orientation.

```yaml
Type: OfficePageOrientation
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Portrait, Landscape

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PageNumberFormat
Page number text format.

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

### -PageSize
PDF page size.

```yaml
Type: PageSize
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PdfOptions
Underlying low-level OfficeIMO PDF options.

```yaml
Type: PdfOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Subject
PDF subject metadata.

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

### -Title
PDF title metadata.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `OfficeIMO.Word.Pdf.WordToPdfOptions`

## RELATED LINKS

- None
