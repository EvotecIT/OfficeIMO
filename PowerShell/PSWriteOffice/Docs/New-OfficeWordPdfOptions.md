---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeWordPdfOptions
## SYNOPSIS
Creates discoverable Word-to-PDF conversion options for Convert-OfficeDocument.

## SYNTAX
### Default (Default)
```powershell
New-OfficeWordPdfOptions [-PdfOptions <PdfCreateOptions>] [-Font <string>] [-PageSize <PdfPageSize>] [-Orientation <OfficePageOrientation>] [-FallbackPageSize <PdfPageSize>] [-FallbackOrientation <OfficePageOrientation>] [-MarginLeft <OfficeLength>] [-MarginTop <OfficeLength>] [-MarginRight <OfficeLength>] [-MarginBottom <OfficeLength>] [-Title <string>] [-Author <string>] [-Subject <string>] [-Keywords <string>] [-IncludePageNumbers] [-PageNumberFormat <string>] [-DefaultTableBorders] [-AllowSystemFontEmbedding] [-AllowDocumentFontEmbedding] [<CommonParameters>]
```

## DESCRIPTION
Creates discoverable Word-to-PDF conversion options for Convert-OfficeDocument.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $options = New-OfficeWordPdfOptions -Title 'Service report' -Author 'Evotec' -IncludePageNumbers -AllowSystemFontEmbedding
Convert-OfficeDocument -Path .\Report.docx -To Pdf -OutputPath .\Report.pdf -Options $options
```


## PARAMETERS

### -AllowDocumentFontEmbedding
Allow embedding installed fonts selected by a trusted Word document. System-font embedding must also be enabled.

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

### -AllowSystemFontEmbedding
Allow embedding fonts discovered on the current system.

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

### -Author
PDF author metadata.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

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
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FallbackOrientation
Fallback page orientation for sections without page settings.

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

### -FallbackPageSize
Fallback physical PDF page size for sections without page settings.

```yaml
Type: PdfPageSize
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Font
Default font family used when the document does not specify one.

```yaml
Type: String
Parameter Sets: Default
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
Parameter Sets: Default
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
Parameter Sets: Default
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

### -MarginLeft
Left page margin in PDF points.

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

### -MarginRight
Right page margin in PDF points.

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

### -MarginTop
Top page margin in PDF points.

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
PDF page orientation.

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

### -PageNumberFormat
Page number text format.

```yaml
Type: String
Parameter Sets: Default
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
Type: PdfPageSize
Parameter Sets: Default
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
Type: PdfCreateOptions
Parameter Sets: Default
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
Parameter Sets: Default
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

- `OfficeIMO.Word.Pdf.WordPdfExportOptions`

## RELATED LINKS

- None
