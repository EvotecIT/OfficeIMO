---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficeRtfPdfOptions
## SYNOPSIS
Creates discoverable RTF-to-PDF conversion options for Export-OfficeDocumentPdf.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeRtfPdfOptions [-PdfOptions <PdfOptions>] [-IncludeHiddenText] [-IncludeImages] [-DefaultImageWidth <Double>] [-DefaultImageHeight <Double>] [-IncludeMetadata] [-IncludeTables] [-IncludeHeaderFooters] [-IncludeNotes] [-MaximumSystemFontFamilies <Int32>] [-AllowSystemFontEmbedding] [-AllowDocumentFontEmbedding] [<CommonParameters>]
```

## DESCRIPTION
Creates discoverable RTF-to-PDF conversion options for Export-OfficeDocumentPdf.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $options = New-OfficeRtfPdfOptions -IncludeImages -IncludeTables -IncludeHeaderFooters -MaximumSystemFontFamilies 32
Export-OfficeDocumentPdf -InputPath .\Report.rtf -Path .\Report.pdf -RtfOptions $options
```


## PARAMETERS

### -AllowDocumentFontEmbedding
Allow embedding fonts referenced by the RTF document.

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

### -DefaultImageHeight
Fallback image height in PDF points.

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

### -DefaultImageWidth
Fallback image width in PDF points.

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

### -IncludeHeaderFooters
Render headers and footers.

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

### -IncludeHiddenText
Include text marked hidden.

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

### -IncludeImages
Render images.

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

### -IncludeMetadata
Copy document metadata into the PDF.

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

### -IncludeNotes
Render document notes.

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

### -IncludeTables
Render tables.

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

### -MaximumSystemFontFamilies
Maximum number of system font families to discover.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `OfficeIMO.Rtf.Pdf.RtfToPdfOptions`

## RELATED LINKS

- None
