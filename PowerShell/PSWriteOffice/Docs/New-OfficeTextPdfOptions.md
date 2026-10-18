---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficeTextPdfOptions
## SYNOPSIS
Creates literal text decoding and PDF layout settings for Export-OfficeDocumentPdf.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeTextPdfOptions [-EncodingName <string>] [-TabSize <int>] [-MaximumCharacters <int>] [-MaximumPages <int>] [-PdfOptions <PdfOptions>] [<CommonParameters>]
```

## DESCRIPTION
Creates literal text decoding and PDF layout settings for Export-OfficeDocumentPdf.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $options = New-OfficeTextPdfOptions -EncodingName utf-16 -TabSize 4
Export-OfficeDocumentPdf -InputPath .\Report.txt -Path .\Report.pdf -TextOptions $options
```


## PARAMETERS

### -EncodingName
Explicit encoding; otherwise use a Unicode BOM or strict UTF-8.

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

### -MaximumCharacters
Maximum decoded and expanded characters.

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

### -MaximumPages
Maximum generated pages.

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
PDF fonts and page geometry; default is ten-point Courier.

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

### -TabSize
Source columns per tab stop.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `OfficeIMO.Pdf.PdfPlainTextOptions`

## RELATED LINKS

- None
