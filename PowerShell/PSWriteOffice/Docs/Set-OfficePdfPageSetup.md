---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficePdfPageSetup
## SYNOPSIS
Sets PDF page size, orientation, and margins.

## SYNTAX
### Context (Default)
```powershell
Set-OfficePdfPageSetup [-PageSize <string>] [-Width <Double>] [-Height <Double>] [-Landscape] [-Margin <Double>] [-Left <Double>] [-Top <Double>] [-Right <Double>] [-Bottom <Double>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Set-OfficePdfPageSetup -Document <PdfDocument> [-PageSize <string>] [-Width <Double>] [-Height <Double>] [-Landscape] [-Margin <Double>] [-Left <Double>] [-Top <Double>] [-Right <Double>] [-Bottom <Double>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Sets PDF page size, orientation, and margins.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficePdf -Path .\Examples\Documents\PdfPageSetup.pdf {
    Set-OfficePdfPageSetup -PageSize A4 -Margin 42
    Add-OfficePdfHeading -Text 'A4 report'
    Add-OfficePdfParagraph -Text 'The report uses custom margins.'
}
```

Applies page setup before adding generated PDF content.

## PARAMETERS

### -Bottom
Bottom margin in PDF points.

```yaml
Type: Double
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Compatibility parameter. Page composition is supported only inside New-OfficePdf.

```yaml
Type: PdfDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Height
Custom page height in PDF points when -PageSize Custom is used.

```yaml
Type: Double
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Landscape
Use landscape orientation.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Left
Left margin in PDF points.

```yaml
Type: Double
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Margin
Uniform margin in PDF points.

```yaml
Type: Double
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PageSize
Page size name: A4, A5, Letter, Legal, or Custom.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the updated document.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Right
Right margin in PDF points.

```yaml
Type: Double
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Top
Top margin in PDF points.

```yaml
Type: Double
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Custom page width in PDF points when -PageSize Custom is used.

```yaml
Type: Double
Parameter Sets: Context, Document
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

- `OfficeIMO.Pdf.PdfDocument`

## OUTPUTS

- `OfficeIMO.Pdf.PdfDocument`

## RELATED LINKS

- None
