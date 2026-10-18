---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Clear-OfficePdfBackgroundShape
## SYNOPSIS
Clears generated PDF page background shapes.

## SYNTAX
### Context (Default)
```powershell
Clear-OfficePdfBackgroundShape [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Clear-OfficePdfBackgroundShape -Document <PdfDocument> [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Clears generated PDF page background shapes.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficePdf -Path .\Examples\Documents\PdfNoBackgroundShape.pdf {
    Add-OfficePdfBackgroundShape -Shape Rectangle -FillColor '#EEF2FF' -X 0 -Y 0 -Width 595 -Height 120
    Clear-OfficePdfBackgroundShape
    Add-OfficePdfHeading -Text 'Clean variant'
}
```

Clears generated page background shapes while the PDF page is being composed.

## PARAMETERS

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Pdf.PdfDocument`

## OUTPUTS

- `OfficeIMO.Pdf.PdfDocument`

## RELATED LINKS

- None
