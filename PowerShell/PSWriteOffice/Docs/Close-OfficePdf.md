---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Close-OfficePdf
## SYNOPSIS
Completes a live PDF document's lifetime without saving.

## SYNTAX
### __AllParameterSets
```powershell
Close-OfficePdf -Document <PdfDocument> [<CommonParameters>]
```

## DESCRIPTION
Completes a live PDF document's lifetime without saving.

## EXAMPLES

### EXAMPLE 1
```powershell
$pdf | Close-OfficePdf
```


## PARAMETERS

### -Document
Live PDF document to close.

```yaml
Type: PdfDocument
Parameter Sets: __AllParameterSets
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

- `OfficeIMO.Pdf.PdfDocument`

## OUTPUTS

- `None`

## RELATED LINKS

- None
