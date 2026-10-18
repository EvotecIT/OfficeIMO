---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficePdfList
## SYNOPSIS
Adds a bullet or numbered list to a PDF document.

## SYNTAX
### Context (Default)
```powershell
Add-OfficePdfList [-Items] <string[]> [-Numbered] [-StartNumber <int>] [-Align <PdfAlign>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficePdfList [-Items] <string[]> -Document <PdfDocument> [-Numbered] [-StartNumber <int>] [-Align <PdfAlign>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a bullet or numbered list to a PDF document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficePdf -Path .\Examples\Documents\PdfList.pdf {
    Add-OfficePdfHeading -Text 'Next actions'
    Add-OfficePdfList -Items 'Confirm owner','Publish summary','Schedule review' -Numbered
}
```

Adds a numbered list in the generated PDF flow.

## PARAMETERS

### -Align
List alignment.

```yaml
Type: PdfAlign
Parameter Sets: Context, Document
Aliases: None
Possible values: Left, Center, Right, Justify

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
PDF document to update outside the DSL context.

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

### -Items
List item text.

```yaml
Type: String[]
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Numbered
Create a numbered list instead of a bullet list.

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

### -StartNumber
Number to use for the first numbered item.

```yaml
Type: Int32
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
