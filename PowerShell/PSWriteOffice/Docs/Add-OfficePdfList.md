---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficePdfList
## SYNOPSIS
Adds a bulleted or numbered list built from PdfListItem children.

## SYNTAX
### Context (Default)
```powershell
Add-OfficePdfList [-Content] <scriptblock> [-Numbered] [-StartNumber <Int32>] [-Alignment <OfficeAlignment>] [<CommonParameters>]
```

### Document
```powershell
Add-OfficePdfList [-Content] <scriptblock> -Document <PdfDocument> [-Numbered] [-StartNumber <Int32>] [-Alignment <OfficeAlignment>] [-PassThru] [<CommonParameters>]
```

### Items
```powershell
Add-OfficePdfList [-Items] <string[]> [-Numbered] [-StartNumber <Int32>] [-Alignment <OfficeAlignment>] [<CommonParameters>]
```

### DocumentItems
```powershell
Add-OfficePdfList [-Items] <string[]> -Document <PdfDocument> [-Numbered] [-StartNumber <Int32>] [-Alignment <OfficeAlignment>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a bulleted or numbered list built from PdfListItem children.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficePdf -Path .\Actions.pdf {
    PdfList {
        PdfListItem 'Confirm the service owner'
        PdfListItem 'Publish the summary'
        PdfListItem 'Schedule the review'
    }
}
```


## PARAMETERS

### -Alignment
Alignment for this operation.

```yaml
Type: OfficeAlignment
Parameter Sets: Context, Document, Items, DocumentItems
Aliases: None
Possible values: Left, Center, Right, Justify, Start, End

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Content
Content for this operation.

```yaml
Type: ScriptBlock
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Document for this operation.

```yaml
Type: PdfDocument
Parameter Sets: Document, DocumentItems
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Items
Items for this operation.

```yaml
Type: String[]
Parameter Sets: Items, DocumentItems
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Numbered
Numbered for this operation.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document, Items, DocumentItems
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Document set only; inside a block the document is returned by New-OfficePdf.

```yaml
Type: SwitchParameter
Parameter Sets: Document, DocumentItems
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StartNumber
StartNumber for this operation.

```yaml
Type: Int32
Parameter Sets: Context, Document, Items, DocumentItems
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
