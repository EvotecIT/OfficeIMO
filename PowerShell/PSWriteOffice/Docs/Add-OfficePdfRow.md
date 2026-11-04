---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficePdfRow
## SYNOPSIS
Adds a row of side-by-side PdfColumn blocks.

## SYNTAX
### Context (Default)
```powershell
Add-OfficePdfRow [-Content] <scriptblock> [-Gap <OfficeLength>] [-SpacingBefore <OfficeLength>] [-SpacingAfter <OfficeLength>] [-KeepTogether] [-KeepWithNext] [-ColumnSeparatorColor <OfficeColor>] [-ColumnSeparatorWidth <OfficeLength>] [<CommonParameters>]
```

### Document
```powershell
Add-OfficePdfRow [-Content] <scriptblock> -Document <PdfDocument> [-Gap <OfficeLength>] [-SpacingBefore <OfficeLength>] [-SpacingAfter <OfficeLength>] [-KeepTogether] [-KeepWithNext] [-ColumnSeparatorColor <OfficeColor>] [-ColumnSeparatorWidth <OfficeLength>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a row of side-by-side PdfColumn blocks.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficePdf -Path .\TwoColumns.pdf {
    PdfRow {
        PdfColumn { PdfParagraph 'Left column' }
        PdfColumn { PdfParagraph 'Right column' }
    }
}
```


## PARAMETERS

### -ColumnSeparatorColor
ColumnSeparatorColor for this operation.

```yaml
Type: OfficeColor
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ColumnSeparatorWidth
ColumnSeparatorWidth for this operation.

```yaml
Type: OfficeLength
Parameter Sets: Context, Document
Aliases: None
Possible values:

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
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Gap
Gap for this operation.

```yaml
Type: OfficeLength
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -KeepTogether
KeepTogether for this operation.

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

### -KeepWithNext
KeepWithNext for this operation.

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
Document set only; inside a block the document is returned by New-OfficePdf.

```yaml
Type: SwitchParameter
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SpacingAfter
SpacingAfter for this operation.

```yaml
Type: OfficeLength
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SpacingBefore
SpacingBefore for this operation.

```yaml
Type: OfficeLength
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
