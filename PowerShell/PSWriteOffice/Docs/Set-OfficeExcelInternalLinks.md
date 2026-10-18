---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelInternalLinks
## SYNOPSIS
Converts cells in a range into internal workbook links.

## SYNTAX
### Context (Default)
```powershell
Set-OfficeExcelInternalLinks [-Range] <string> [-DestinationSheetScript <scriptblock>] [-DisplayScript <scriptblock>] [-TargetAddress <string>] [-NoStyle] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Set-OfficeExcelInternalLinks [-Range] <string> -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <Int32>] [-DestinationSheetScript <scriptblock>] [-DisplayScript <scriptblock>] [-TargetAddress <string>] [-NoStyle] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Converts cells in a range into internal workbook links.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ExcelSheet 'Summary' { Set-OfficeExcelInternalLinks -Range 'A2:A10' }
```

Turns each non-empty cell in A2:A10 into an internal link to the sheet with the same name.

## PARAMETERS

### -DestinationSheetScript
Optional mapping from cell text to destination sheet name.

```yaml
Type: ScriptBlock
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DisplayScript
Optional mapping from cell text to display text.

```yaml
Type: ScriptBlock
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
Workbook to operate on outside the DSL context.

```yaml
Type: ExcelDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -NoStyle
Skip hyperlink styling (blue + underline).

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
Emit the worksheet after creating links.

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

### -Range
A1 range containing values to convert into internal links.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Worksheet name when using Document.

```yaml
Type: String
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SheetIndex
Worksheet index (0-based) when using Document.

```yaml
Type: Int32
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TargetAddress
Destination cell on the target sheet.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `OfficeIMO.Excel.ExcelSheet`

## RELATED LINKS

- None
