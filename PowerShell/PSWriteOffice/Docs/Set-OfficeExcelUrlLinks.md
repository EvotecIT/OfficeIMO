---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelUrlLinks
## SYNOPSIS
Converts cells in a range into external URL hyperlinks.

## SYNTAX
### Context (Default)
```powershell
Set-OfficeExcelUrlLinks [-Range] <string> -UrlScript <scriptblock> [-TitleScript <scriptblock>] [-NoStyle] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Set-OfficeExcelUrlLinks [-Range] <string> -Document <ExcelDocument> -UrlScript <scriptblock> [-Sheet <string>] [-SheetIndex <Int32>] [-TitleScript <scriptblock>] [-NoStyle] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Converts cells in a range into external URL hyperlinks.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ExcelSheet 'Summary' { Set-OfficeExcelUrlLinks -Range 'D2:D10' -UrlScript { param($text) "https://datatracker.ietf.org/doc/html/$text" } }
```

Turns each non-empty cell in D2:D10 into an external hyperlink.

## PARAMETERS

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
A1 range containing values to convert into external links.

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

### -TitleScript
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

### -UrlScript
Maps the cell text to a URL.

```yaml
Type: ScriptBlock
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: True
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
