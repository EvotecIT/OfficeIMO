---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeExcelTableOfContents
## SYNOPSIS
Adds or refreshes a table-of-contents sheet with links to every sheet.

## SYNTAX
### Default (Default)
```powershell
Add-OfficeExcelTableOfContents [-Document <ExcelDocument>] [-Name <string>] [-NoMoveFirst] [-NoHyperlinks] [-IncludeNamedRanges] [-IncludeHiddenNamedRanges] [-NoStyle] [-BackLink] [-BackLinkAddress <string>] [-BackLinkText <string>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds or refreshes a table-of-contents sheet with links to every sheet.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeExcel -Path .\Reports\Board-Pack.xlsx {
    ExcelTableOfContents
}
```

Adds a first sheet with a link to every other sheet.

### EXAMPLE 2
```powershell
$orders = Import-OfficeCsv -Path .\Data\Q3-Orders.csv
New-OfficeExcel -Path .\Reports\Q3-Regions.xlsx {
    foreach ($group in $orders | Group-Object Region) {
        ExcelSheet $group.Name {
            ExcelTable $group.Group -Address A3
        }
    }
    ExcelTableOfContents -Name 'Contents' -BackLink -BackLinkAddress A1 -BackLinkText 'Back to contents'
}
```

Writes one sheet per region, then adds a contents sheet with links and a back link on each regional sheet.

## PARAMETERS

### -BackLink
Adds a link from each listed worksheet back to the table of contents.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BackLinkAddress
A1 destination for worksheet back links; omission chooses an unused cell.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BackLinkText
Text used for back links.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Caller-owned workbook; omission uses the enclosing Excel content block.

```yaml
Type: ExcelDocument
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -IncludeHiddenNamedRanges
Include hidden named ranges when listing named ranges.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeNamedRanges
Include named ranges in the TOC.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Name
Named range name.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoHyperlinks
Disable internal hyperlinks in the TOC sheet.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoMoveFirst
Keeps the table-of-contents worksheet at its current position.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoStyle
Disable formatted TOC styling.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emits the native result after the operation.

```yaml
Type: SwitchParameter
Parameter Sets: Default
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

- `None`

## RELATED LINKS

- None
