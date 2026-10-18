---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeWordTableCell
## SYNOPSIS
Gets cells from an OfficeIMO Word table.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeWordTableCell [-Table] <WordTable> [-Row <Int32>] [-Column <Int32>] [<CommonParameters>]
```

## DESCRIPTION
Gets cells from an OfficeIMO Word table.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $table = Get-OfficeWordTable -Path .\Report.docx | Select-Object -First 1
$cell = $table | Get-OfficeWordTableCell -Row 1 -Column 2
$cell.Paragraphs |
    Select-Object -Property Text
```

Gets a zero-based table cell from an OfficeIMO table object and inspects its paragraphs.

## PARAMETERS

### -Column
Optional zero-based column index.

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

### -Row
Optional zero-based row index.

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

### -Table
Table to inspect.

```yaml
Type: WordTable
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Word.WordTable`

## OUTPUTS

- `OfficeIMO.Word.WordTableCell`

## RELATED LINKS

- None
