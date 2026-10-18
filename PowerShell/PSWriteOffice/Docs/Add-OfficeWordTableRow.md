---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeWordTableRow
## SYNOPSIS
Appends a row to an existing Word table.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeWordTableRow [-Table] <WordTable> [[-Values] <Object>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a new row to a WordTable that was already created or found in an existing
document. The command accepts scalar values, arrays, dictionaries, ordered dictionaries, and
PowerShell objects. Values are expanded from left to right across cells; missing values become empty
cells. This keeps existing-document editing simple without forcing callers back into the Word DSL.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $doc = Get-OfficeWord -Path .\Report.docx
$doc | Get-OfficeWordTable | Select-Object -First 1 |
    Add-OfficeWordTableRow -Values 'Service', 'Ready', 'Low'
```

Adds one table row and writes the supplied values into its cells.

### EXAMPLE 2
```powershell
PS> $doc = Get-OfficeWord -Path .\Report.docx
$table = Find-OfficeWordTable -Document $doc -Text 'Risk marker' | Select-Object -First 1
$table | Add-OfficeWordTableRow -Values ([ordered]@{
    Item  = 'Mitigation plan'
    Owner = 'Service Desk'
    State = 'Ready'
})
$doc | Close-OfficeWord -Save
```

Uses an ordered dictionary so values are written into predictable table columns.

## PARAMETERS

### -PassThru
Emit the created row for additional OfficeIMO-level edits.

```yaml
Type: SwitchParameter
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
Existing Word table to append to, usually from Get-OfficeWordTable or Find-OfficeWordTable.

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

### -Values
Values to write into the new row. Arrays, dictionaries, ordered dictionaries, and objects are expanded across cells.

```yaml
Type: Object
Parameter Sets: __AllParameterSets
Aliases: Data, InputObject
Possible values:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Word.WordTable`

## OUTPUTS

- `OfficeIMO.Word.WordTableRow`

## RELATED LINKS

- None
