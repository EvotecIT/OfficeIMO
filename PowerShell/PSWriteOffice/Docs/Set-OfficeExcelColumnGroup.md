---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelColumnGroup
## SYNOPSIS
Configures collapsible Excel outline grouping for worksheet columns.

## SYNTAX
### __AllParameterSets
```powershell
Set-OfficeExcelColumnGroup [[-StartColumn] <Object>] [[-EndColumn] <Object>] [-StartColumnName <string>] [-EndColumnName <string>] [-OutlineLevel <int>] [-Collapsed] [-Hidden] [-Clear] [-KeepHidden] [-SummaryRight <Boolean>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Configures collapsible Excel outline grouping for worksheet columns.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ExcelSheet 'Data' { Set-OfficeExcelColumnGroup -StartColumn B -EndColumn D -Collapsed }
```

Applies Excel column outline metadata using OfficeIMO.

## PARAMETERS

### -Clear
Clear column grouping metadata from the target range.

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

### -Collapsed
Hide the grouped columns and mark the following summary column as collapsed.

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

### -EndColumn
Last 1-based column in the group. Defaults to the start column.

```yaml
Type: Object
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -EndColumnName
Last column letter in the group.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: EndColumnLetter
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Hidden
Hide the grouped columns without marking the group collapsed.

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

### -KeepHidden
Keep hidden columns hidden when clearing column grouping metadata.

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

### -OutlineLevel
Excel outline level from 1 through 7.

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

### -PassThru
Emit the object created or changed by the command.

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

### -StartColumn
First 1-based column in the group.

```yaml
Type: Object
Parameter Sets: __AllParameterSets
Aliases: Column
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StartColumnName
First column letter in the group.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: StartColumnLetter, ColumnName, ColumnLetter, Letter
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SummaryRight
Set whether column summary controls appear to the right of grouped columns.

```yaml
Type: Boolean
Parameter Sets: __AllParameterSets
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

- `None`

## OUTPUTS

- `None`

## RELATED LINKS

- None
