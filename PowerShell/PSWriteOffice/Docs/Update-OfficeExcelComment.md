---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Update-OfficeExcelComment
## SYNOPSIS
Updates legacy worksheet comments (notes) that match a filter.

## SYNTAX
### Context (Default)
```powershell
Update-OfficeExcelComment [-Sheet <string>] [-SheetIndex <Int32>] [-Address <string>] [-Range <string>] [-MatchAuthor <string>] [-TextContains <string>] [-All] [-Text <string>] [-Run <Object[]>] [-Author <string>] [-Initials <string>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
Update-OfficeExcelComment [-Path] <string> [-Sheet <string>] [-SheetIndex <Int32>] [-Address <string>] [-Range <string>] [-MatchAuthor <string>] [-TextContains <string>] [-All] [-Text <string>] [-Run <Object[]>] [-Author <string>] [-Initials <string>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Update-OfficeExcelComment -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <Int32>] [-Address <string>] [-Range <string>] [-MatchAuthor <string>] [-TextContains <string>] [-All] [-Text <string>] [-Run <Object[]>] [-Author <string>] [-Initials <string>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Updates legacy worksheet comments (notes) that match a filter.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $updated = Update-OfficeExcelComment -Path .\Report.xlsx -Sheet Data -Address B2 -Text 'Reviewed' -Author Carol -Initials CC -PassThru
Get-OfficeExcelComment -Path .\Report.xlsx -Sheet Data -Address B2 |
    Select-Object Address, Author, Text
```

Replaces matching comment text and optionally changes the author.

## PARAMETERS

### -Address
A1 cell address to match.

```yaml
Type: String
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -All
Allow updating all comments on the selected worksheet(s) when no filter is supplied.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Author
Replacement author name.

```yaml
Type: String
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Workbook to update outside the DSL context.

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

### -Initials
Replacement author initials.

```yaml
Type: String
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MatchAuthor
Existing comment author to match, ignoring case.

```yaml
Type: String
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Returns the number of comments updated.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Workbook path to update.

```yaml
Type: String
Parameter Sets: Path
Aliases: InputPath, FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Range
A1 cell or range to match.

```yaml
Type: String
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Run
Replacement rich text runs.

```yaml
Type: Object[]
Parameter Sets: Context, Path, Document
Aliases: Runs
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Worksheet name to update. Defaults to the current DSL sheet or all workbook sheets.

```yaml
Type: String
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SheetIndex
Worksheet index (0-based) to update. Defaults to the current DSL sheet or all workbook sheets.

```yaml
Type: Int32
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Replacement plain text.

```yaml
Type: String
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TextContains
Existing text fragment to match, ignoring case.

```yaml
Type: String
Parameter Sets: Context, Path, Document
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

- `System.Int32`

## RELATED LINKS

- None
