---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelDataValidationMessage
## SYNOPSIS
Sets prompt and error messages on existing Excel data validation rules.

## SYNTAX
### Context (Default)
```powershell
Set-OfficeExcelDataValidationMessage [-Sheet <string>] [-SheetIndex <Int32>] [-Range <string>] [-HeaderName <string>] [-TableName <string>] [-HeaderRow <int>] [-IncludeHeader] [-PromptTitle <string>] [-Prompt <string>] [-ErrorTitle <string>] [-ErrorMessage <string>] [-ShowInputMessage] [-ShowErrorMessage] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
Set-OfficeExcelDataValidationMessage [-Path] <string> [-Sheet <string>] [-SheetIndex <Int32>] [-Range <string>] [-HeaderName <string>] [-TableName <string>] [-HeaderRow <int>] [-IncludeHeader] [-PromptTitle <string>] [-Prompt <string>] [-ErrorTitle <string>] [-ErrorMessage <string>] [-ShowInputMessage] [-ShowErrorMessage] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Set-OfficeExcelDataValidationMessage -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <Int32>] [-Range <string>] [-HeaderName <string>] [-TableName <string>] [-HeaderRow <int>] [-IncludeHeader] [-PromptTitle <string>] [-Prompt <string>] [-ErrorTitle <string>] [-ErrorMessage <string>] [-ShowInputMessage] [-ShowErrorMessage] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Sets prompt and error messages on existing Excel data validation rules.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $rules = Set-OfficeExcelDataValidationMessage -Path .\Report.xlsx -Sheet Data -HeaderName Sales -TableName ServiceHealth -PromptTitle 'Sales' -Prompt 'Enter 1-1000' -ErrorTitle 'Invalid sales' -ErrorMessage 'Enter a whole number from 1 to 1000' -ShowInputMessage -ShowErrorMessage -PassThru
$rules |
    Select-Object Range, PromptTitle, ErrorTitle
```

Updates validation metadata for matching rules and saves the workbook.

## PARAMETERS

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

### -ErrorMessage
Error message text. Omit or pass null to clear the message.

```yaml
Type: String
Parameter Sets: Context, Path, Document
Aliases: Error
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ErrorTitle
Error title. Omit or pass null to clear the title.

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

### -HeaderName
Header or table column name used to resolve the validation rules to update.

```yaml
Type: String
Parameter Sets: Context, Path, Document
Aliases: ColumnName
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HeaderRow
Worksheet header row used when resolving HeaderName without a table. Use 0 for the first row of the used range.

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

### -IncludeHeader
Include the header cell in the resolved range.

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

### -PassThru
Returns matching validation rules after updating them.

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

### -Prompt
Input prompt text. Omit or pass null to clear the prompt.

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

### -PromptTitle
Input prompt title. Omit or pass null to clear the title.

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

### -Range
A1 range used to select existing validation rules.

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

### -Sheet
Worksheet name to update. Defaults to the current DSL sheet.

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
Worksheet index (0-based) to update. Defaults to the current DSL sheet.

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

### -ShowErrorMessage
Forces Excel to show the validation error.

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

### -ShowInputMessage
Forces Excel to show the input prompt.

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

### -TableName
Optional table name for header-based range resolution.

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

- `System.Management.Automation.PSObject`

## RELATED LINKS

- None
