---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficeExcelValidation
## SYNOPSIS
Sets prompt and error messages on existing validation rules.

## SYNTAX
### Range (Default)
```powershell
Set-OfficeExcelValidation [[-Range] <string>] [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-PromptTitle <string>] [-Prompt <string>] [-ErrorTitle <string>] [-ErrorMessage <string>] [-ShowInputMessage] [-ShowErrorMessage] [-PassThru] [<CommonParameters>]
```

### Header
```powershell
Set-OfficeExcelValidation -Header <string> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-PromptTitle <string>] [-Prompt <string>] [-ErrorTitle <string>] [-ErrorMessage <string>] [-ShowInputMessage] [-ShowErrorMessage] [-PassThru] [-Table <string>] [-HeaderRow <Int32>] [-IncludeHeader] [<CommonParameters>]
```

## DESCRIPTION
Sets prompt and error messages on existing validation rules.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeExcel -Path .\Reports\Helpdesk.xlsx {
    ExcelSheet 'Tickets' {
        Set-OfficeExcelValidation -Header Priority -Table Tickets -PromptTitle 'Priority' -Prompt 'P1 is a production outage; P4 is a question.' -ShowInputMessage
    }
}
```

Shows a hint when someone selects a cell in the Priority column.

### EXAMPLE 2
```powershell
Open-OfficeExcel -Path .\Forms\Equipment-Requests.xlsx {
    ExcelSheet 'Requests' {
        Set-OfficeExcelValidation -Range E2:E500 -ErrorTitle 'Cost centre' -ErrorMessage 'Pick a cost centre from the list. Ask Finance if yours is missing.' -ShowErrorMessage
    }
}
```

Replaces the generic error text of the cost-centre rule in an existing form with guidance for the requester.

## PARAMETERS

### -Document
Caller-owned workbook; omission uses the enclosing Excel content block.

```yaml
Type: ExcelDocument
Parameter Sets: Range, Header
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -ErrorMessage
Error message text. Omit or pass null to clear the message.

```yaml
Type: String
Parameter Sets: Range, Header
Aliases: None
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
Parameter Sets: Range, Header
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Header
Header text identifying the target column.

```yaml
Type: String
Parameter Sets: Header
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HeaderRow
One-based row containing the column headers.

```yaml
Type: Int32
Parameter Sets: Header
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeHeader
Includes the header cell in the resolved column range.

```yaml
Type: SwitchParameter
Parameter Sets: Header
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
Parameter Sets: Range, Header
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Prompt
Input prompt text. Omit or pass null to clear the prompt.

```yaml
Type: String
Parameter Sets: Range, Header
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
Parameter Sets: Range, Header
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Range
A1 cell range on the selected worksheet.

```yaml
Type: String
Parameter Sets: Range
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.

```yaml
Type: OfficeSheetSelector
Parameter Sets: Range, Header
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
Parameter Sets: Range, Header
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
Parameter Sets: Range, Header
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Table
Native Excel table used to resolve the source range or header.

```yaml
Type: String
Parameter Sets: Header
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
