---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeExcelMergeOptions
## SYNOPSIS
Creates native import policies and typed worksheet selections for Merge-OfficeDocument.

## SYNTAX
### Default (Default)
```powershell
New-OfficeExcelMergeOptions [-Sheet <OfficeSheetSelector[]>] [-NamePrefix <string>] [-ValidationMode <ExcelSheetNameValidationMode>] [-CopyMode <ExcelWorksheetCopyMode>] [<CommonParameters>]
```

## DESCRIPTION
The first input or live Document remains the base. Sheet selects appended source worksheets;
sheet objects identify worksheets in saved source files and remain caller-owned.

## EXAMPLES

### EXAMPLE 1
```powershell
$options = New-OfficeExcelMergeOptions -Sheet 'Sales*', 3 -NamePrefix 'Imported ' -CopyMode Package
Merge-OfficeDocument -Path ./Cover.xlsx, ./North.xlsx, ./South.xlsx -OutputPath ./Combined.xlsx -Options $options
```


## PARAMETERS

### -CopyMode
Native value or package copy policy.

```yaml
Type: ExcelWorksheetCopyMode
Parameter Sets: Default
Aliases: None
Possible values: Values, Package

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NamePrefix
Prefix added to imported worksheet names.

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

### -Sheet
Source worksheet names/wildcards, one-based positions or sheet objects from saved appended sources.

```yaml
Type: OfficeSheetSelector[]
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ValidationMode
Native imported-sheet name validation policy.

```yaml
Type: ExcelSheetNameValidationMode
Parameter Sets: Default
Aliases: None
Possible values: None, Sanitize, Strict

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

- `PSWriteOffice.Models.Excel.ExcelMergeOptions`: Combines native workbook import policies with PowerShell worksheet selectors.

## RELATED LINKS

- None
