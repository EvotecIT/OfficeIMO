---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Invoke-OfficeExcelTemplateRow
## SYNOPSIS
Repeats an Excel template row for pipeline data and replaces markers in each inserted row.

## SYNTAX
### Context (Default)
```powershell
Invoke-OfficeExcelTemplateRow [-InputObject] <Object> -TemplateRow <int> [-Sheet <string>] [-SheetIndex <Int32>] [-CultureName <string>] [-MissingValueBehavior <ExcelTemplateMissingValueBehavior>] [-ThrowOnMissing] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
Invoke-OfficeExcelTemplateRow [-Path] <string> [-InputObject] <Object> -TemplateRow <int> [-Sheet <string>] [-SheetIndex <Int32>] [-CultureName <string>] [-MissingValueBehavior <ExcelTemplateMissingValueBehavior>] [-ThrowOnMissing] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Invoke-OfficeExcelTemplateRow [-InputObject] <Object> -Document <ExcelDocument> -TemplateRow <int> [-Sheet <string>] [-SheetIndex <Int32>] [-CultureName <string>] [-MissingValueBehavior <ExcelTemplateMissingValueBehavior>] [-ThrowOnMissing] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Repeats an Excel template row for pipeline data and replaces markers in each inserted row.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $items | Invoke-OfficeExcelTemplateRow -Path .\Invoice.xlsx -Sheet Invoice -TemplateRow 12 -CultureName en-US
```

Copies the template row once per input object, applies marker values, and saves the workbook.

## PARAMETERS

### -CultureName
Culture name used for built-in marker format aliases such as currency and date.

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

### -InputObject
Pipeline data. Hashtables, dictionaries, PSCustomObjects, and typed objects are supported.

```yaml
Type: Object
Parameter Sets: Context, Path, Document
Aliases: Rows, Data
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -MissingValueBehavior
Behavior used when a marker is not supplied by each input row.

```yaml
Type: ExcelTemplateMissingValueBehavior
Parameter Sets: Context, Path, Document
Aliases: None
Possible values: PreserveMarker, EmptyString, Throw

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Returns the number of marker replacements.

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

### -Sheet
Worksheet name. Defaults to the current sheet inside an ExcelSheet block.

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
Worksheet index when using a workbook object or path.

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

### -TemplateRow
1-based row number that contains template markers to repeat.

```yaml
Type: Int32
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ThrowOnMissing
Throws when a marker is not supplied by an input row.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelDocument`
- `System.Object`

## OUTPUTS

- `System.Int32`

## RELATED LINKS

- None
