---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Invoke-OfficeExcelTemplateOptionalRow
## SYNOPSIS
Includes or removes an optional Excel template row block.

## SYNTAX
### Context (Default)
```powershell
Invoke-OfficeExcelTemplateOptionalRow -FirstRow <int> [-Sheet <string>] [-SheetIndex <Int32>] [-RowCount <int>] [-Value <hashtable>] [-Remove] [-CultureName <string>] [-MissingValueBehavior <ExcelTemplateMissingValueBehavior>] [-ThrowOnMissing] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
Invoke-OfficeExcelTemplateOptionalRow [-Path] <string> -FirstRow <int> [-Sheet <string>] [-SheetIndex <Int32>] [-RowCount <int>] [-Value <hashtable>] [-Remove] [-CultureName <string>] [-MissingValueBehavior <ExcelTemplateMissingValueBehavior>] [-ThrowOnMissing] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Invoke-OfficeExcelTemplateOptionalRow -Document <ExcelDocument> -FirstRow <int> [-Sheet <string>] [-SheetIndex <Int32>] [-RowCount <int>] [-Value <hashtable>] [-Remove] [-CultureName <string>] [-MissingValueBehavior <ExcelTemplateMissingValueBehavior>] [-ThrowOnMissing] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Includes or removes an optional Excel template row block.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Invoke-OfficeExcelTemplateOptionalRow -Path .\Invoice.xlsx -Sheet Invoice -FirstRow 10 -Value @{ Discount = '10%' }
```

Leaves the optional row block in place, replaces its markers, and saves the workbook.

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
Accept pipeline input: False
Accept wildcard characters: False
```

### -FirstRow
1-based first row in the optional block.

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

### -MissingValueBehavior
Behavior used when a marker in the optional block is not supplied by -Value.

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

### -Remove
Removes the optional row block instead of keeping and binding it.

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

### -RowCount
Number of rows in the optional block.

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

### -ThrowOnMissing
Throws when a marker in the optional block is not supplied by -Value.

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

### -Value
Template marker values used when the optional row block is included.

```yaml
Type: Hashtable
Parameter Sets: Context, Path, Document
Aliases: Values
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

- `System.Int32`

## RELATED LINKS

- None
