---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Invoke-OfficeExcelTemplateSheet
## SYNOPSIS
Repeats an Excel template worksheet for pipeline data and applies markers in each generated sheet.

## SYNTAX
### Context (Default)
```powershell
Invoke-OfficeExcelTemplateSheet [-Item] <Object> [-TemplateSheet <string>] [-SheetNameProperty <string>] [-CultureName <string>] [-MissingValueBehavior <ExcelTemplateMissingValueBehavior>] [-ThrowOnMissing] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
Invoke-OfficeExcelTemplateSheet [-Path] <string> [-Item] <Object> [-TemplateSheet <string>] [-SheetNameProperty <string>] [-CultureName <string>] [-MissingValueBehavior <ExcelTemplateMissingValueBehavior>] [-ThrowOnMissing] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Invoke-OfficeExcelTemplateSheet [-Item] <Object> -Document <ExcelDocument> [-TemplateSheet <string>] [-SheetNameProperty <string>] [-CultureName <string>] [-MissingValueBehavior <ExcelTemplateMissingValueBehavior>] [-ThrowOnMissing] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Repeats an Excel template worksheet for pipeline data and applies markers in each generated sheet.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $invoices | Invoke-OfficeExcelTemplateSheet -Path .\Invoices.xlsx -TemplateSheet Template -SheetNameProperty SheetName
```

Uses the template sheet for the first object, copies it for later objects, binds markers, and saves the workbook.

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

### -Item
Pipeline data. Hashtables, dictionaries, PSCustomObjects, and typed objects are supported.

```yaml
Type: Object
Parameter Sets: Context, Path, Document
Aliases: Rows, Data, InputObject
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -MissingValueBehavior
Behavior used when a marker is not supplied by each input item.

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

### -SheetNameProperty
Input property used as the generated worksheet name.

```yaml
Type: String
Parameter Sets: Context, Path, Document
Aliases: NameProperty
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TemplateSheet
Template worksheet name. Defaults to the current sheet inside an ExcelSheet block or the first sheet for path/document use.

```yaml
Type: String
Parameter Sets: Context, Path, Document
Aliases: Sheet, WorksheetName
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ThrowOnMissing
Throws when a marker is not supplied by an input item.

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

- `System.Object`

## OUTPUTS

- `System.Int32`

## RELATED LINKS

- None
