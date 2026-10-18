---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Export-OfficeExcel
## SYNOPSIS
Exports PowerShell objects to an Excel workbook using an operator-friendly surface.

## SYNTAX
### Create (Default)
```powershell
Export-OfficeExcel [-Path] <string> [-InputObject <Object>] [-WorksheetName <string>] [-TableName <string>] [-TableStyle <ExcelTableStyle>] [-ShowFirstColumn] [-ShowLastColumn] [-NoRowStripes] [-ShowColumnStripes] [-StartRow <int>] [-StartColumn <int>] [-NoHeader] [-NoTable] [-NoAutoFilter] [-AutoFit] [-FreezeTopRow] [-FreezeFirstColumn] [-BoldTopRow] [-Title <string>] [-NoClobber] [-ExcludeProperty <string[]>] [-CollectionSeparator <string>] [-DictionaryEntrySeparator <string>] [-DictionaryKeyValueSeparator <string>] [-MaxCollectionItems <int>] [-MaxNestingDepth <int>] [-ColumnFormat <hashtable>] [-TextColumn <string[]>] [-NumberColumn <string[]>] [-IntegerColumn <string[]>] [-PercentColumn <string[]>] [-CurrencyColumn <string[]>] [-DateColumn <string[]>] [-DateTimeColumn <string[]>] [-FormatDecimals <int>] [-FormatCultureName <string>] [-IncludeHeaderInColumnFormat] [-AutoFitFormattedColumn] [-IgnoreMissingColumnFormat] [-IncludeUnexportableProperties] [-PropertyConversionErrorAction <ActionPreference>] [-Open] [-PassThru] [-DocumentTitle <string>] [-Author <string>] [-Subject <string>] [-Keywords <string>] [-Description <string>] [-Category <string>] [-Company <string>] [-Manager <string>] [-ApplicationName <string>] [-LastModifiedBy <string>] [-SafePreflight] [-SafeRepairDefinedNames] [-ValidateOpenXml] [-DisableFastPackageWriter] [-EvaluateFormulas] [-ClearCachedFormulaResults] [-MarkFormulasDirty] [-ForceFullCalculationOnOpen] [-DateSystem <string>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Append
```powershell
Export-OfficeExcel [-Path] <string> -Append [-InputObject <Object>] [-WorksheetName <string>] [-TableName <string>] [-TableStyle <ExcelTableStyle>] [-ShowFirstColumn] [-ShowLastColumn] [-NoRowStripes] [-ShowColumnStripes] [-StartRow <int>] [-StartColumn <int>] [-NoHeader] [-NoTable] [-NoAutoFilter] [-AutoFit] [-FreezeTopRow] [-FreezeFirstColumn] [-BoldTopRow] [-Title <string>] [-AppendToTable] [-NoClobber] [-ExcludeProperty <string[]>] [-CollectionSeparator <string>] [-DictionaryEntrySeparator <string>] [-DictionaryKeyValueSeparator <string>] [-MaxCollectionItems <int>] [-MaxNestingDepth <int>] [-ColumnFormat <hashtable>] [-TextColumn <string[]>] [-NumberColumn <string[]>] [-IntegerColumn <string[]>] [-PercentColumn <string[]>] [-CurrencyColumn <string[]>] [-DateColumn <string[]>] [-DateTimeColumn <string[]>] [-FormatDecimals <int>] [-FormatCultureName <string>] [-IncludeHeaderInColumnFormat] [-AutoFitFormattedColumn] [-IgnoreMissingColumnFormat] [-IncludeUnexportableProperties] [-PropertyConversionErrorAction <ActionPreference>] [-Open] [-PassThru] [-DocumentTitle <string>] [-Author <string>] [-Subject <string>] [-Keywords <string>] [-Description <string>] [-Category <string>] [-Company <string>] [-Manager <string>] [-ApplicationName <string>] [-LastModifiedBy <string>] [-SafePreflight] [-SafeRepairDefinedNames] [-ValidateOpenXml] [-DisableFastPackageWriter] [-EvaluateFormulas] [-ClearCachedFormulaResults] [-MarkFormulasDirty] [-ForceFullCalculationOnOpen] [-DateSystem <string>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### ClearSheet
```powershell
Export-OfficeExcel [-Path] <string> -ClearSheet [-InputObject <Object>] [-WorksheetName <string>] [-TableName <string>] [-TableStyle <ExcelTableStyle>] [-ShowFirstColumn] [-ShowLastColumn] [-NoRowStripes] [-ShowColumnStripes] [-StartRow <int>] [-StartColumn <int>] [-NoHeader] [-NoTable] [-NoAutoFilter] [-AutoFit] [-FreezeTopRow] [-FreezeFirstColumn] [-BoldTopRow] [-Title <string>] [-NoClobber] [-ExcludeProperty <string[]>] [-CollectionSeparator <string>] [-DictionaryEntrySeparator <string>] [-DictionaryKeyValueSeparator <string>] [-MaxCollectionItems <int>] [-MaxNestingDepth <int>] [-ColumnFormat <hashtable>] [-TextColumn <string[]>] [-NumberColumn <string[]>] [-IntegerColumn <string[]>] [-PercentColumn <string[]>] [-CurrencyColumn <string[]>] [-DateColumn <string[]>] [-DateTimeColumn <string[]>] [-FormatDecimals <int>] [-FormatCultureName <string>] [-IncludeHeaderInColumnFormat] [-AutoFitFormattedColumn] [-IgnoreMissingColumnFormat] [-IncludeUnexportableProperties] [-PropertyConversionErrorAction <ActionPreference>] [-Open] [-PassThru] [-DocumentTitle <string>] [-Author <string>] [-Subject <string>] [-Keywords <string>] [-Description <string>] [-Category <string>] [-Company <string>] [-Manager <string>] [-ApplicationName <string>] [-LastModifiedBy <string>] [-SafePreflight] [-SafeRepairDefinedNames] [-ValidateOpenXml] [-DisableFastPackageWriter] [-EvaluateFormulas] [-ClearCachedFormulaResults] [-MarkFormulasDirty] [-ForceFullCalculationOnOpen] [-DateSystem <string>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Provides a fast PowerShell export path while keeping OfficeIMO as the workbook engine.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $rows | Export-OfficeExcel -Path .\Report.xlsx -WorksheetName Data -TableName Data -AutoFit -FreezeTopRow
```

Creates a workbook, writes the objects as a table, auto-fits columns, and freezes the header row.

### EXAMPLE 2
```powershell
PS> $rows | Export-OfficeExcel -Path .\Report.xlsx -WorksheetName Data -TableName Sales -TextColumn Id -CurrencyColumn Revenue -ColumnFormat @{ Rate = @{ Style = 'Percent'; Decimals = 1 }; Created = 'Date' } -FormatCultureName en-US -AutoFitFormattedColumn
```

Formats ID values as text, Revenue as currency, Rate as a one-decimal percentage, and Created as a short date while keeping formatting logic in OfficeIMO.

## PARAMETERS

### -Append
Append rows to an existing worksheet when the workbook exists.

```yaml
Type: SwitchParameter
Parameter Sets: Append
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AppendToTable
Require append operations to extend an existing Excel table instead of writing after the used range.

```yaml
Type: SwitchParameter
Parameter Sets: Append
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ApplicationName
Workbook application-name metadata.

```yaml
Type: String
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Author
Workbook author metadata.

```yaml
Type: String
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AutoFit
Auto-fit exported columns.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: AutoSize
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AutoFitFormattedColumn
Auto-fit only columns that receive export-time column formats.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BoldTopRow
Bold the exported header row.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Category
Workbook category metadata.

```yaml
Type: String
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ClearCachedFormulaResults
Remove cached formula results before saving.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ClearSheet
Replace the target worksheet inside an existing workbook.

```yaml
Type: SwitchParameter
Parameter Sets: ClearSheet
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CollectionSeparator
Text used between items when a cell contains a collection.

```yaml
Type: String
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ColumnFormat
Header-to-format map. Values may be preset names such as Text, Currency, Percent, Date, or custom Excel number formats.

```yaml
Type: Hashtable
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Company
Workbook company metadata.

```yaml
Type: String
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CurrencyColumn
Headers that should be formatted as currency.

```yaml
Type: String[]
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DateColumn
Headers that should be formatted as dates.

```yaml
Type: String[]
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DateSystem
Workbook date system for Excel date serials.

```yaml
Type: String
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values: 1900, 1904, NineteenHundred, NineteenFour

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DateTimeColumn
Headers that should be formatted as date/time values.

```yaml
Type: String[]
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Description
Workbook description metadata.

```yaml
Type: String
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DictionaryEntrySeparator
Text used between entries when a cell contains a dictionary.

```yaml
Type: String
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DictionaryKeyValueSeparator
Text used between a dictionary key and value.

```yaml
Type: String
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DisableFastPackageWriter
Disable OfficeIMO fast package writers for this save.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DocumentTitle
Workbook document title metadata.

```yaml
Type: String
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -EvaluateFormulas
Evaluate supported formulas and write cached values before saving.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExcludeProperty
Exclude specific properties from exported objects.

```yaml
Type: String[]
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ForceFullCalculationOnOpen
Request a full workbook recalculation when opened in Excel-compatible applications.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FormatCultureName
Culture used by friendly currency column presets, such as en-US or pl-PL.

```yaml
Type: String
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FormatDecimals
Decimal places used by friendly number, percent, and currency column presets.

```yaml
Type: Int32
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FreezeFirstColumn
Freeze the first exported column.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FreezeTopRow
Freeze the exported header row.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IgnoreMissingColumnFormat
Continue when a requested export-time column format header is missing.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeHeaderInColumnFormat
Include header cells when applying export-time column formats.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeUnexportableProperties
Include properties that cannot be read by exporting a descriptive placeholder value.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputObject
Objects to write. Accepts pipeline input.

```yaml
Type: Object
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -IntegerColumn
Headers that should be formatted as whole numbers.

```yaml
Type: String[]
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Keywords
Workbook keyword metadata.

```yaml
Type: String
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LastModifiedBy
Workbook last-modified-by metadata.

```yaml
Type: String
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Manager
Workbook manager metadata.

```yaml
Type: String
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MarkFormulasDirty
Mark formula cells dirty before saving.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxCollectionItems
Maximum number of items allowed in one nested collection or dictionary cell. Defaults to 1,048,575; increase explicitly for trusted larger values.

```yaml
Type: Int32
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxNestingDepth
Maximum nesting depth allowed while normalizing one cell value. Defaults to 64; increase explicitly for trusted deeper values.

```yaml
Type: Int32
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoAutoFilter
Disable AutoFilter dropdowns on the created table.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoClobber
Do not overwrite an existing workbook unless appending or clearing a sheet.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoHeader
Do not emit a header row.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoRowStripes
Disable alternating row stripes for newly created tables.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoTable
Do not create an Excel table around the exported data.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NumberColumn
Headers that should be formatted as decimal numbers.

```yaml
Type: String[]
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Open
Open the workbook after saving.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: Show
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the saved FileInfo.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Destination workbook path.

```yaml
Type: String
Parameter Sets: Create, Append, ClearSheet
Aliases: FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PercentColumn
Headers that should be formatted as percentages.

```yaml
Type: String[]
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PropertyConversionErrorAction
Controls how unreadable PowerShell properties are handled while projecting export rows.

```yaml
Type: ActionPreference
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values: SilentlyContinue, Stop, Continue, Inquire, Ignore, Suspend, Break

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SafePreflight
Run OfficeIMO worksheet preflight cleanup before saving.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SafeRepairDefinedNames
Repair common defined-name issues before saving.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ShowColumnStripes
Enable alternating column stripes for newly created tables.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ShowFirstColumn
Emphasize the first table column when the selected style supports it.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ShowLastColumn
Emphasize the last table column when the selected style supports it.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StartColumn
Starting column for new exports.

```yaml
Type: Int32
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StartRow
Starting row for new exports. When appending and left at 1, rows are written after the used range.

```yaml
Type: Int32
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Subject
Workbook subject metadata.

```yaml
Type: String
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TableName
Optional Excel table name.

```yaml
Type: String
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TableStyle
Built-in Excel table style name.

```yaml
Type: ExcelTableStyle
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values: TableStyleLight1, TableStyleLight2, TableStyleLight3, TableStyleLight4, TableStyleLight5, TableStyleLight6, TableStyleLight7, TableStyleLight8, TableStyleLight9, TableStyleLight10, TableStyleLight11, TableStyleLight12, TableStyleLight13, TableStyleLight14, TableStyleLight15, TableStyleLight16, TableStyleLight17, TableStyleLight18, TableStyleLight19, TableStyleLight20, TableStyleLight21, TableStyleMedium1, TableStyleMedium2, TableStyleMedium3, TableStyleMedium4, TableStyleMedium5, TableStyleMedium6, TableStyleMedium7, TableStyleMedium8, TableStyleMedium9, TableStyleMedium10, TableStyleMedium11, TableStyleMedium12, TableStyleMedium13, TableStyleMedium14, TableStyleMedium15, TableStyleMedium16, TableStyleMedium17, TableStyleMedium18, TableStyleMedium19, TableStyleMedium20, TableStyleMedium21, TableStyleMedium22, TableStyleMedium23, TableStyleMedium24, TableStyleMedium25, TableStyleMedium26, TableStyleMedium27, TableStyleMedium28, TableStyleDark1, TableStyleDark2, TableStyleDark3, TableStyleDark4, TableStyleDark5, TableStyleDark6, TableStyleDark7, TableStyleDark8, TableStyleDark9, TableStyleDark10, TableStyleDark11

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TextColumn
Headers that should be formatted as text, useful for IDs, zip codes, and leading-zero values.

```yaml
Type: String[]
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Title
Write a title above the exported table.

```yaml
Type: String
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ValidateOpenXml
Validate the saved package with OpenXmlValidator and throw on errors.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WorksheetName
Worksheet name to create or update.

```yaml
Type: String
Parameter Sets: Create, Append, ClearSheet
Aliases: Sheet
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

- `None`

## RELATED LINKS

- None
