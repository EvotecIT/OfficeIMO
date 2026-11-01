---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeExcelGoogleSheetsOptions
## SYNOPSIS
Creates native Excel-to-Google-Sheets planning and export settings without contacting Google.

## SYNTAX
### Default (Default)
```powershell
New-OfficeExcelGoogleSheetsOptions [-Title <string>] [-FolderId <string>] [-DriveId <string>] [-SpreadsheetId <string>] [-ExpectedDriveVersion <Int64>] [-UnsupportedFormula <GoogleSheetsUnsupportedFormulaMode>] [-Charts <UnsupportedFeatureMode>] [-PivotTables <UnsupportedFeatureMode>] [-PrintLayout <UnsupportedFeatureMode>] [-Locale <string>] [-TimeZone <string>] [-RecalculationInterval <GoogleSheetsRecalculationInterval>] [-FunctionMapping <hashtable>] [-Editor <string[]>] [-DomainUsersCanEdit] [<CommonParameters>]
```

## DESCRIPTION
Creates native Excel-to-Google-Sheets planning and export settings without contacting Google.

## EXAMPLES

### EXAMPLE 1
```powershell
$options = New-OfficeExcelGoogleSheetsOptions -Title 'Quarterly forecast' -UnsupportedFormula PreserveWithWarning
Export-OfficeExcelGoogleSpreadsheet -Document $workbook -Options $options -PlanOnly
```


## PARAMETERS

### -Charts
Policy for charts without a reliable native equivalent. Omission retains the native default.

```yaml
Type: UnsupportedFeatureMode
Parameter Sets: Default
Aliases: None
Possible values: Error, WarnAndSkip, Flatten, Rasterize

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DomainUsersCanEdit
Whether domain users may edit protected content. Omission preserves the native policy.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DriveId
Shared-drive identifier. Omission retains the native default.

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

### -Editor
Editor email addresses for translated protected ranges.

```yaml
Type: String[]
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExpectedDriveVersion
Drive version observed before replacement; the native exporter checks for changes. Omission retains the native default.

```yaml
Type: Int64
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FolderId
Drive destination folder identifier. Omission retains the native default.

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

### -FunctionMapping
Function-name replacements; keys and values are native formula function names.

```yaml
Type: Hashtable
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Locale
Target spreadsheet locale. Omission retains the native default.

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

### -PivotTables
Policy for unsupported pivot tables. Omission retains the native default.

```yaml
Type: UnsupportedFeatureMode
Parameter Sets: Default
Aliases: None
Possible values: Error, WarnAndSkip, Flatten, Rasterize

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PrintLayout
Policy for print layout without an equivalent. Omission retains the native default.

```yaml
Type: UnsupportedFeatureMode
Parameter Sets: Default
Aliases: None
Possible values: Error, WarnAndSkip, Flatten, Rasterize

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RecalculationInterval
When the target recalculates formulas. Omission retains the native default.

```yaml
Type: GoogleSheetsRecalculationInterval
Parameter Sets: Default
Aliases: None
Possible values: OnChange, Minute, Hour

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SpreadsheetId
Existing spreadsheet identifier for replacement. Omission retains the native default.

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

### -TimeZone
Target spreadsheet time zone. Omission retains the native default.

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

### -Title
Title of a new spreadsheet. Omission retains the native default.

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

### -UnsupportedFormula
Policy for formulas that cannot be translated. Omission retains the native default.

```yaml
Type: GoogleSheetsUnsupportedFormulaMode
Parameter Sets: Default
Aliases: None
Possible values: Error, PreserveWithWarning, UseCachedValue

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

- `OfficeIMO.Excel.GoogleSheets.GoogleSheetsSaveOptions`

## RELATED LINKS

- None
