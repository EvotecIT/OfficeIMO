---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeExcelDataSet
## SYNOPSIS
Writes every table in a DataSet to separate Excel worksheets.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeExcelDataSet [-DataSet] <DataSet> [-NoTable] [-NoHeader] [-TableStyle <ExcelTableStyle>] [-ShowFirstColumn] [-ShowLastColumn] [-NoRowStripes] [-ShowColumnStripes] [-NoAutoFilter] [-AutoFit] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Uses OfficeIMO.Excel DataSet ingestion so callers can provide data from any .NET provider without PSWriteOffice owning database connections.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficeExcel -Path .\report.xlsx { Add-OfficeExcelDataSet -DataSet $dataSet -AutoFit }
```

Creates one worksheet per DataTable and formats each range as an Excel table.

## PARAMETERS

### -AutoFit
Auto-fit imported table columns.

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

### -DataSet
Source DataSet whose tables will become worksheets.

```yaml
Type: DataSet
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -NoAutoFilter
Disable AutoFilter dropdowns.

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

### -NoHeader
Skip writing headers.

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

### -NoRowStripes
Disable alternating row stripes for created tables.

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

### -NoTable
Write plain ranges instead of Excel tables.

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

### -PassThru
Return import metadata for each worksheet.

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

### -ShowColumnStripes
Enable alternating column stripes for created tables.

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

### -ShowFirstColumn
Emphasize the first table column when the selected style supports it.

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

### -ShowLastColumn
Emphasize the last table column when the selected style supports it.

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

### -TableStyle
Built-in table style to apply.

```yaml
Type: ExcelTableStyle
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: TableStyleLight1, TableStyleLight2, TableStyleLight3, TableStyleLight4, TableStyleLight5, TableStyleLight6, TableStyleLight7, TableStyleLight8, TableStyleLight9, TableStyleLight10, TableStyleLight11, TableStyleLight12, TableStyleLight13, TableStyleLight14, TableStyleLight15, TableStyleLight16, TableStyleLight17, TableStyleLight18, TableStyleLight19, TableStyleLight20, TableStyleLight21, TableStyleMedium1, TableStyleMedium2, TableStyleMedium3, TableStyleMedium4, TableStyleMedium5, TableStyleMedium6, TableStyleMedium7, TableStyleMedium8, TableStyleMedium9, TableStyleMedium10, TableStyleMedium11, TableStyleMedium12, TableStyleMedium13, TableStyleMedium14, TableStyleMedium15, TableStyleMedium16, TableStyleMedium17, TableStyleMedium18, TableStyleMedium19, TableStyleMedium20, TableStyleMedium21, TableStyleMedium22, TableStyleMedium23, TableStyleMedium24, TableStyleMedium25, TableStyleMedium26, TableStyleMedium27, TableStyleMedium28, TableStyleDark1, TableStyleDark2, TableStyleDark3, TableStyleDark4, TableStyleDark5, TableStyleDark6, TableStyleDark7, TableStyleDark8, TableStyleDark9, TableStyleDark10, TableStyleDark11

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.Data.DataSet`

## OUTPUTS

- `OfficeIMO.Excel.ExcelDataSetImportResult`

## RELATED LINKS

- None
