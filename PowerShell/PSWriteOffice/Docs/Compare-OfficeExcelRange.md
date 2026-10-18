---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Compare-OfficeExcelRange
## SYNOPSIS
Compares two Excel worksheets or ranges and returns cell-level differences.

## SYNTAX
### Path (Default)
```powershell
Compare-OfficeExcelRange [-Path] <string> [-RightPath <string>] [-LeftSheet <string>] [-LeftSheetIndex <Int32>] [-RightSheet <string>] [-RightSheetIndex <Int32>] [-LeftRange <string>] [-RightRange <string>] [-TrimStrings] [-IgnoreCase] [-StrictNullEmpty] [<CommonParameters>]
```

### Document
```powershell
Compare-OfficeExcelRange -Document <ExcelDocument> [-RightDocument <ExcelDocument>] [-LeftSheet <string>] [-LeftSheetIndex <Int32>] [-RightSheet <string>] [-RightSheetIndex <Int32>] [-LeftRange <string>] [-RightRange <string>] [-TrimStrings] [-IgnoreCase] [-StrictNullEmpty] [<CommonParameters>]
```

### Context
```powershell
Compare-OfficeExcelRange [-RightDocument <ExcelDocument>] [-LeftSheet <string>] [-LeftSheetIndex <Int32>] [-RightSheet <string>] [-RightSheetIndex <Int32>] [-LeftRange <string>] [-RightRange <string>] [-TrimStrings] [-IgnoreCase] [-StrictNullEmpty] [<CommonParameters>]
```

## DESCRIPTION
Compares two Excel worksheets or ranges and returns cell-level differences.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $differences = Compare-OfficeExcelRange -Path .\Report.xlsx -LeftSheet Current -RightSheet Expected -LeftRange A1:D20 -RightRange A1:D20 -TrimStrings -IgnoreCase
if ($differences) {
    $differences | Export-Csv -Path .\RangeDifferences.csv -NoTypeInformation
}
```

Compares two ranges and exports cell-level differences when the workbook does not match the expected sheet.

## PARAMETERS

### -Document
Left workbook object.

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

### -IgnoreCase
Compare strings case-insensitively.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document, Context
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LeftRange
Left A1 range. Defaults to the left worksheet used range.

```yaml
Type: String
Parameter Sets: Path, Document, Context
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LeftSheet
Left worksheet name. Defaults to the current sheet in a sheet block, otherwise the first sheet.

```yaml
Type: String
Parameter Sets: Path, Document, Context
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LeftSheetIndex
Left worksheet index.

```yaml
Type: Int32
Parameter Sets: Path, Document, Context
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Left workbook path.

```yaml
Type: String
Parameter Sets: Path
Aliases: InputPath, FilePath, LeftPath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RightDocument
Optional right workbook object. Defaults to the left workbook.

```yaml
Type: ExcelDocument
Parameter Sets: Document, Context
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RightPath
Optional right workbook path. Defaults to the left workbook.

```yaml
Type: String
Parameter Sets: Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RightRange
Right A1 range. Defaults to the right worksheet used range.

```yaml
Type: String
Parameter Sets: Path, Document, Context
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RightSheet
Right worksheet name. Defaults to the left sheet.

```yaml
Type: String
Parameter Sets: Path, Document, Context
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RightSheetIndex
Right worksheet index.

```yaml
Type: Int32
Parameter Sets: Path, Document, Context
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StrictNullEmpty
Treat null and empty string values as different.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document, Context
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TrimStrings
Compare strings after trimming whitespace.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document, Context
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

- `OfficeIMO.Excel.ExcelRangeDifference`

## RELATED LINKS

- None
