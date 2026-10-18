---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Import-OfficeExcelDelimitedText
## SYNOPSIS
Imports normalized CSV/TSV text into an Excel workbook through OfficeIMO.

## SYNTAX
### Path (Default)
```powershell
Import-OfficeExcelDelimitedText [-Path] <string> -SourcePath <string> [-Delimiter <Char>] [-SheetName <string>] [-CultureName <string>] [-NoHeader] [-SkipRows <int>] [-NoTable] [-NoTypeConversion] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Import-OfficeExcelDelimitedText -Document <ExcelDocument> -SourcePath <string> [-Delimiter <Char>] [-SheetName <string>] [-CultureName <string>] [-NoHeader] [-SkipRows <int>] [-NoTable] [-NoTypeConversion] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Imports normalized CSV/TSV text into an Excel workbook through OfficeIMO.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $result = Import-OfficeExcelDelimitedText -Path .\Report.xlsx `
                -SourcePath .\sales-pl.csv `
                -Delimiter ';' `
                -CultureName 'pl-PL' `
                -SheetName Sales `
                -PassThru
            $result | Format-List SheetName,Range,RowCount,ColumnCount,Delimiter
```

Normalizes delimited text through OfficeIMO, performs culture-aware value conversion, and writes the result as an Excel table unless -NoTable is used.

## PARAMETERS

### -CultureName
Culture name for number and date conversion.

```yaml
Type: String
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Delimiter
Delimiter character. When omitted, it is detected.

```yaml
Type: Char
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Workbook document.

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

### -NoHeader
Treat the first row as data.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoTable
Import rows without creating an Excel table.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoTypeConversion
Keep imported values as text.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit a result object.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Workbook path.

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

### -SheetName
Worksheet name to create or inspect.

```yaml
Type: String
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SkipRows
Number of parsed delimited records to skip before header discovery or data import.

```yaml
Type: Int32
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SourcePath
Delimited text source path.

```yaml
Type: String
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: True
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
