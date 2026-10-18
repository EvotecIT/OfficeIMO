---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertTo-OfficeExcelHtml
## SYNOPSIS
Converts an Excel workbook to an HTML review document.

## SYNTAX
### Path (Default)
```powershell
ConvertTo-OfficeExcelHtml [-Path] <string> [-Password <string>] [-OutputPath <string>] [-Profile <OfficeExcelHtmlProfile>] [-Theme <OfficeVisualThemeKind>] [-Title <string>] [-MaxRowsPerSheet <Int32>] [-EmptyCellText <string>] [-NoDefaultStyles] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Workbook
```powershell
ConvertTo-OfficeExcelHtml -Workbook <ExcelDocument> [-OutputPath <string>] [-Profile <OfficeExcelHtmlProfile>] [-Theme <OfficeVisualThemeKind>] [-Title <string>] [-MaxRowsPerSheet <Int32>] [-EmptyCellText <string>] [-NoDefaultStyles] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Converts an Excel workbook to an HTML review document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ConvertTo-OfficeExcelHtml -Path .\Report.xlsx -OutputPath .\Report.html -Title 'Workbook Review' -PassThru
```

Loads the workbook and writes an HTML file with tables, formulas, comments, charts, and image inventory where available.

### EXAMPLE 2
```powershell
PS> ConvertTo-OfficeExcelHtml -Path .\Report.xlsx -Profile VisualReview -OutputPath .\Report.visual.html
```

Uses the OfficeIMO Excel visual review profile.

## PARAMETERS

### -EmptyCellText
Text used for empty cells in semantic table output.

```yaml
Type: String
Parameter Sets: Path, Workbook
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxRowsPerSheet
Maximum number of used rows to emit per worksheet.

```yaml
Type: Int32
Parameter Sets: Path, Workbook
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoDefaultStyles
Do not include OfficeIMO default CSS styles.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Workbook
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputPath
Optional output HTML path. When omitted, HTML text is returned.

```yaml
Type: String
Parameter Sets: Path, Workbook
Aliases: OutPath
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit a FileInfo when saving to disk.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Workbook
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Password used to open an encrypted workbook package.

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

### -Path
Path to the workbook to convert.

```yaml
Type: String
Parameter Sets: Path
Aliases: FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Profile
HTML conversion profile.

```yaml
Type: OfficeExcelHtmlProfile
Parameter Sets: Path, Workbook
Aliases: None
Possible values: SemanticTables, VisualReview

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Theme
Built-in HTML document theme.

```yaml
Type: OfficeVisualThemeKind
Parameter Sets: Path, Workbook
Aliases: None
Possible values: Plain, WordLike, TechnicalDocument, GitHubLike, Compact, Report

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Title
Optional HTML document title.

```yaml
Type: String
Parameter Sets: Path, Workbook
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Workbook
Workbook instance to convert.

```yaml
Type: ExcelDocument
Parameter Sets: Workbook
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String`
- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `System.String`
- `System.IO.FileInfo`

## RELATED LINKS

- None
