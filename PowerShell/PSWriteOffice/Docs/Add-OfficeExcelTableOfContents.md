---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeExcelTableOfContents
## SYNOPSIS
Adds or refreshes a workbook table of contents sheet.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeExcelTableOfContents [-SheetName <string>] [-DoNotPlaceFirst] [-NoHyperlinks] [-IncludeNamedRanges] [-IncludeHiddenNamedRanges] [-NoStyle] [-AddBackLinks] [-BackLinkRow <int>] [-BackLinkColumn <int>] [-BackLinkText <string>] [-Open] [-PassThru] [<CommonParameters>]
```

### Path
```powershell
Add-OfficeExcelTableOfContents [-Path] <string> [-SheetName <string>] [-DoNotPlaceFirst] [-NoHyperlinks] [-IncludeNamedRanges] [-IncludeHiddenNamedRanges] [-NoStyle] [-AddBackLinks] [-BackLinkRow <int>] [-BackLinkColumn <int>] [-BackLinkText <string>] [-Open] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeExcelTableOfContents -Document <ExcelDocument> [-SheetName <string>] [-DoNotPlaceFirst] [-NoHyperlinks] [-IncludeNamedRanges] [-IncludeHiddenNamedRanges] [-NoStyle] [-AddBackLinks] [-BackLinkRow <int>] [-BackLinkColumn <int>] [-BackLinkText <string>] [-Open] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Can run inside the Excel DSL, against an open workbook, or directly against a file path.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $proof = @(
    Add-OfficeExcelTableOfContents -Path .\report.xlsx -IncludeNamedRanges -AddBackLinks -PassThru
    Get-OfficeExcelSummary -Path .\report.xlsx -IncludeSheets |
        Select-Object -Property SheetCount, NamedRangeCount, Sheets
)
$proof
```

Creates or refreshes a TOC sheet, adds back links, and reads back workbook navigation metadata.

## PARAMETERS

### -AddBackLinks
Add a quick link back to the TOC on each worksheet.

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

### -BackLinkColumn
Column for the back link when AddBackLinks is used.

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

### -BackLinkRow
Row for the back link when AddBackLinks is used.

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

### -BackLinkText
Text used for back links.

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
Workbook to update.

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

### -DoNotPlaceFirst
Keep the TOC sheet in its current position instead of moving it first.

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

### -IncludeHiddenNamedRanges
Include hidden named ranges when listing named ranges.

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

### -IncludeNamedRanges
Include named ranges in the TOC.

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

### -NoHyperlinks
Disable internal hyperlinks in the TOC sheet.

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

### -NoStyle
Disable formatted TOC styling.

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

### -Open
Open the workbook after saving when Path is used.

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

### -PassThru
Emit the updated document or file info.

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
Path to the workbook to update in place.

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
Name of the TOC sheet.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `OfficeIMO.Excel.ExcelDocument`
- `System.IO.FileInfo`

## RELATED LINKS

- None
