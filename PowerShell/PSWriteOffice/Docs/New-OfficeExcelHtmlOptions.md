---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeExcelHtmlOptions
## SYNOPSIS
Creates Excel-to-HTML settings for Convert-OfficeDocument -To Html.

## SYNTAX
### Default (Default)
```powershell
New-OfficeExcelHtmlOptions [-Profile <OfficeExcelHtmlProfile>] [-Theme <OfficeThemeInput>] [-Title <string>] [-MaxRowsPerSheet <Int32>] [-EmptyCellText <string>] [-NoDefaultStyles] [<CommonParameters>]
```

## DESCRIPTION
Creates Excel-to-HTML settings for Convert-OfficeDocument -To Html.

## EXAMPLES

### EXAMPLE 1
```powershell
Convert-OfficeDocument -Path .\Reports\Q3-Sales.xlsx -To Html -OutputPath .\Site\Q3-Sales.html -Options (New-OfficeExcelHtmlOptions -Title 'Q3 sales')
```

Converts a workbook to a web page with a title.

### EXAMPLE 2
```powershell
$html = New-OfficeExcelHtmlOptions -Profile VisualReview -Theme Boardroom -Title 'Operations dashboard' -MaxRowsPerSheet 500
Convert-OfficeDocument -Path .\Reports\Operations.xlsx -To Html -OutputPath .\Site\Operations.html -Options $html
```

Renders the workbook as a visual review page in the Boardroom theme, capped at 500 rows per sheet for a fast page.

## PARAMETERS

### -EmptyCellText
Text used for empty cells in semantic table output.

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

### -MaxRowsPerSheet
Maximum number of used rows to emit per worksheet.

```yaml
Type: Int32
Parameter Sets: Default
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
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Profile
HTML conversion profile.

```yaml
Type: OfficeExcelHtmlProfile
Parameter Sets: Default
Aliases: None
Possible values: SemanticTables, VisualReview

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Theme
Built-in Office theme, native theme object or a document whose visual theme is read.

```yaml
Type: OfficeThemeInput
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
Chart title.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `None`

## RELATED LINKS

- None
