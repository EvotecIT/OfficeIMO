---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficeExcelTheme
## SYNOPSIS
Applies an OfficeTheme or theme XML to a workbook.

## SYNTAX
### Name (Default)
```powershell
Set-OfficeExcelTheme -Name <string> [-Document <ExcelDocument>] [-PassThru] [<CommonParameters>]
```

### Theme
```powershell
Set-OfficeExcelTheme [-Theme] <OfficeThemeInput> [-Document <ExcelDocument>] [-Name <string>] [-PassThru] [<CommonParameters>]
```

### Xml
```powershell
Set-OfficeExcelTheme -Xml <string> [-Document <ExcelDocument>] [-Name <string>] [-PassThru] [<CommonParameters>]
```

### XmlPath
```powershell
Set-OfficeExcelTheme -XmlPath <string> [-Document <ExcelDocument>] [-Name <string>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Applies an OfficeTheme or theme XML to a workbook.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeExcel -Path .\Reports\Q3-Sales.xlsx {
    ExcelTheme Minimal
}
```

Restyles the workbook's cell styles, table styles and charts with the Minimal theme.

### EXAMPLE 2
```powershell
Open-OfficeExcel -Path .\Reports\Legacy-Report.xlsx {
    ExcelTheme .\Templates\Northwind.xltx
}
```

Takes the colours and fonts from the company template so an older workbook matches current branding.

## PARAMETERS

### -Document
Caller-owned workbook; omission uses the enclosing Excel content block.

```yaml
Type: ExcelDocument
Parameter Sets: Name, Theme, Xml, XmlPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Name
Optional workbook theme name to apply after writing the theme.

```yaml
Type: String
Parameter Sets: Name, Theme, Xml, XmlPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emits the native result after the operation.

```yaml
Type: SwitchParameter
Parameter Sets: Name, Theme, Xml, XmlPath
Aliases: None
Possible values:

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
Parameter Sets: Theme
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Xml
Theme XML to write to the workbook theme part.

```yaml
Type: String
Parameter Sets: Xml
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -XmlPath
Literal path to an XML definition.

```yaml
Type: String
Parameter Sets: XmlPath
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

- `None`

## RELATED LINKS

- None
