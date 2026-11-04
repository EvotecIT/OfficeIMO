---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeExcelDocumentProperty
## SYNOPSIS
Gets built-in, application and custom document properties.

## SYNTAX
### Document (Default)
```powershell
Get-OfficeExcelDocumentProperty [-Name <string[]>] [-Kind <ExcelDocumentPropertyKind[]>] [-Document <ExcelDocument>] [<CommonParameters>]
```

### Path
```powershell
Get-OfficeExcelDocumentProperty [-Path] <string[]> [-Name <string[]>] [-Kind <ExcelDocumentPropertyKind[]>] [<CommonParameters>]
```

### LiteralPath
```powershell
Get-OfficeExcelDocumentProperty -LiteralPath <string[]> [-Name <string[]>] [-Kind <ExcelDocumentPropertyKind[]>] [<CommonParameters>]
```

## DESCRIPTION
Gets built-in, application and custom document properties.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-OfficeExcelDocumentProperty -Path .\Reports\Q3-Sales.xlsx
```

Lists the built-in, application and custom properties of a workbook.

### EXAMPLE 2
```powershell
Get-ChildItem -Path \\files\Reports -Filter *.xlsx |
                Get-OfficeExcelDocumentProperty -Name Owner, Classification -Kind Custom |
                Format-Table
```

Reads the custom Owner and Classification properties of every workbook on the share.

## PARAMETERS

### -Document
Caller-owned workbook; omission uses the enclosing Excel content block.

```yaml
Type: ExcelDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Kind
Metadata kind to add.

```yaml
Type: ExcelDocumentPropertyKind[]
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values: BuiltIn, Application, Custom

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LiteralPath
Literal workbook paths, including FileInfo.FullName and provider PSPath.

```yaml
Type: String[]
Parameter Sets: LiteralPath
Aliases: PSPath, FullName
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Name
Property name filter (wildcards supported).

```yaml
Type: String[]
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Workbook paths resolved with filesystem wildcard support.

```yaml
Type: String[]
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String[]`
- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `None`

## RELATED LINKS

- None
