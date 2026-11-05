---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeExcelConnection
## SYNOPSIS
Adds Power Query or connection metadata that Excel can own and refresh.

## SYNTAX
### PowerQuery (Default)
```powershell
Add-OfficeExcelConnection -Name <string> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-PassThru] [-CommandText <string>] [-Description <string>] [-QueryTable <string>] [-RefreshOnOpen] [<CommonParameters>]
```

### Xml
```powershell
Add-OfficeExcelConnection -Kind <OfficeExcelPackageMetadataKind> -Xml <string> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds Power Query or connection metadata that Excel can own and refresh.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeExcel -Path .\Reports\Sales-2026.xlsx {
    ExcelConnection -Name SalesQuery -CommandText 'let Source = Excel.CurrentWorkbook(){[Name="Orders"]}[Content] in Source'
}
```

Stores a Power Query query that Excel can refresh; OfficeIMO writes the definition and never runs it.

### EXAMPLE 2
```powershell
$query = 'let Source = Sql.Database("dw.example.com", "Sales"), Orders = Source{[Schema="dbo", Item="Orders"]}[Data] in Orders'
Open-OfficeExcel -Path .\Reports\Live-Sales.xlsx {
    ExcelConnection -Sheet Data -Name WarehouseOrders -CommandText $query -Description 'Orders from the sales warehouse' -QueryTable OrdersQuery -RefreshOnOpen
}
```

Adds a described query bound to a query table on the Data sheet and marks it to refresh when the workbook opens.

## PARAMETERS

### -CommandText
Power Query M expression stored as metadata.

```yaml
Type: String
Parameter Sets: PowerQuery
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Description
Connection description.

```yaml
Type: String
Parameter Sets: PowerQuery
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Caller-owned workbook; omission uses the enclosing Excel content block.

```yaml
Type: ExcelDocument
Parameter Sets: PowerQuery, Xml
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
Type: OfficeExcelPackageMetadataKind
Parameter Sets: Xml
Aliases: None
Possible values: Connection, QueryTable

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Name
Connection name stored in workbook metadata.

```yaml
Type: String
Parameter Sets: PowerQuery
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emits the native result after the operation.

```yaml
Type: SwitchParameter
Parameter Sets: PowerQuery, Xml
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -QueryTable
Native query-table metadata used by the connection.

```yaml
Type: String
Parameter Sets: PowerQuery
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RefreshOnOpen
Request refresh-on-open metadata.

```yaml
Type: SwitchParameter
Parameter Sets: PowerQuery
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.

```yaml
Type: OfficeSheetSelector
Parameter Sets: PowerQuery, Xml
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Xml
XML payload to add as package metadata.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `None`

## RELATED LINKS

- None
