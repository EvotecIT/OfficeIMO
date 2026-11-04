---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficeExcelSheetView
## SYNOPSIS
Sets frozen panes, gridlines, direction, zoom and view mode for a worksheet.

## SYNTAX
### Default (Default)
```powershell
Set-OfficeExcelSheetView [-Sheet <OfficeSheetSelector>] [-FreezeRows <Int32>] [-FreezeColumns <Int32>] [-NoGridlines] [-RightToLeft] [-Zoom <UInt32>] [-ZoomNormal <UInt32>] [-View <ExcelWorksheetViewKind>] [-PassThru] [-Document <ExcelDocument>] [<CommonParameters>]
```

## DESCRIPTION
Omitted settings keep their values; explicitly false switches reverse their setting.
Set both frozen dimensions to zero to unfreeze the sheet.

## EXAMPLES

### EXAMPLE 1
```powershell
ExcelSheet 'Dashboard' { ExcelSheetView -FreezeRows 3 -FreezeColumns 1 -NoGridlines -Zoom 85 }
```


## PARAMETERS

### -Document
Gets or sets an explicit live document; omission uses the enclosing content block.

```yaml
Type: ExcelDocument
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -FreezeColumns
Gets or sets the number of frozen left columns; omission preserves the current column count.

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

### -FreezeRows
Gets or sets the number of frozen top rows; omission preserves the current row count.

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

### -NoGridlines
Gets or sets whether gridlines are hidden; explicitly false shows them.

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

### -PassThru
Gets or sets whether to emit the edited native worksheet.

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

### -RightToLeft
Gets or sets right-to-left layout; explicitly false restores left-to-right.

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

### -Sheet
Gets or sets a worksheet name, one-based position or native worksheet.

```yaml
Type: OfficeSheetSelector
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -View
Gets or sets the worksheet view mode.

```yaml
Type: ExcelWorksheetViewKind
Parameter Sets: Default
Aliases: None
Possible values: Normal, PageBreakPreview, PageLayout

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Zoom
Gets or sets the current zoom percentage from 10 to 400.

```yaml
Type: UInt32
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ZoomNormal
Gets or sets the normal-view zoom percentage from 10 to 400.

```yaml
Type: UInt32
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

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `OfficeIMO.Excel.ExcelSheet`

## RELATED LINKS

- None
