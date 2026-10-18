---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelWorksheetView
## SYNOPSIS
Sets worksheet view options such as gridlines, direction, zoom, and view mode.

## SYNTAX
### Context (Default)
```powershell
Set-OfficeExcelWorksheetView [-ShowGridlines] [-HideGridlines] [-RightToLeft] [-LeftToRight] [-ZoomScale <UInt32>] [-ZoomScaleNormal <UInt32>] [-View <ExcelWorksheetViewKind>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Set-OfficeExcelWorksheetView -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <Int32>] [-ShowGridlines] [-HideGridlines] [-RightToLeft] [-LeftToRight] [-ZoomScale <UInt32>] [-ZoomScaleNormal <UInt32>] [-View <ExcelWorksheetViewKind>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Sets worksheet view options such as gridlines, direction, zoom, and view mode.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ExcelSheet 'Data' { Set-OfficeExcelWorksheetView -HideGridlines -ZoomScale 125 -View PageLayout }
```

Hides gridlines, sets zoom, and switches the sheet view.

## PARAMETERS

### -Document
Workbook to operate on outside the DSL context.

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

### -HideGridlines
Hide worksheet gridlines.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LeftToRight
Show worksheet left-to-right.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the worksheet after applying view options.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RightToLeft
Show worksheet right-to-left.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Worksheet name when using Document.

```yaml
Type: String
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SheetIndex
Worksheet index (0-based) when using Document.

```yaml
Type: Int32
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ShowGridlines
Show worksheet gridlines.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -View
Worksheet view mode.

```yaml
Type: ExcelWorksheetViewKind
Parameter Sets: Context, Document
Aliases: None
Possible values: Normal, PageBreakPreview, PageLayout

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ZoomScale
Active worksheet zoom percentage. Excel supports values from 10 to 400.

```yaml
Type: UInt32
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ZoomScaleNormal
Normal-view worksheet zoom percentage. Excel supports values from 10 to 400.

```yaml
Type: UInt32
Parameter Sets: Context, Document
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

- `None`

## RELATED LINKS

- None
