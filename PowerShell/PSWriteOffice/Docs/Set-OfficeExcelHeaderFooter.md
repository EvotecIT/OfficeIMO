---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelHeaderFooter
## SYNOPSIS
Sets worksheet header and footer text and optional images.

## SYNTAX
### Context (Default)
```powershell
Set-OfficeExcelHeaderFooter [-HeaderLeft <string>] [-HeaderCenter <string>] [-HeaderRight <string>] [-FooterLeft <string>] [-FooterCenter <string>] [-FooterRight <string>] [-DifferentFirstPage] [-DifferentOddEven] [-AlignWithMargins <bool>] [-ScaleWithDocument <bool>] [-HeaderImagePath <string>] [-HeaderImageUrl <string>] [-HeaderImagePosition <ExcelHeaderFooterPosition>] [-FooterImagePath <string>] [-FooterImageUrl <string>] [-FooterImagePosition <ExcelHeaderFooterPosition>] [-ImageWidthPoints <Double>] [-ImageHeightPoints <Double>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Set-OfficeExcelHeaderFooter -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <Int32>] [-HeaderLeft <string>] [-HeaderCenter <string>] [-HeaderRight <string>] [-FooterLeft <string>] [-FooterCenter <string>] [-FooterRight <string>] [-DifferentFirstPage] [-DifferentOddEven] [-AlignWithMargins <bool>] [-ScaleWithDocument <bool>] [-HeaderImagePath <string>] [-HeaderImageUrl <string>] [-HeaderImagePosition <ExcelHeaderFooterPosition>] [-FooterImagePath <string>] [-FooterImageUrl <string>] [-FooterImagePosition <ExcelHeaderFooterPosition>] [-ImageWidthPoints <Double>] [-ImageHeightPoints <Double>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Uses OfficeIMO.Excel header/footer APIs and supports DSL or document usage.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ExcelSheet 'Data' { Set-OfficeExcelHeaderFooter -HeaderCenter 'Demo' -FooterRight 'Page &P of &N' }
```

Applies header and footer text to the worksheet.

## PARAMETERS

### -AlignWithMargins
Align header/footer with margins (default: true).

```yaml
Type: Boolean
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DifferentFirstPage
Use a different header/footer on the first page.

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

### -DifferentOddEven
Use different headers/footers on odd/even pages.

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

### -FooterCenter
Center footer text.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FooterImagePath
Footer image file path.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FooterImagePosition
Footer image position.

```yaml
Type: ExcelHeaderFooterPosition
Parameter Sets: Context, Document
Aliases: None
Possible values: Left, Center, Right

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FooterImageUrl
Footer image URL.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FooterLeft
Left footer text.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FooterRight
Right footer text.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HeaderCenter
Center header text.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HeaderImagePath
Header image file path.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HeaderImagePosition
Header image position.

```yaml
Type: ExcelHeaderFooterPosition
Parameter Sets: Context, Document
Aliases: None
Possible values: Left, Center, Right

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HeaderImageUrl
Header image URL.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HeaderLeft
Left header text.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HeaderRight
Right header text.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ImageHeightPoints
Image height in points.

```yaml
Type: Double
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ImageWidthPoints
Image width in points.

```yaml
Type: Double
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
Emit the worksheet after updating.

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

### -ScaleWithDocument
Scale header/footer with document (default: true).

```yaml
Type: Boolean
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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `None`

## RELATED LINKS

- None
