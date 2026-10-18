---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelHyperlink
## SYNOPSIS
Sets a hyperlink on a worksheet cell.

## SYNTAX
### ContextExternal (Default)
```powershell
Set-OfficeExcelHyperlink -Url <string> [-Row <Int32>] [-Column <Int32>] [-Address <string>] [-Display <string>] [-NoStyle] [-PassThru] [<CommonParameters>]
```

### DocumentExternal
```powershell
Set-OfficeExcelHyperlink -Document <ExcelDocument> -Url <string> [-Sheet <string>] [-SheetIndex <Int32>] [-Row <Int32>] [-Column <Int32>] [-Address <string>] [-Display <string>] [-NoStyle] [-PassThru] [<CommonParameters>]
```

### DocumentInternal
```powershell
Set-OfficeExcelHyperlink -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <Int32>] [-Row <Int32>] [-Column <Int32>] [-Address <string>] [-Location <string>] [-TargetSheet <string>] [-TargetAddress <string>] [-Display <string>] [-NoStyle] [-PassThru] [<CommonParameters>]
```

### ContextInternal
```powershell
Set-OfficeExcelHyperlink [-Row <Int32>] [-Column <Int32>] [-Address <string>] [-Location <string>] [-TargetSheet <string>] [-TargetAddress <string>] [-Display <string>] [-NoStyle] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Sets a hyperlink on a worksheet cell.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ExcelSheet 'Data' { Set-OfficeExcelHyperlink -Address 'A1' -Url 'https://example.org' -Display 'Example' }
```

Creates a styled hyperlink in A1.

### EXAMPLE 2
```powershell
PS> ExcelSheet 'Data' { Set-OfficeExcelHyperlink -Row 2 -Column 1 -TargetSheet 'Summary' -TargetAddress 'A1' -Display 'Go to Summary' }
```

Links A2 to Summary!A1.

## PARAMETERS

### -Address
A1-style cell address (e.g., A1, C5).

```yaml
Type: String
Parameter Sets: ContextExternal, DocumentExternal, DocumentInternal, ContextInternal
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Column
1-based column index.

```yaml
Type: Int32
Parameter Sets: ContextExternal, DocumentExternal, DocumentInternal, ContextInternal
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Display
Optional display text.

```yaml
Type: String
Parameter Sets: ContextExternal, DocumentExternal, DocumentInternal, ContextInternal
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
Parameter Sets: DocumentExternal, DocumentInternal
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Location
Internal location to link to (e.g., "'Summary'!A1").

```yaml
Type: String
Parameter Sets: DocumentInternal, ContextInternal
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoStyle
Skip hyperlink styling (blue + underline).

```yaml
Type: SwitchParameter
Parameter Sets: ContextExternal, DocumentExternal, DocumentInternal, ContextInternal
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the worksheet after setting the link.

```yaml
Type: SwitchParameter
Parameter Sets: ContextExternal, DocumentExternal, DocumentInternal, ContextInternal
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Row
1-based row index.

```yaml
Type: Int32
Parameter Sets: ContextExternal, DocumentExternal, DocumentInternal, ContextInternal
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
Parameter Sets: DocumentExternal, DocumentInternal
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
Parameter Sets: DocumentExternal, DocumentInternal
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TargetAddress
Target A1 address for internal links.

```yaml
Type: String
Parameter Sets: DocumentInternal, ContextInternal
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TargetSheet
Target worksheet name for internal links.

```yaml
Type: String
Parameter Sets: DocumentInternal, ContextInternal
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Url
External URL to link to.

```yaml
Type: String
Parameter Sets: ContextExternal, DocumentExternal
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
