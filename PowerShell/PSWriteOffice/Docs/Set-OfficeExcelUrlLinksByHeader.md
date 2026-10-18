---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelUrlLinksByHeader
## SYNOPSIS
Converts cells under a header into external URL hyperlinks.

## SYNTAX
### ContextUsedRange (Default)
```powershell
Set-OfficeExcelUrlLinksByHeader [-Header] <string> -UrlScript <scriptblock> [-TitleScript <scriptblock>] [-NoStyle] [-PassThru] [<CommonParameters>]
```

### DocumentUsedRange
```powershell
Set-OfficeExcelUrlLinksByHeader [-Header] <string> -Document <ExcelDocument> -UrlScript <scriptblock> [-Sheet <string>] [-SheetIndex <Int32>] [-TitleScript <scriptblock>] [-NoStyle] [-PassThru] [<CommonParameters>]
```

### DocumentTable
```powershell
Set-OfficeExcelUrlLinksByHeader [-Header] <string> -Document <ExcelDocument> -TableName <string> -UrlScript <scriptblock> [-Sheet <string>] [-SheetIndex <Int32>] [-TitleScript <scriptblock>] [-NoStyle] [-PassThru] [<CommonParameters>]
```

### DocumentRange
```powershell
Set-OfficeExcelUrlLinksByHeader [-Header] <string> -Document <ExcelDocument> -Range <string> -UrlScript <scriptblock> [-Sheet <string>] [-SheetIndex <Int32>] [-TitleScript <scriptblock>] [-NoStyle] [-PassThru] [<CommonParameters>]
```

### ContextTable
```powershell
Set-OfficeExcelUrlLinksByHeader [-Header] <string> -TableName <string> -UrlScript <scriptblock> [-TitleScript <scriptblock>] [-NoStyle] [-PassThru] [<CommonParameters>]
```

### ContextRange
```powershell
Set-OfficeExcelUrlLinksByHeader [-Header] <string> -Range <string> -UrlScript <scriptblock> [-TitleScript <scriptblock>] [-NoStyle] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Converts cells under a header into external URL hyperlinks.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ExcelSheet 'Summary' { Set-OfficeExcelUrlLinksByHeader -Header 'RFC' -TableName 'Links' -UrlScript { param($text) "https://datatracker.ietf.org/doc/html/$text" } }
```

Uses the RFC column values to create external hyperlinks.

## PARAMETERS

### -Document
Workbook to operate on outside the DSL context.

```yaml
Type: ExcelDocument
Parameter Sets: DocumentUsedRange, DocumentTable, DocumentRange
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Header
Header text to locate.

```yaml
Type: String
Parameter Sets: ContextUsedRange, DocumentUsedRange, DocumentTable, DocumentRange, ContextTable, ContextRange
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoStyle
Skip hyperlink styling (blue + underline).

```yaml
Type: SwitchParameter
Parameter Sets: ContextUsedRange, DocumentUsedRange, DocumentTable, DocumentRange, ContextTable, ContextRange
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the worksheet after creating links.

```yaml
Type: SwitchParameter
Parameter Sets: ContextUsedRange, DocumentUsedRange, DocumentTable, DocumentRange, ContextTable, ContextRange
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Range
Restrict linking to a specific A1 range whose first row contains headers.

```yaml
Type: String
Parameter Sets: DocumentRange, ContextRange
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Worksheet name when using Document.

```yaml
Type: String
Parameter Sets: DocumentUsedRange, DocumentTable, DocumentRange
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
Parameter Sets: DocumentUsedRange, DocumentTable, DocumentRange
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TableName
Restrict linking to a named table.

```yaml
Type: String
Parameter Sets: DocumentTable, ContextTable
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TitleScript
Optional mapping from cell text to display text.

```yaml
Type: ScriptBlock
Parameter Sets: ContextUsedRange, DocumentUsedRange, DocumentTable, DocumentRange, ContextTable, ContextRange
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -UrlScript
Maps the cell text to a URL.

```yaml
Type: ScriptBlock
Parameter Sets: ContextUsedRange, DocumentUsedRange, DocumentTable, DocumentRange, ContextTable, ContextRange
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

- `OfficeIMO.Excel.ExcelSheet`

## RELATED LINKS

- None
