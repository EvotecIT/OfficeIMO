---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficeExcelHyperlink
## SYNOPSIS
Links a cell, or every cell of a range or header column, to a URL or a place in the workbook.

## SYNTAX
### Url (Default)
```powershell
Set-OfficeExcelHyperlink [-Address] <string> [-Url] <string> [-Display <string>] [-DisplayMode <ExcelHyperlinkDisplay>] [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-NoStyle] [-PassThru] [<CommonParameters>]
```

### Internal
```powershell
Set-OfficeExcelHyperlink [-Address] <string> -TargetAddress <string> [-Display <string>] [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-NoStyle] [-PassThru] [-TargetSheet <OfficeSheetSelector>] [<CommonParameters>]
```

### InternalRange
```powershell
Set-OfficeExcelHyperlink [-Range] <string> -Internal [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-NoStyle] [-PassThru] [-TargetAddress <string>] [-DisplayScript <scriptblock>] [-TargetSheetScript <scriptblock>] [<CommonParameters>]
```

### InternalHeader
```powershell
Set-OfficeExcelHyperlink [-Header] <string> -Internal [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-NoStyle] [-PassThru] [-TargetAddress <string>] [-Range <string>] [-DisplayScript <scriptblock>] [-Table <string>] [-TargetSheetScript <scriptblock>] [<CommonParameters>]
```

### UrlRange
```powershell
Set-OfficeExcelHyperlink [-Range] <string> -UrlScript <scriptblock> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-NoStyle] [-PassThru] [-DisplayScript <scriptblock>] [<CommonParameters>]
```

### UrlHeader
```powershell
Set-OfficeExcelHyperlink [-Header] <string> -UrlScript <scriptblock> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-NoStyle] [-PassThru] [-Range <string>] [-DisplayScript <scriptblock>] [-Table <string>] [<CommonParameters>]
```

## DESCRIPTION
Links a cell, or every cell of a range or header column, to a URL or a place in the workbook.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeExcel -Path .\Reports\Status.xlsx {
    ExcelSheet 'Summary' {
        ExcelHyperlink A2 'https://status.example.com' -Display 'Service status'
    }
}
```

Turns A2 into a link to the status page with readable text.

### EXAMPLE 2
```powershell
Open-OfficeExcel -Path .\Reports\Helpdesk.xlsx {
    ExcelSheet 'Tickets' {
        ExcelHyperlink -Header Ticket -Table Tickets -UrlScript { param($text) "https://helpdesk.example.com/tickets/$text" }
    }
}
```

Links each value in the Ticket column of the Tickets table to its page in the helpdesk system.

### EXAMPLE 3
```powershell
$servers = 'SQL01', 'SQL02', 'WEB01'
Open-OfficeExcel -Path .\Reports\Disk-Capacity.xlsx {
    ExcelSheet 'Summary' {
        ExcelHyperlink -Header Server -Internal
    }
    foreach ($server in $servers) {
        ExcelSheet $server {
            ExcelHyperlink H1 -TargetSheet Summary -TargetAddress A1 -Display 'Back to summary'
        }
    }
}
```

Links each server name in the Summary sheet to the sheet of the same name, and adds a link back to the summary on every server sheet.

## PARAMETERS

### -Address
A1 address of the target cell.

```yaml
Type: String
Parameter Sets: Url, Internal
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Display
Optional display text.

```yaml
Type: String
Parameter Sets: Url, Internal
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DisplayMode
Selects original text, full URL, host name or smart hyperlink display.

```yaml
Type: ExcelHyperlinkDisplay
Parameter Sets: Url
Aliases: None
Possible values: Text, Url, Host, Smart

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DisplayScript
Optional mapping from cell text to display text.

```yaml
Type: ScriptBlock
Parameter Sets: InternalRange, InternalHeader, UrlRange, UrlHeader
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
Parameter Sets: Url, Internal, InternalRange, InternalHeader, UrlRange, UrlHeader
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Header
Header text identifying the target column.

```yaml
Type: String
Parameter Sets: InternalHeader, UrlHeader
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Internal
Uses worksheet destinations instead of external URLs for header-based links.

```yaml
Type: SwitchParameter
Parameter Sets: InternalRange, InternalHeader
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoStyle
Skip hyperlink styling (blue + underline).

```yaml
Type: SwitchParameter
Parameter Sets: Url, Internal, InternalRange, InternalHeader, UrlRange, UrlHeader
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
Parameter Sets: Url, Internal, InternalRange, InternalHeader, UrlRange, UrlHeader
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Range
A1 cell range on the selected worksheet.

```yaml
Type: String
Parameter Sets: InternalRange, InternalHeader, UrlRange, UrlHeader
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.

```yaml
Type: OfficeSheetSelector
Parameter Sets: Url, Internal, InternalRange, InternalHeader, UrlRange, UrlHeader
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Table
Native Excel table used to resolve the source range or header.

```yaml
Type: String
Parameter Sets: InternalHeader, UrlHeader
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
Parameter Sets: Internal, InternalRange, InternalHeader
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
Type: OfficeSheetSelector
Parameter Sets: Internal
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TargetSheetScript
Script that maps each source cell text to a destination worksheet name.

```yaml
Type: ScriptBlock
Parameter Sets: InternalRange, InternalHeader
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
Parameter Sets: Url
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -UrlScript
Maps the cell text to a URL.

```yaml
Type: ScriptBlock
Parameter Sets: UrlRange, UrlHeader
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
