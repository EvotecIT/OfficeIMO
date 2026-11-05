---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Update-OfficeExcelComment
## SYNOPSIS
Replaces the text or author of notes that match a filter.

## SYNTAX
### Text (Default)
```powershell
Update-OfficeExcelComment [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector[]>] [-Range <string>] [-MatchAuthor <string>] [-TextContains <string>] [-All] [-Author <string>] [-Initials <string>] [-PassThru] [-Text <string>] [<CommonParameters>]
```

### RichText
```powershell
Update-OfficeExcelComment [-Content] <scriptblock> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector[]>] [-Range <string>] [-MatchAuthor <string>] [-TextContains <string>] [-All] [-Author <string>] [-Initials <string>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Replaces the text or author of notes that match a filter.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeExcel -Path .\Finance\Budget-2027.xlsx {
    Update-OfficeExcelComment -Sheet Assumptions -TextContains 'TBD' -Text 'Confirmed with the vendor on 24 October.'
}
```

Replaces the text of every note that says TBD on the Assumptions sheet.

### EXAMPLE 2
```powershell
Open-OfficeExcel -Path .\Reports\Helpdesk.xlsx {
    Update-OfficeExcelComment -MatchAuthor 'jnowak' -Author 'Service Desk' -Initials 'SD'
    Update-OfficeExcelComment -Sheet Tickets -Range F40 {
        ExcelText 'Escalated: ' -Bold -Color Danger
        ExcelText 'vendor case 88412 opened.'
    }
}
```

Moves every note written by a departing colleague to the team account and rewrites one escalation note with rich text.

## PARAMETERS

### -All
Explicitly selects every item when no filter is supplied.

```yaml
Type: SwitchParameter
Parameter Sets: Text, RichText
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Author
Replacement author name.

```yaml
Type: String
Parameter Sets: Text, RichText
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Content
Excel child commands executed in the current content scope.

```yaml
Type: ScriptBlock
Parameter Sets: RichText
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Caller-owned workbook; omission uses the enclosing Excel content block.

```yaml
Type: ExcelDocument
Parameter Sets: Text, RichText
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Initials
Author initials, supplied together with Author.

```yaml
Type: String
Parameter Sets: Text, RichText
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MatchAuthor
Existing comment author to match, ignoring case.

```yaml
Type: String
Parameter Sets: Text, RichText
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
Parameter Sets: Text, RichText
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
Parameter Sets: Text, RichText
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
Type: OfficeSheetSelector[]
Parameter Sets: Text, RichText
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Replacement plain text.

```yaml
Type: String
Parameter Sets: Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TextContains
Existing text fragment to match, ignoring case.

```yaml
Type: String
Parameter Sets: Text, RichText
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
