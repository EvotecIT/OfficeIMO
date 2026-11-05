---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeExcelComment
## SYNOPSIS
Adds a note or a threaded comment (or reply) to a cell.

## SYNTAX
### Note (Default)
```powershell
Add-OfficeExcelComment [-Address] <string> [-Text] <string> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-Author <string>] [-PassThru] [-Initials <string>] [<CommonParameters>]
```

### Threaded
```powershell
Add-OfficeExcelComment [-Address] <string> [-Text] <string> -Threaded [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-Author <string>] [-PassThru] [-ParentId <string>] [-Id <string>] [-Date <DateTime>] [-Done] [<CommonParameters>]
```

## DESCRIPTION
Adds a note or a threaded comment (or reply) to a cell.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeExcel -Path .\Finance\Budget-2027.xlsx {
    ExcelSheet 'Assumptions' {
        ExcelComment B4 'Inflation from the October central bank projection.' -Author 'Finance'
    }
}
```

Explains the source of a figure in a note on B4.

### EXAMPLE 2
```powershell
Open-OfficeExcel -Path .\Finance\Forecast-2027.xlsx {
    ExcelSheet 'Forecast' {
        ExcelComment F12 'Can we confirm the Q2 hiring plan before this goes to the board?' -Threaded -Author 'Anna Kowalska'
    }
}
```

Adds a threaded comment asking for confirmation of a forecast figure, attributed to the reviewer.

## PARAMETERS

### -Address
A1 address of the target cell.

```yaml
Type: String
Parameter Sets: Note, Threaded
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Author
Author name (optional).

```yaml
Type: String
Parameter Sets: Note, Threaded
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Date
Optional timestamp for the threaded comment.

```yaml
Type: DateTime
Parameter Sets: Threaded
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
Parameter Sets: Note, Threaded
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Done
Mark the threaded comment as done/resolved.

```yaml
Type: SwitchParameter
Parameter Sets: Threaded
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Id
Optional stable threaded-comment id.

```yaml
Type: String
Parameter Sets: Threaded
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Initials
Author initials, supplied together with Author.

```yaml
Type: String
Parameter Sets: Note
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ParentId
Optional parent threaded-comment id when adding a reply.

```yaml
Type: String
Parameter Sets: Threaded
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
Parameter Sets: Note, Threaded
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
Parameter Sets: Note, Threaded
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Comment text.

```yaml
Type: String
Parameter Sets: Note, Threaded
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Threaded
Creates a threaded comment instead of a note.

```yaml
Type: SwitchParameter
Parameter Sets: Threaded
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
