---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeExcelCommentAudit
## SYNOPSIS
Audits legacy notes and threaded comments preserved in an Excel workbook.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeExcelCommentAudit [-Path] <string> [-IncludeComments] [<CommonParameters>]
```

### Document
```powershell
Get-OfficeExcelCommentAudit -Document <ExcelDocument> [-IncludeComments] [<CommonParameters>]
```

## DESCRIPTION
Audits legacy notes and threaded comments preserved in an Excel workbook.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $audit = Get-OfficeExcelCommentAudit -Path .\ReviewWorkbook.xlsx -IncludeComments
$audit.Comments | Sort-Object SheetName,CellReference | Format-Table SheetName,CellReference,Author,Text
$audit.Issues | Format-Table Severity,Category,SheetName,Address,Message
```

Returns workbook-level note/threaded-comment counts, optional comment records, and metadata issues such as missing authors or orphaned threaded replies.

## PARAMETERS

### -Document
Workbook document.

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

### -IncludeComments
Include legacy and threaded comment records in the output.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Workbook path.

```yaml
Type: String
Parameter Sets: Path
Aliases: InputPath, FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `System.Management.Automation.PSObject`

## RELATED LINKS

- None
