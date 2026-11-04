---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Open-OfficeExcel
## SYNOPSIS
Opens a workbook for explicit editing, or atomically saves a completed content block.

## SYNTAX
### Path (Default)
```powershell
Open-OfficeExcel [-Path] <string> [[-Content] <scriptblock>] [-ReadOnly] [-Password <securestring>] [-SaveOptions <ExcelSaveOptions>] [-AllowLoss] [-MaxDegreeOfParallelism <Int32>] [-PassThru] [-Show] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPath
```powershell
Open-OfficeExcel [[-Content] <scriptblock>] -LiteralPath <string> [-ReadOnly] [-Password <securestring>] [-SaveOptions <ExcelSaveOptions>] [-AllowLoss] [-MaxDegreeOfParallelism <Int32>] [-PassThru] [-Show] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Uri
```powershell
Open-OfficeExcel [-Uri] <uri> [-AllowHttp] [-ReadOnly] [-Password <securestring>] [-MaxDegreeOfParallelism <Int32>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Opening preserves the workbook's theme and encryption. A failed block leaves its source unchanged. Remote workbooks remain detached; save them with Save-OfficeExcel -Path.

## EXAMPLES

### EXAMPLE 1
```powershell
$workbook = Get-Item -LiteralPath './Review[final].xlsx' | Open-OfficeExcel
$workbook | Close-OfficeExcel
```


### EXAMPLE 2
```powershell
Open-OfficeExcel -Path ./Report.xlsx {
    ExcelSheet 'Data' { ExcelCell -Address B2 -Value 'Reviewed' }
}
```


### EXAMPLE 3
```powershell
$workbook = Open-OfficeExcel -Uri https://intranet.example.com/Budget.xlsx -ReadOnly
Import-OfficeExcel -Document $workbook -Sheet Budget
Close-OfficeExcel -Document $workbook
```


### EXAMPLE 4
```powershell
Open-OfficeExcel -Path .\Reports\Q3-Sales.xlsx {
    ExcelSheet 'Summary' {
        ExcelCell B1 (Get-Date)
    }
}
```

Opens the Q3 sales workbook, stamps the review date on the Summary sheet and saves it back when the block ends.

### EXAMPLE 5
```powershell
$newTickets = Import-OfficeCsv -Path .\Exports\Tickets-2026-10-26.csv
Open-OfficeExcel -Path .\Reports\Helpdesk.xlsx {
    ExcelSheet 'Tickets' {
        Add-OfficeExcelTableRow -Table Tickets -InputObject $newTickets
        ExcelConditionalFormat -Header AgeDays -Table Tickets -Color Accent2
    }
}
```

Appends the tickets exported overnight to the Tickets table and adds a data bar to the age column, then saves the tracker in place.

### EXAMPLE 6
```powershell
$book = Open-OfficeExcel -Uri https://intranet.example.com/finance/Budget-2027.xlsx -ReadOnly
Import-OfficeExcel -Document $book -Sheet Budget | Where-Object Department -eq 'IT'
Close-OfficeExcel -Document $book
```

Opens a published budget read-only from HTTPS, reads the IT rows and closes it without writing anything back.

## PARAMETERS

### -AllowHttp
Allow HTTP in addition to HTTPS for a remote workbook.

```yaml
Type: SwitchParameter
Parameter Sets: Uri
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AllowLoss
Explicitly accept native-save reductions in a completed local content block.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Content
Edit a local workbook and save only after the block completes successfully.

```yaml
Type: ScriptBlock
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LiteralPath
Literal workbook path; accepts FileInfo.FullName and provider PSPath.

```yaml
Type: String
Parameter Sets: LiteralPath
Aliases: PSPath, FullName
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -MaxDegreeOfParallelism
Maximum parallel workers; omission preserves the native execution policy.

```yaml
Type: Int32
Parameter Sets: Path, LiteralPath, Uri
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Return the saved FileInfo from a completed local content block.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Source password as a SecureString, credential, or warned plaintext value.

```yaml
Type: SecureString
Parameter Sets: Path, LiteralPath, Uri
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Workbook path with FileSystem wildcard expansion.

```yaml
Type: String
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -ReadOnly
Return a read-only live workbook; authoring blocks require writable input.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath, Uri
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SaveOptions
Native save options or a hashtable, applied only to a completed local content block.

```yaml
Type: ExcelSaveOptions
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Show
Open the saved file in its registered application after a local block completes.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Uri
HTTP or HTTPS workbook loaded without a local save destination.

```yaml
Type: Uri
Parameter Sets: Uri
Aliases: None
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

- `System.String`

## OUTPUTS

- `OfficeIMO.Excel.ExcelDocument`
- `System.IO.FileInfo`

## RELATED LINKS

- None
