---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeExcel
## SYNOPSIS
Creates a themed Excel workbook, returning a live workbook or saving a completed content block.

## SYNTAX
### Live (Default)
```powershell
New-OfficeExcel [[-Path] <string>] [-Theme <OfficeThemeInput>] [-Template <string>] [-Force] [-DocumentProperty <hashtable>] [-DateSystem <ExcelDateSystem>] [-MaxDegreeOfParallelism <Int32>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Content
```powershell
New-OfficeExcel [-Path] <string> [-Content] <scriptblock> [-Theme <OfficeThemeInput>] [-Template <string>] [-Force] [-Password <securestring>] [-SaveOptions <ExcelSaveOptions>] [-DocumentProperty <hashtable>] [-DateSystem <ExcelDateSystem>] [-MaxDegreeOfParallelism <Int32>] [-PassThru] [-AllowLoss] [-Show] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Without a block no file is written, including on disposal. A block saves atomically; an existing destination requires Force. Template copies content; Theme reads only the visual theme. SaveOptions and Password apply to a completed content block.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficeExcel -Path ./Report.xlsx -Theme Modern -PassThru {
    ExcelSheet 'Data' { ExcelCell -Address A1 -Value 'Region' }
}
```


### EXAMPLE 2
```powershell
$workbook = New-OfficeExcel -Theme Boardroom
$sheet = Add-OfficeExcelSheet -Document $workbook -Name Data -PassThru
Set-OfficeExcelCell -Sheet $sheet -Address A1 -Value 'Region'
Save-OfficeExcel -Document $workbook -Path ./Report.xlsx
Close-OfficeExcel -Document $workbook
```


### EXAMPLE 3
```powershell
New-OfficeExcel -Path .\Services.xlsx {
    ExcelSheet 'Services' {
        ExcelTable (Get-Service | Select-Object Name, DisplayName, Status, StartType)
    }
}
```

Creates Services.xlsx with a themed, filterable table of the local services and saves it when the block ends.

### EXAMPLE 4
```powershell
$orders = Import-OfficeCsv -Path .\Data\Q3-Orders.csv
$byRegion = $orders | Group-Object Region | ForEach-Object {
    [pscustomobject]@{
        Region  = $_.Name
        Orders  = $_.Count
        Revenue = ($_.Group | Measure-Object Amount -Sum).Sum
    }
}

$properties = @{ Title = 'Q3 sales'; Company = 'Northwind Traders' }
New-OfficeExcel -Path .\Reports\Q3-Sales.xlsx -Theme Boardroom -DocumentProperty $properties -Force -Show {
    ExcelSheet 'Summary' {
        ExcelTable $byRegion -Name RegionSummary
        ExcelChart -Table RegionSummary -Address F2 -Type Column -Title 'Revenue by region'
    }
    ExcelSheet 'Orders' {
        ExcelTable $orders -Name Orders
    }
}
```

Builds a Q3 sales workbook in the Boardroom theme with a summary chart and the order detail, sets the document title and opens it in Excel.

### EXAMPLE 5
```powershell
$book = New-OfficeExcel -Theme Modern
foreach ($server in 'SQL01', 'SQL02', 'WEB01') {
    $disks = Invoke-Command -ComputerName $server { Get-Volume | Where-Object DriveLetter }
    Add-OfficeExcelSheet -Document $book -Name $server {
        ExcelTable ($disks | Select-Object DriveLetter, FileSystemLabel, Size, SizeRemaining)
    }
}
Save-OfficeExcel -Document $book -Path .\Reports\Disk-Capacity.xlsx
Close-OfficeExcel -Document $book
```

Without a content block the live workbook is returned, so a script can add a sheet per server in a loop and choose the file name when it saves.

## PARAMETERS

### -AllowLoss
Explicitly accept reductions reported by the content block's native save.

```yaml
Type: SwitchParameter
Parameter Sets: Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Content
Authoring block to complete before the document is atomically saved and closed.

```yaml
Type: ScriptBlock
Parameter Sets: Content
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DateSystem
Native workbook date system.

```yaml
Type: ExcelDateSystem
Parameter Sets: Live, Content
Aliases: None
Possible values: NineteenHundred, NineteenFour

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DocumentProperty
Workbook properties such as Title, Author, Subject, Keywords, Company and Manager.

```yaml
Type: Hashtable
Parameter Sets: Live, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Force
Allow a completed content block to replace an existing destination.

```yaml
Type: SwitchParameter
Parameter Sets: Live, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxDegreeOfParallelism
Maximum parallel workers; 1 runs sequentially. Omission retains the native execution policy.

```yaml
Type: Int32
Parameter Sets: Live, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Return the saved FileInfo after a completed content block.

```yaml
Type: SwitchParameter
Parameter Sets: Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Password for a completed content block's saved file; use Save-OfficeExcel for a live workbook.

```yaml
Type: SecureString
Parameter Sets: Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Optional associated destination for a live document; required when a content block saves a file.

```yaml
Type: String
Parameter Sets: Live, Content
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SaveOptions
Native save options or a hashtable, applied to a completed content block.

```yaml
Type: ExcelSaveOptions
Parameter Sets: Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Show
Open the saved result in its registered application after the content block succeeds.

```yaml
Type: SwitchParameter
Parameter Sets: Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Template
Existing workbook copied as starting content. Its styling is preserved unless Theme is supplied.

```yaml
Type: String
Parameter Sets: Live, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Theme
Built-in theme, theme object or a document/template whose visual theme is read.

```yaml
Type: OfficeThemeInput
Parameter Sets: Live, Content
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

- `None`

## OUTPUTS

- `OfficeIMO.Excel.ExcelDocument`
- `System.IO.FileInfo`

## RELATED LINKS

- None
