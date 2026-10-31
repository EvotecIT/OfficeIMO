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
### __AllParameterSets
```powershell
New-OfficeExcel [[-Path] <string>] [[-Content] <scriptblock>] [-Theme <OfficeThemeInput>] [-Template <string>] [-Force] [-Password <securestring>] [-SaveOptions <ExcelSaveOptions>] [-DocumentProperty <hashtable>] [-DateSystem <ExcelDateSystem>] [-MaxDegreeOfParallelism <Int32>] [-PassThru] [-AllowLoss] [-Show] [-WhatIf] [-Confirm] [<CommonParameters>]
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


## PARAMETERS

### -AllowLoss
Gets or sets whether native-save loss is explicitly accepted.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Content
Gets or sets a content block; omission returns the live document.

```yaml
Type: ScriptBlock
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DateSystem
Native workbook date system.

```yaml
Type: ExcelDateSystem
Parameter Sets: __AllParameterSets
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
Parameter Sets: __AllParameterSets
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
Parameter Sets: __AllParameterSets
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
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Gets or sets whether a completed content block returns the saved FileInfo.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
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
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Optional associated destination; required with a content block.

```yaml
Type: String
Parameter Sets: __AllParameterSets
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
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Show
Gets or sets whether the saved file opens in the system default application.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
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
Parameter Sets: __AllParameterSets
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
Parameter Sets: __AllParameterSets
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
