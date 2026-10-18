---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Send-OfficePdfPrinter
## SYNOPSIS
Prepares PDF sheets and submits them to an explicitly named printer. Requires PowerShell 7.4 or newer.

## SYNTAX
### __AllParameterSets
```powershell
Send-OfficePdfPrinter [-Path] <string> -PrinterName <string> [-Pages <string>] [-PagesPerSheet <int>] [-Copies <int>] [-Duplex <string>] [-PaperSourceId <string>] [-OutputFilePath <string>] [-Dpi <double>] [-Margin <double>] [-Orientation <string>] [-ScaleMode <string>] [-PaperSize <PageSize>] [-Password <string>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
The returned receipt proves queue acceptance. It does not prove physical delivery. Check the queue after an interrupted submission before retrying.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Send-OfficePdfPrinter -Path .\Report.pdf -PrinterName 'Office printer' -Pages '1-3' -PagesPerSheet 2 -Duplex LongEdge
```


## PARAMETERS

### -Copies
Copies submitted to the queue.

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

### -Dpi
Prepared sheet resolution.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Duplex
Duplex setting.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: PrinterDefault, SingleSided, LongEdge, ShortEdge

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Margin
Printable margin in points.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Orientation
Paper orientation.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Automatic, Portrait, Landscape

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputFilePath
New destination file for a Windows file printer.

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

### -Pages
Page selection such as 1-3,last.

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

### -PagesPerSheet
Source pages per sheet.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: 1, 2, 4

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PaperSize
PDF paper size, default A4.

```yaml
Type: PageSize
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PaperSourceId
Paper source identifier reported for this queue.

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

### -Password
Password for an encrypted PDF; printing permissions remain enforced.

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

### -Path
Local PDF source.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PrinterName
Installed queue from Get-OfficePrinter.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ScaleMode
Scaling of source pages.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Fit, ActualSize, Fill

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

- `OfficeIMO.Workflows.PdfPrintSubmission`

## RELATED LINKS

- None
