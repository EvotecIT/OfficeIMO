---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Send-OfficePdfPrinter
## SYNOPSIS
Prepares PDF sheets and submits them to an explicitly named printer. Requires PowerShell 7.4 or newer.

## SYNTAX
### Path (Default)
```powershell
Send-OfficePdfPrinter [-Path] <string[]> -PrinterName <string> [-Page <OfficePageSelector>] [-PagesPerSheet <int>] [-Copies <int>] [-Duplex <string>] [-PaperSourceId <string>] [-OutputPath <string>] [-Dpi <double>] [-Margin <OfficeLength>] [-Orientation <OfficePageOrientation>] [-ScaleMode <string>] [-PaperSize <OfficePageSize>] [-Password <securestring>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPath
```powershell
Send-OfficePdfPrinter -LiteralPath <string[]> -PrinterName <string> [-Page <OfficePageSelector>] [-PagesPerSheet <int>] [-Copies <int>] [-Duplex <string>] [-PaperSourceId <string>] [-OutputPath <string>] [-Dpi <double>] [-Margin <OfficeLength>] [-Orientation <OfficePageOrientation>] [-ScaleMode <string>] [-PaperSize <OfficePageSize>] [-Password <securestring>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
The returned receipt proves queue acceptance. It does not prove physical delivery. Check the queue after an interrupted submission before retrying.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Send-OfficePdfPrinter -Path .\Report.pdf -PrinterName 'Office printer' -Page '1-3' -PagesPerSheet 2 -Duplex LongEdge
```


## PARAMETERS

### -Copies
Copies submitted to the queue.

```yaml
Type: Int32
Parameter Sets: Path, LiteralPath
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
Parameter Sets: Path, LiteralPath
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
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values: PrinterDefault, SingleSided, LongEdge, ShortEdge

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LiteralPath
Literal PDF sources, including FileInfo.FullName.

```yaml
Type: String[]
Parameter Sets: LiteralPath
Aliases: FullName, PSPath
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Margin
Printable margin in points.

```yaml
Type: OfficeLength
Parameter Sets: Path, LiteralPath
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
Type: OfficePageOrientation
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values: Portrait, Landscape

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputPath
New destination file for a Windows file printer.

```yaml
Type: String
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Page
Page selection such as 1-3,last.

```yaml
Type: OfficePageSelector
Parameter Sets: Path, LiteralPath
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
Parameter Sets: Path, LiteralPath
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
Type: OfficePageSize
Parameter Sets: Path, LiteralPath
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
Password for an encrypted PDF; printing permissions remain enforced.

```yaml
Type: SecureString
Parameter Sets: Path, LiteralPath
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
Type: String[]
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PrinterName
Installed queue from Get-OfficePrinter.

```yaml
Type: String
Parameter Sets: Path, LiteralPath
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
Parameter Sets: Path, LiteralPath
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

- `System.String[]`

## OUTPUTS

- `OfficeIMO.Workflows.PdfPrintSubmission`

## RELATED LINKS

- None
