---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Send-OfficePrinter
## SYNOPSIS
Prints a document of any format OfficeIMO converts to PDF on an explicitly named printer queue or IPP printer. Requires PowerShell 7.4 or newer.

## SYNTAX
### Path (Default)
```powershell
Send-OfficePrinter [-Path] <string[]> -PrinterName <string> [-Page <OfficePageSelector>] [-PagesPerSheet <int>] [-PageSubset <string>] [-Grayscale] [-RasterizeGrayscale] [-Copies <int>] [-Duplex <string>] [-PaperSourceId <string>] [-OutputPath <string>] [-Dpi <double>] [-Margin <OfficeLength>] [-Orientation <OfficePageOrientation>] [-ScaleMode <string>] [-PaperSize <OfficePageSize>] [-Password <securestring>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPath
```powershell
Send-OfficePrinter -LiteralPath <string[]> -PrinterName <string> [-Page <OfficePageSelector>] [-PagesPerSheet <int>] [-PageSubset <string>] [-Grayscale] [-RasterizeGrayscale] [-Copies <int>] [-Duplex <string>] [-PaperSourceId <string>] [-OutputPath <string>] [-Dpi <double>] [-Margin <OfficeLength>] [-Orientation <OfficePageOrientation>] [-ScaleMode <string>] [-PaperSize <OfficePageSize>] [-Password <securestring>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
PDF files are printed directly; other documents are converted to PDF in memory first, and conversion warnings are written as warnings. Send-OfficePdfPrinter remains an alias of this command. PrinterName accepts an installed queue name or an ipp:// or ipps:// printer URI.

Without PaperSize, sheets use the printer's default media. The returned receipt proves queue acceptance, not physical delivery. Follow it with Get-OfficePrintJob, and check the queue after an interrupted submission before retrying.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Send-OfficePrinter -Path .\Report.pdf -PrinterName 'Office printer' -Page '1-3' -PagesPerSheet 2 -Duplex LongEdge
```


### EXAMPLE 2
```powershell
PS> Send-OfficePrinter -Path .\Letter.docx -PrinterName 'ipp://printer.local/ipp/print' | Get-OfficePrintJob -Wait
```


### EXAMPLE 3
```powershell
PS> Send-OfficePrinter -Path .\Report.pdf -PrinterName 'Office printer' -PageSubset Odd -PagesPerSheet 4 -Grayscale
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
Highest resolution used when sheets are rasterized for delivery.

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

### -Grayscale
Print in grayscale, regardless of the device's color support. IPP printers and CUPS queues reached over IPP that list monochrome printing receive the vector PDF and convert it to gray themselves; others receive gray raster sheets.

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

### -LiteralPath
Literal document paths, including FileInfo.FullName.

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
Paper orientation. When omitted, each sheet takes the orientation that places its pages largest, so two-up and six-up sheets of portrait pages print in landscape.

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
Aliases: OutputFilePath
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
Aliases: Pages
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PagesPerSheet
Source pages per sheet, in row-major order.

```yaml
Type: Int32
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values: 1, 2, 4, 6, 9

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PageSubset
Keep all selected pages, or only odd or even original page numbers.

```yaml
Type: String
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values: All, Odd, Even

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PaperSize
Paper size. Defaults to the printer's default media, or A4 when the printer reports none.

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
Paper source identifier reported by Get-OfficePrinterCapability.

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
Local document to print: PDF or any format OfficeIMO converts to PDF.

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
Installed queue from Get-OfficePrinter, or an ipp:// or ipps:// printer URI.

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

### -RasterizeGrayscale
With Grayscale, always send gray raster sheets, so the printout matches the preview pixel for pixel. Slower than letting the printer convert vector PDF. Windows queues always receive raster sheets.

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
