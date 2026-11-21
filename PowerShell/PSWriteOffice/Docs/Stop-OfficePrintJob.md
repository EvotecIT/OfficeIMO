---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Stop-OfficePrintJob
## SYNOPSIS
Cancels a print job submitted with Send-OfficePrinter and returns its state afterwards. Requires PowerShell 7.4 or newer.

## SYNTAX
### Default (Default)
```powershell
Stop-OfficePrintJob [-PrinterName] <string> [-JobId] <string> [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
A job that has already finished is reported unchanged. Pages the printer has already received may still print.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Send-OfficePrinter -Path .\Report.pdf -PrinterName 'Office printer' | Stop-OfficePrintJob
```


## PARAMETERS

### -JobId
The job identifier from the receipt.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -PrinterName
The printer queue name or IPP URI from the receipt.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String`

## OUTPUTS

- `OfficeIMO.Workflows.PdfPrintJobStatus`

## RELATED LINKS

- None
