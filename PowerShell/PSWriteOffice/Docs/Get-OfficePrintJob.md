---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficePrintJob
## SYNOPSIS
Reports the state of a print job submitted with Send-OfficePrinter, optionally waiting until it finishes. Requires PowerShell 7.4 or newer.

## SYNTAX
### Default (Default)
```powershell
Get-OfficePrintJob [-PrinterName] <string> [-JobId] <string> [-Wait] [-TimeoutSeconds <int>] [<CommonParameters>]
```

## DESCRIPTION
States are the same for every transport: Pending, Held, Processing, Stopped, Canceled, Aborted, Completed and Removed. Removed means the queue no longer reports the job: it finished and was purged, was deleted, or never existed.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $receipt = Send-OfficePrinter -Path .\Report.pdf -PrinterName 'Office printer'
$receipt | Get-OfficePrintJob -Wait -TimeoutSeconds 120
```


### EXAMPLE 2
```powershell
PS> Get-OfficePrintJob -PrinterName 'Office printer' -JobId 'Office_printer-42'
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

### -TimeoutSeconds
Longest wait in seconds when Wait is set.

```yaml
Type: Int32
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Wait
Wait until the job is canceled, aborted, completed or removed.

```yaml
Type: SwitchParameter
Parameter Sets: Default
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

- `System.String`

## OUTPUTS

- `OfficeIMO.Workflows.PdfPrintJobStatus`

## RELATED LINKS

- None
