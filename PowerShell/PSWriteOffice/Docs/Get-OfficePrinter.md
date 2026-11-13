---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficePrinter
## SYNOPSIS
Lists system printer queues, or the paper sources of a named queue. Requires PowerShell 7.4 or newer.

## SYNTAX
### Printer (Default)
```powershell
Get-OfficePrinter [[-Name] <string[]>] [<CommonParameters>]
```

### PaperSource
```powershell
Get-OfficePrinter [-Name] <string[]> -PaperSource [<CommonParameters>]
```

## DESCRIPTION
Use Get-OfficePrinterCapability for media, duplex, color and resolution details of a queue or IPP printer.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficePrinter
            Get-OfficePrinter -Name 'Office printer' -PaperSource
```


## PARAMETERS

### -Name
Printer name patterns; PaperSource requires one exact printer name or IPP printer URI.

```yaml
Type: String[]
Parameter Sets: Printer, PaperSource
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PaperSource
List paper sources of the single printer selected by Name.

```yaml
Type: SwitchParameter
Parameter Sets: PaperSource
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

- `None`

## OUTPUTS

- `OfficeIMO.Workflows.PdfPrinterInfo`
- `OfficeIMO.Workflows.PdfPaperSourceInfo`

## RELATED LINKS

- None
