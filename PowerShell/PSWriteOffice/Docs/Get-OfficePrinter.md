---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficePrinter
## SYNOPSIS
Lists system printer queues, or the capabilities or paper sources of a named queue. Requires PowerShell 7.4 or newer.

## SYNTAX
### Printer (Default)
```powershell
Get-OfficePrinter [[-Name] <string[]>] [<CommonParameters>]
```

### PaperSource
```powershell
Get-OfficePrinter [-Name] <string[]> -PaperSource [<CommonParameters>]
```

### Capability
```powershell
Get-OfficePrinter [-Name] <string[]> -Capability [<CommonParameters>]
```

## DESCRIPTION
Lists system printer queues, or the capabilities or paper sources of a named queue. Requires PowerShell 7.4 or newer.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficePrinter
            Get-OfficePrinter -Name 'Office printer' -Capability
            Get-OfficePrinter -Name 'Office printer' -PaperSource
```


## PARAMETERS

### -Capability
Return the media, trays, duplex, color, resolution, collation and copy capabilities of the single printer selected by Name.

```yaml
Type: SwitchParameter
Parameter Sets: Capability
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Name
Printer name patterns; Capability and PaperSource require one exact printer name.

```yaml
Type: String[]
Parameter Sets: Printer, PaperSource, Capability
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
- `OfficeIMO.Workflows.PdfPrinterCapabilities`

## RELATED LINKS

- None
