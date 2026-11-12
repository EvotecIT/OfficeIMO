---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficePrinterCapability
## SYNOPSIS
Reports the media, paper sources, duplex, color, resolution and copy capabilities of a printer queue or IPP printer. Requires PowerShell 7.4 or newer.

## SYNTAX
### Default (Default)
```powershell
Get-OfficePrinterCapability [-Name] <string> [<CommonParameters>]
```

## DESCRIPTION
Name accepts an installed queue name or an ipp:// or ipps:// printer URI. Reading capabilities does not change printer settings.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficePrinterCapability -Name 'Office printer'
            Get-OfficePrinterCapability -Name 'ipp://printer.local/ipp/print'
```


## PARAMETERS

### -Name
Printer queue name or IPP printer URI.

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

- `OfficeIMO.Workflows.PdfPrinterCapabilities`

## RELATED LINKS

- None
