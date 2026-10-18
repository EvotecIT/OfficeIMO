---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficePrinter
## SYNOPSIS
Lists system printer queues, or the paper sources of a named queue. Requires PowerShell 7.4 or newer.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficePrinter [-PaperSources <string>] [<CommonParameters>]
```

## DESCRIPTION
Lists system printer queues, or the paper sources of a named queue. Requires PowerShell 7.4 or newer.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficePrinter
            Get-OfficePrinter -PaperSources 'Office printer'
```


## PARAMETERS

### -PaperSources
Named printer whose paper sources should be listed.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `OfficeIMO.Workflows.PdfPrinterInfo`
- `OfficeIMO.Workflows.PdfPaperSourceInfo`

## RELATED LINKS

- None
