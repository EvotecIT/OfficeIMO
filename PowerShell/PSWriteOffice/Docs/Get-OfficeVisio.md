---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeVisio
## SYNOPSIS
Loads an existing .vsdx file as an OfficeIMO.Visio document.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeVisio [-Path] <string> [<CommonParameters>]
```

## DESCRIPTION
Loads an existing .vsdx file as an OfficeIMO.Visio document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $diagram = Get-OfficeVisio -Path .\ServiceMap.vsdx
Get-OfficeVisioInfo -Document $diagram -AsText
```

Loads an existing .vsdx file and creates a deterministic inspection snapshot.

## PARAMETERS

### -Path
Visio .vsdx path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String`

## OUTPUTS

- `OfficeIMO.Visio.VisioDocument`

## RELATED LINKS

- None
