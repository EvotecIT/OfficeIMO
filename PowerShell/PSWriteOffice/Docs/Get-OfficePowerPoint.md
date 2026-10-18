---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficePowerPoint
## SYNOPSIS
Loads an existing PowerPoint presentation.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficePowerPoint -Path <string> [-Password <string>] [<CommonParameters>]
```

## DESCRIPTION
Returns an OfficeIMO PowerPointPresentation for downstream slide operations.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $ppt = Get-OfficePowerPoint -Path .\Quarterly.pptx
```

Reads Quarterly.pptx and exposes the presentation object.

## PARAMETERS

### -Password
Password used to open an encrypted presentation package.

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
Path to the .pptx file.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: FilePath
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

- `None`

## RELATED LINKS

- None
