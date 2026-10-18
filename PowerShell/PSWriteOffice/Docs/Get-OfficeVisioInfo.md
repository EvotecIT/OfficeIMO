---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeVisioInfo
## SYNOPSIS
Creates a deterministic inspection snapshot for a Visio document.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeVisioInfo [-Path] <string> [-AsText] [<CommonParameters>]
```

### Document
```powershell
Get-OfficeVisioInfo -Document <VisioDocument> [-AsText] [<CommonParameters>]
```

## DESCRIPTION
Creates a deterministic inspection snapshot for a Visio document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> OfficeVisio -Path .\ServiceMap.vsdx { VisioRectangle -Text 'API' -X 2 -Y 4 }
            Get-OfficeVisioInfo -Path .\ServiceMap.vsdx -AsText
```

Returns stable line-oriented text that is useful for tests and release notes.

## PARAMETERS

### -AsText
Emit the stable line-oriented inspection text.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Visio document object.

```yaml
Type: VisioDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Path
Visio .vsdx path.

```yaml
Type: String
Parameter Sets: Path
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
- `OfficeIMO.Visio.VisioDocument`

## OUTPUTS

- `OfficeIMO.Visio.VisioInspectionSnapshot`
- `System.String`

## RELATED LINKS

- None
