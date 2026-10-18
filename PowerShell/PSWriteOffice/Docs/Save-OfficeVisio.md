---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Save-OfficeVisio
## SYNOPSIS
Saves an OfficeIMO.Visio document.

## SYNTAX
### __AllParameterSets
```powershell
Save-OfficeVisio [-Document] <VisioDocument> [[-Path] <string>] [-Open] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Saves an OfficeIMO.Visio document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $diagram = Get-OfficeVisio -Path .\ServiceMap.vsdx
$diagram | Save-OfficeVisio -Path .\ServiceMap-copy.vsdx -PassThru
```

Saves an existing OfficeIMO.Visio document to another .vsdx file.

## PARAMETERS

### -Document
Visio document to save.

```yaml
Type: VisioDocument
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Open
Open the document after saving.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: Show
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the document object for further processing.

```yaml
Type: SwitchParameter
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
Optional save-as path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: FilePath
Possible values:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Visio.VisioDocument`

## OUTPUTS

- `OfficeIMO.Visio.VisioDocument`

## RELATED LINKS

- None
