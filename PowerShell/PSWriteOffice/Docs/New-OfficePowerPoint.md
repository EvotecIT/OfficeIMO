---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficePowerPoint
## SYNOPSIS
Creates a PowerPoint presentation using the DSL.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficePowerPoint [-Path] <string> [[-Content] <scriptblock>] [-Open] [-NoSave] [-PassThru] [-Password <string>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Initializes a presentation, runs the DSL script block, and optionally saves the deck.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $ppt = New-OfficePowerPoint -Path .\deck.pptx -NoSave
```

Creates a live presentation associated with deck.pptx for incremental composition.

### EXAMPLE 2
```powershell
PS> New-OfficePowerPoint -Path .\deck.pptx { PptSlide { PptTitle -Title 'Status Update' } } -Open
```

Creates, saves, and opens a deck with one titled slide.

## PARAMETERS

### -Content
DSL scriptblock describing presentation content.

```yaml
Type: ScriptBlock
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoSave
Skip saving after executing the DSL.

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

### -Open
Open the presentation after saving.

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

### -PassThru
Emit the saved FileInfo for chaining.

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

### -Password
Password used to save the presentation as an encrypted package.

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
Destination path for the new .pptx.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: FilePath
Possible values:

Required: True
Position: 0
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
