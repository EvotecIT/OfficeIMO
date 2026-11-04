---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeRtf
## SYNOPSIS
Creates an RTF document, optionally with paragraphs from -Text or the pipeline.

## SYNTAX
### Live (Default)
```powershell
New-OfficeRtf [[-Path] <string>] [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Text
```powershell
New-OfficeRtf [-Path] <string> [-Text] <string[]> [-Force] [-PassThru] [-Show] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Content
```powershell
New-OfficeRtf [-Path] <string> [-Content] <scriptblock> [-Force] [-PassThru] [-Show] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Creates an RTF document, optionally with paragraphs from -Text or the pipeline.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficeRtf -Path .\Note.rtf -Text 'Server maintenance starts at 22:00.'
```


## PARAMETERS

### -Content
Content for this operation.

```yaml
Type: ScriptBlock
Parameter Sets: Content
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Force
Force for this operation.

```yaml
Type: SwitchParameter
Parameter Sets: Live, Text, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
PassThru for this operation.

```yaml
Type: SwitchParameter
Parameter Sets: Text, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Path for this operation.

```yaml
Type: String
Parameter Sets: Live, Text, Content
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Show
Show for this operation.

```yaml
Type: SwitchParameter
Parameter Sets: Text, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Text for this operation.

```yaml
Type: String[]
Parameter Sets: Text
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String[]`

## OUTPUTS

- `OfficeIMO.Rtf.RtfDocument`

## RELATED LINKS

- None
