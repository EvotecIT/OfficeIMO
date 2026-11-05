---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Save-OfficeRtf
## SYNOPSIS
Saves an RtfDocument to its original path or to a new path.

## SYNTAX
### Document (Default)
```powershell
Save-OfficeRtf [[-Path] <string>] -Document <RtfDocument> [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Saves an RtfDocument to its original path or to a new path.

## EXAMPLES

### EXAMPLE 1
```powershell
Save-OfficeRtf -Document $rtf.Document -Path .\Copy.rtf
```


## PARAMETERS

### -Document
Document for this operation.

```yaml
Type: RtfDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -PassThru
PassThru for this operation.

```yaml
Type: SwitchParameter
Parameter Sets: Document
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
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Rtf.RtfDocument`

## OUTPUTS

- `System.IO.FileInfo`

## RELATED LINKS

- None
