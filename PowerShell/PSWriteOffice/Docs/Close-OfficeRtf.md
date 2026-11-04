---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Close-OfficeRtf
## SYNOPSIS
Releases an RtfDocument without saving.

## SYNTAX
### Document (Default)
```powershell
Close-OfficeRtf -Document <RtfDocument> [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Releases an RtfDocument without saving.

## EXAMPLES

### EXAMPLE 1
```powershell
Close-OfficeRtf -Document $rtf.Document
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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Rtf.RtfDocument`

## OUTPUTS

- `System.Void`

## RELATED LINKS

- None
