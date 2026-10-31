---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Save-OfficeEpub
## SYNOPSIS
Saves a native EPUB publication, rejecting loss unless explicitly allowed.

## SYNTAX
### __AllParameterSets
```powershell
Save-OfficeEpub [[-Path] <string>] -Document <EpubPublication> [-AllowLoss] [-RemoveInvalidatedSignatures] [-MaxOutputBytes <long>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Save replaces its destination after validation. PassThru returns the native write report.

## EXAMPLES

### EXAMPLE 1
```powershell
$report = $book | Save-OfficeEpub -PassThru
```


## PARAMETERS

### -AllowLoss
Permit reductions reported by the package writer.

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

### -Document
Live editable publication.

```yaml
Type: EpubPublication
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -MaxOutputBytes
Maximum compressed output size.

```yaml
Type: Int64
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
Return the native write report.

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
Destination path; omission uses the native publication's associated path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RemoveInvalidatedSignatures
Explicitly remove signatures invalidated by edits. Reported loss still requires AllowLoss.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Epub.EpubPublication`

## OUTPUTS

- `OfficeIMO.Epub.EpubWriteReport`

## RELATED LINKS

- None
