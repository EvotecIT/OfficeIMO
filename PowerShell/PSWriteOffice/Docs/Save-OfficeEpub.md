---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Save-OfficeEpub
## SYNOPSIS
Writes a validated EPUB package with explicit conflict and loss policies.

## SYNTAX
### __AllParameterSets
```powershell
Save-OfficeEpub [-Path] <string> -Publication <EpubPublication> [-Force] [-AllowLoss] [-RemoveInvalidatedSignatures] [-MaxOutputBytes <long>] [-ReportVariable <string>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Writes a validated EPUB package with explicit conflict and loss policies.

## EXAMPLES

### EXAMPLE 1
```powershell
$book | Save-OfficeEpub -Path ./operations.epub
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

### -Force
Replace an existing destination.

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
Return the created or changed object.

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
Destination EPUB path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Publication
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

### -RemoveInvalidatedSignatures
Explicitly remove signatures invalidated by package edits. Requires AllowLoss if reported as loss.

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

### -ReportVariable
Variable receiving the writer preservation report.

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

- `OfficeIMO.Epub.EpubPublication`

## OUTPUTS

- `OfficeIMO.Epub.EpubPublication`

## RELATED LINKS

- None
