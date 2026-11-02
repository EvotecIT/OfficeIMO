---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Save-OfficeProject
## SYNOPSIS
Saves a project through the qualified native codec with loss protection.

## SYNTAX
### Document (Default)
```powershell
Save-OfficeProject [-Path] <string> -Document <ProjectDocument> [-Format <ProjectFileFormat>] [-AllowLoss] [-Force] [-ReportVariable <string>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Saves a project through the qualified native codec with loss protection.

## EXAMPLES

### EXAMPLE 1
```powershell
$project | Save-OfficeProject -Path ./delivery.mpp
```


## PARAMETERS

### -AllowLoss
Permit reported conversion loss. Unsupported codecs still fail.

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

### -Document
Project to inspect or change.

```yaml
Type: ProjectDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Force
Replace an existing destination.

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

### -Format
Explicit generation or exchange format.

```yaml
Type: ProjectFileFormat
Parameter Sets: Document
Aliases: None
Possible values: Automatic, Xml, Mpp14, Mpt14, Mpp12, Mpt12, Mpp9, Mpt9, Mpp8, Mpt8, Mpx4

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
Destination filename; its extension selects the format.

```yaml
Type: String
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ReportVariable
Variable receiving the preservation and conversion assessment.

```yaml
Type: String
Parameter Sets: Document
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

- `OfficeIMO.Project.ProjectDocument`

## OUTPUTS

- `OfficeIMO.Project.ProjectDocument`

## RELATED LINKS

- None
