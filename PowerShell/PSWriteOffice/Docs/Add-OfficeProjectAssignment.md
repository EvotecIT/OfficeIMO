---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeProjectAssignment
## SYNOPSIS
Assigns a project resource to a task by stable UID.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeProjectAssignment -Document <ProjectDocument> -TaskUid <int> -ResourceUid <int> [-UnitsPercent <decimal>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Assigns a project resource to a task by stable UID.

## EXAMPLES

### EXAMPLE 1
```powershell
$project | Add-OfficeProjectAssignment -TaskUid 1 -ResourceUid 1 -UnitsPercent 50
```


## PARAMETERS

### -Document
Project to inspect or change.

```yaml
Type: ProjectDocument
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
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

### -ResourceUid
Stable resource UID.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TaskUid
Stable task UID.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -UnitsPercent
Assignment units as a percentage.

```yaml
Type: Decimal
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

- `OfficeIMO.Project.ProjectDocument`

## OUTPUTS

- `OfficeIMO.Project.ProjectAssignment`

## RELATED LINKS

- None
