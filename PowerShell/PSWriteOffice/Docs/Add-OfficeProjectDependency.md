---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeProjectDependency
## SYNOPSIS
Adds a finish-to-start relationship between two tasks.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeProjectDependency -Document <ProjectDocument> -PredecessorUid <int> -SuccessorUid <int> [-Type <ProjectDependencyType>] [-LagDays <decimal>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a finish-to-start relationship between two tasks.

## EXAMPLES

### EXAMPLE 1
```powershell
$project | Add-OfficeProjectDependency -PredecessorUid 1 -SuccessorUid 2
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

### -LagDays
Working-day delay; negative values represent lead time.

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

### -PredecessorUid
Predecessor task UID.

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

### -SuccessorUid
Successor task UID.

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

### -Type
Dependency relationship.

```yaml
Type: ProjectDependencyType
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: FinishToFinish, FinishToStart, StartToFinish, StartToStart

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

- `OfficeIMO.Project.ProjectDependency`

## RELATED LINKS

- None
