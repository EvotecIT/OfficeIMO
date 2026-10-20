---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Invoke-OfficeProjectSchedule
## SYNOPSIS
Explicitly calculates a schedule and optionally applies its dates.

## SYNTAX
### __AllParameterSets
```powershell
Invoke-OfficeProjectSchedule -Document <ProjectDocument> [-CalculateAssignments] [-Apply] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Explicitly calculates a schedule and optionally applies its dates.

## EXAMPLES

### EXAMPLE 1
```powershell
$project | Invoke-OfficeProjectSchedule -CalculateAssignments -Apply
```


## PARAMETERS

### -Apply
Apply an error-free result to this project. Otherwise return a calculation without changing stored dates.

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

### -CalculateAssignments
Also calculate assignment work and costs.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Project.ProjectDocument`

## OUTPUTS

- `OfficeIMO.Project.ProjectScheduleResult`

## RELATED LINKS

- None
