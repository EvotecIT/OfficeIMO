---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeProjectView
## SYNOPSIS
Captures a portable report from a current calculated schedule.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeProjectView -Document <ProjectDocument> [-Schedule <ProjectScheduleResult>] [-Kind <ProjectViewKind>] [-Options <ProjectViewOptions>] [<CommonParameters>]
```

## DESCRIPTION
Captures a portable report from a current calculated schedule.

## EXAMPLES

### EXAMPLE 1
```powershell
$view = $project | Get-OfficeProjectView -Schedule $schedule -Kind Gantt
```


## PARAMETERS

### -Document
Project document.

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

### -Kind
Report presentation.

```yaml
Type: ProjectViewKind
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Gantt, TaskUsage, ResourceUsage, ResourceHistogram, Network, Timeline, Table

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Options
Optional explicit report settings.

```yaml
Type: ProjectViewOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Schedule
Current schedule result for this project revision. Omission calculates a report snapshot without changing the document.

```yaml
Type: ProjectScheduleResult
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

- `OfficeIMO.Project.ProjectView`

## RELATED LINKS

- None
