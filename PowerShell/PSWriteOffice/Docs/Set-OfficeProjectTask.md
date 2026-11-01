---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficeProjectTask
## SYNOPSIS
Updates selected task fields without implicitly recalculating its schedule.

## SYNTAX
### Task (Default)
```powershell
Set-OfficeProjectTask -Task <ProjectTask> [-Name <string>] [-DurationDays <Decimal>] [-Start <DateTime>] [-Finish <DateTime>] [-PercentComplete <Int32>] [-Manual <Boolean>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Updates selected task fields without implicitly recalculating its schedule.

## EXAMPLES

### EXAMPLE 1
```powershell
$project | Get-OfficeProjectTask -Uid 7 | Set-OfficeProjectTask -DurationDays 4 -PercentComplete 25
```


## PARAMETERS

### -DurationDays
Working duration in days.

```yaml
Type: Decimal
Parameter Sets: Task
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Finish
Stored local finish date.

```yaml
Type: DateTime
Parameter Sets: Task
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Manual
Explicit manual scheduling choice.

```yaml
Type: Boolean
Parameter Sets: Task
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Name
New task name.

```yaml
Type: String
Parameter Sets: Task
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Return the changed task.

```yaml
Type: SwitchParameter
Parameter Sets: Task
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PercentComplete
Duration completion from zero to one hundred.

```yaml
Type: Int32
Parameter Sets: Task
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Start
Stored local start date.

```yaml
Type: DateTime
Parameter Sets: Task
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Task
Editable task returned by Get-OfficeProjectTask.

```yaml
Type: ProjectTask
Parameter Sets: Task
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

- `OfficeIMO.Project.ProjectTask`

## OUTPUTS

- `OfficeIMO.Project.ProjectTask`

## RELATED LINKS

- None
