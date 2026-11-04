---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeProjectViewOptions
## SYNOPSIS
Creates selection, timescale and page settings for Get-OfficeProjectView -Options.

## SYNTAX
### Options (Default)
```powershell
New-OfficeProjectViewOptions [-Kind <ProjectViewKind>] [-Timescale <ProjectViewTimescale>] [-Task <int[]>] [-Resource <int[]>] [-Column <ProjectViewColumn[]>] [-NameContains <string>] [-CriticalOnly] [-NoSummaries] [-Grouping <ProjectViewGrouping>] [-NetworkLayout <ProjectNetworkLayout>] [-StatusDate <DateTime>] [-Baseline <Int32>] [-Start <DateTime>] [-Finish <DateTime>] [-NoProgress] [-NoLegend] [-PageWidth <OfficeLength>] [-PageHeight <OfficeLength>] [-Margin <OfficeLength>] [-NoFitToContent] [-MaximumRows <Int32>] [-MaximumBuckets <Int32>] [-MaximumCells <Int32>] [-MaximumIntervalVisits <Int32>] [-MaximumPages <Int32>] [<CommonParameters>]
```

## DESCRIPTION
Creates selection, timescale and page settings for Get-OfficeProjectView -Options.

## EXAMPLES

### EXAMPLE 1
```powershell
$options = New-OfficeProjectViewOptions -Kind ResourceUsage -Timescale Month
Get-OfficeProjectView -Document $project -Options $options
```


## PARAMETERS

### -Baseline
Baseline number 0-10 to compare against.

```yaml
Type: Int32
Parameter Sets: Options
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Column
Uid, Name, Start, Finish, WorkHours, Cost, PercentComplete, Critical; default depends on the layout.

```yaml
Type: ProjectViewColumn[]
Parameter Sets: Options
Aliases: None
Possible values: Uid, Name, Start, Finish, WorkHours, Cost, PercentComplete, Critical

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CriticalOnly
CriticalOnly for this operation.

```yaml
Type: SwitchParameter
Parameter Sets: Options
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Finish
Last visible date (exclusive); default from the schedule.

```yaml
Type: DateTime
Parameter Sets: Options
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Grouping
None (outline order, default) or ParentTask.

```yaml
Type: ProjectViewGrouping
Parameter Sets: Options
Aliases: None
Possible values: None, ParentTask

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Kind
Native report kind; defaults to Gantt.

```yaml
Type: ProjectViewKind
Parameter Sets: Options
Aliases: None
Possible values: Gantt, TaskUsage, ResourceUsage, ResourceHistogram, Network, Timeline, Table

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Margin
Default 28pt.

```yaml
Type: OfficeLength
Parameter Sets: Options
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumBuckets
Time buckets before pagination; default 1000.

```yaml
Type: Int32
Parameter Sets: Options
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumCells
Default 1000000.

```yaml
Type: Int32
Parameter Sets: Options
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumIntervalVisits
Default 2000000.

```yaml
Type: Int32
Parameter Sets: Options
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumPages
Default 1000.

```yaml
Type: Int32
Parameter Sets: Options
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumRows
Default 10000.

```yaml
Type: Int32
Parameter Sets: Options
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NameContains
Case-insensitive task name filter.

```yaml
Type: String
Parameter Sets: Options
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NetworkLayout
Network view: Dependency (default) or Compact.

```yaml
Type: ProjectNetworkLayout
Parameter Sets: Options
Aliases: None
Possible values: Dependency, Compact

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoFitToContent
Keep the full -PageHeight instead of trimming unused space (trimmed by default).

```yaml
Type: SwitchParameter
Parameter Sets: Options
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoLegend
NoLegend for this operation.

```yaml
Type: SwitchParameter
Parameter Sets: Options
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoProgress
Do not fill task bars with completion (shown by default).

```yaml
Type: SwitchParameter
Parameter Sets: Options
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoSummaries
Omit summary task rows (shown by default).

```yaml
Type: SwitchParameter
Parameter Sets: Options
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PageHeight
Default 595pt; the maximum height when pages fit their content.

```yaml
Type: OfficeLength
Parameter Sets: Options
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PageWidth
Default A4 landscape (842pt).

```yaml
Type: OfficeLength
Parameter Sets: Options
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Resource
Stable native resource UIDs. Pass the Uid values of selected model objects; these are identities, not collection positions.

```yaml
Type: Int32[]
Parameter Sets: Options
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Start
First visible date (inclusive); default from the schedule.

```yaml
Type: DateTime
Parameter Sets: Options
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StatusDate
Status-date marker in Gantt and Timeline views.

```yaml
Type: DateTime
Parameter Sets: Options
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Task
Stable native task UIDs. Pass the Uid values of selected model objects; these are identities, not collection positions.

```yaml
Type: Int32[]
Parameter Sets: Options
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Timescale
Day, Week (default) or Month buckets; weeks begin on Monday.

```yaml
Type: ProjectViewTimescale
Parameter Sets: Options
Aliases: None
Possible values: Day, Week, Month

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `OfficeIMO.Project.ProjectViewOptions`

## RELATED LINKS

- None
