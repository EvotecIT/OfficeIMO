---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Import-OfficeProjectTable
## SYNOPSIS
Imports explicitly mapped tables into a new project without changing existing documents.

## SYNTAX
### __AllParameterSets
```powershell
Import-OfficeProjectTable -Table <ProjectMappedTable[]> [-Name <string>] [-StartDate <DateTime>] [-CalendarUid <Int32>] [-AllowUnmappedColumns] [-MaxRows <int>] [-ReportVariable <string>] [<CommonParameters>]
```

## DESCRIPTION
Imports explicitly mapped tables into a new project without changing existing documents.

## EXAMPLES

### EXAMPLE 1
```powershell
$taskTable | Import-OfficeProjectTable -Name Delivery -ReportVariable importReport
```


## PARAMETERS

### -AllowUnmappedColumns
Explicitly permit ignored source columns and return notices.

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

### -CalendarUid
Imported project calendar UID.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxRows
Maximum rows across all tables.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Name
Project name.

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

### -ReportVariable
Variable receiving the import result and notices.

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

### -StartDate
Explicit local start date.

```yaml
Type: DateTime
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Table
At most one table for each entity kind.

```yaml
Type: ProjectMappedTable[]
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

- `OfficeIMO.Project.ProjectMappedTable[]`

## OUTPUTS

- `OfficeIMO.Project.ProjectDocument`

## RELATED LINKS

- None
