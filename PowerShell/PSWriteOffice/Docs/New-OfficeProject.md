---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeProject
## SYNOPSIS
Creates a live project with an optional standard working calendar.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeProject [[-Name] <string>] [-StartDate <DateTime>] [-NoStandardCalendar] [<CommonParameters>]
```

## DESCRIPTION
Creates a live project with an optional standard working calendar.

## EXAMPLES

### EXAMPLE 1
```powershell
$project = New-OfficeProject -Name Delivery -StartDate '2026-10-20 08:00'
```


## PARAMETERS

### -Name
Project name.

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

### -NoStandardCalendar
Do not create a Monday-to-Friday working calendar.

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

### -StartDate
Explicit project start; omitted dates remain unspecified.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `OfficeIMO.Project.ProjectDocument`

## RELATED LINKS

- None
