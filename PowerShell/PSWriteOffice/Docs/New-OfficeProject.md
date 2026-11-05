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
### Live (Default)
```powershell
New-OfficeProject [[-Path] <string>] [-Force] [-Name <string>] [-StartDate <DateTime>] [-NoStandardCalendar] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Content
```powershell
New-OfficeProject [-Path] <string> [-Content] <scriptblock> [-AllowLoss] [-PassThru] [-Force] [-Name <string>] [-StartDate <DateTime>] [-NoStandardCalendar] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Creates a live project with an optional standard working calendar.

## EXAMPLES

### EXAMPLE 1
```powershell
$project = New-OfficeProject -Name Delivery -StartDate '2026-10-20 08:00'
```


## PARAMETERS

### -AllowLoss
Accept reductions reported by native save.

```yaml
Type: SwitchParameter
Parameter Sets: Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Content
Authoring block completed before atomic publication.

```yaml
Type: ScriptBlock
Parameter Sets: Content
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Force
Permit replacement of the associated destination.

```yaml
Type: SwitchParameter
Parameter Sets: Live, Content
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
Parameter Sets: Live, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoStandardCalendar
Do not create a Monday-to-Friday working calendar.

```yaml
Type: SwitchParameter
Parameter Sets: Live, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Return the saved FileInfo after the block completes.

```yaml
Type: SwitchParameter
Parameter Sets: Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Optional live association; required for a completed content block.

```yaml
Type: String
Parameter Sets: Live, Content
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StartDate
Explicit project start; omitted dates remain unspecified.

```yaml
Type: DateTime
Parameter Sets: Live, Content
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
