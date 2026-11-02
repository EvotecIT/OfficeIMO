---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# ConvertTo-OfficeProjectTable
## SYNOPSIS
Maps PowerShell rows to a bounded project entity table.

## SYNTAX
### InputObject (Default)
```powershell
ConvertTo-OfficeProjectTable -InputObject <psobject[]> -Kind <ProjectDataKind> -ColumnMap <IDictionary> [-MaxRows <int>] [<CommonParameters>]
```

## DESCRIPTION
Maps PowerShell rows to a bounded project entity table.

## EXAMPLES

### EXAMPLE 1
```powershell
Import-Csv ./tasks.csv | ConvertTo-OfficeProjectTable -Kind Tasks -ColumnMap @{ Uid = 'ID'; Name = 'Task' }
```


## PARAMETERS

### -ColumnMap
Project field name to source column header. Field names are discovered through ProjectDataField.

```yaml
Type: IDictionary
Parameter Sets: InputObject
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputObject
Rows from Import-Csv, dictionaries or PowerShell objects.

```yaml
Type: PSObject[]
Parameter Sets: InputObject
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Kind
Entity collection represented by the rows.

```yaml
Type: ProjectDataKind
Parameter Sets: InputObject
Aliases: None
Possible values: Tasks, Resources, Assignments, Calendars

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxRows
Maximum accumulated rows.

```yaml
Type: Int32
Parameter Sets: InputObject
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

- `System.Management.Automation.PSObject[]`

## OUTPUTS

- `OfficeIMO.Project.ProjectMappedTable`

## RELATED LINKS

- None
