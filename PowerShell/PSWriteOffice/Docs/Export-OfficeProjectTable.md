---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Export-OfficeProjectTable
## SYNOPSIS
Returns mapped project tables and explicit projection-loss notices.

## SYNTAX
### __AllParameterSets
```powershell
Export-OfficeProjectTable -Document <ProjectDocument> [-AllowLossyProjection] [-MaxRows <int>] [-MaxCells <int>] [<CommonParameters>]
```

## DESCRIPTION
Returns mapped project tables and explicit projection-loss notices.

## EXAMPLES

### EXAMPLE 1
```powershell
$tables = $project | Export-OfficeProjectTable -AllowLossyProjection
```


## PARAMETERS

### -AllowLossyProjection
Explicitly accept omitted settings, dependencies and advanced/native data.

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

### -MaxCells
Maximum total cells.

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
Maximum total rows.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Project.ProjectDocument`

## OUTPUTS

- `OfficeIMO.Project.ProjectDataExportResult`

## RELATED LINKS

- None
