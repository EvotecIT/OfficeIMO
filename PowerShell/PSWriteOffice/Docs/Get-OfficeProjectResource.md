---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeProjectResource
## SYNOPSIS
Returns project resources.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeProjectResource -Document <ProjectDocument> [-Uid <int[]>] [<CommonParameters>]
```

## DESCRIPTION
Returns project resources.

## EXAMPLES

### EXAMPLE 1
```powershell
$project | Get-OfficeProjectResource
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

### -Uid
Optional stable identities.

```yaml
Type: Int32[]
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

- `OfficeIMO.Project.ProjectResource`

## RELATED LINKS

- None
