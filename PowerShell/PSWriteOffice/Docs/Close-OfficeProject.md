---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Close-OfficeProject
## SYNOPSIS
Closes a project without implicitly saving it.

## SYNTAX
### __AllParameterSets
```powershell
Close-OfficeProject -Document <ProjectDocument> [<CommonParameters>]
```

## DESCRIPTION
Closes a project without implicitly saving it.

## EXAMPLES

### EXAMPLE 1
```powershell
$project | Close-OfficeProject
```


## PARAMETERS

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

- `None`

## RELATED LINKS

- None
