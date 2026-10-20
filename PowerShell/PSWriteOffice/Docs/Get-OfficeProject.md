---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeProject
## SYNOPSIS
Loads a project without calculating or executing its content.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeProject [-Path] <string> [-ReadOnly] [-MaxInputBytes <long>] [<CommonParameters>]
```

## DESCRIPTION
Loads a project without calculating or executing its content.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-OfficeProject -Path ./delivery.mpp -ReadOnly
```


## PARAMETERS

### -MaxInputBytes
Maximum source size in bytes.

```yaml
Type: Int64
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Literal filesystem path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: FullName, LiteralPath, FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue, ByPropertyName)
Accept wildcard characters: False
```

### -ReadOnly
Open without allowing edits.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String`

## OUTPUTS

- `OfficeIMO.Project.ProjectDocument`

## RELATED LINKS

- None
