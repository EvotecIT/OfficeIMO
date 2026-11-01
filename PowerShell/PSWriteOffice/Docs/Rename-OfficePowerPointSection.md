---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Rename-OfficePowerPointSection
## SYNOPSIS
Renames the first section matching the supplied name.

## SYNTAX
### Document (Default)
```powershell
Rename-OfficePowerPointSection [-Name] <string> [-NewName] <string> [-Document <PowerPointDocument>] [-CaseSensitive] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Renames the first section matching the supplied name.

## EXAMPLES

### EXAMPLE 1
```powershell
Rename-OfficePowerPointSection -Document $deck -Name Results -NewName 'Deep dive'
```


### EXAMPLE 2
```powershell
$section = Rename-OfficePowerPointSection -Document $deck -Name Appendix -NewName Reference -PassThru
```


## PARAMETERS

### -CaseSensitive
Gets or sets whether name matching is case-sensitive.

```yaml
Type: SwitchParameter
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Gets or sets the live document; omission uses its enclosing content block.

```yaml
Type: PowerPointDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Name
Gets or sets the existing section name.

```yaml
Type: String
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NewName
Gets or sets the new section name.

```yaml
Type: String
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Gets or sets whether the changed section is returned.

```yaml
Type: SwitchParameter
Parameter Sets: Document
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

- `OfficeIMO.PowerPoint.PowerPointDocument`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointSectionInfo`

## RELATED LINKS

- None
