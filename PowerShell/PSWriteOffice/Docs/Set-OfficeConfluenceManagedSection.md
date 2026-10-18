---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeConfluenceManagedSection
## SYNOPSIS
Safely replaces one marker-delimited section in a Confluence storage body.

## SYNTAX
### __AllParameterSets
```powershell
Set-OfficeConfluenceManagedSection -ExistingBody <string> -SectionId <string> -Replacement <string> [-AppendIfMissing] [-PassThruBody] [<CommonParameters>]
```

## DESCRIPTION
Safely replaces one marker-delimited section in a Confluence storage body.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $result = Set-OfficeConfluenceManagedSection -ExistingBody $storage -SectionId daily-report -Replacement $html -AppendIfMissing
```

Returns before/after hashes and the updated body without contacting Confluence.

## PARAMETERS

### -AppendIfMissing
Append a new marker pair when the section does not exist.

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

### -ExistingBody
Existing Confluence storage body.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -PassThruBody
Return only the updated body string.

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

### -Replacement
Replacement storage-format content.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SectionId
Stable marker identifier.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
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

- `OfficeIMO.Confluence.ConfluenceManagedSectionResult`
- `System.String`

## RELATED LINKS

- None
