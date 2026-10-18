---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeProtectionCapability
## SYNOPSIS
Returns OfficeIMO's machine-readable protected-content support contract.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeProtectionCapability [[-Id] <string>] [-Format <string>] [-Kind <OfficeProtectionKind>] [-IncompleteOnly] [-AsJson] [-AsMarkdown] [<CommonParameters>]
```

## DESCRIPTION
Returns OfficeIMO's machine-readable protected-content support contract.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficeProtectionCapability -IncompleteOnly | Format-Table FormatId, Kind, Open, Create, Validate
```

Shows formats whose encrypted, signed, restricted, or obfuscated workflows still have unsupported operations.

## PARAMETERS

### -AsJson
Return the complete catalog as deterministic JSON. Filtering parameters cannot be combined with this switch.

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

### -AsMarkdown
Return the complete catalog as a Markdown table. Filtering parameters cannot be combined with this switch.

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

### -Format
Filter by format identifier or family, such as PDF, DOCX, or EML.

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

### -Id
Exact stable capability identifier.

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

### -IncompleteOnly
Return only rows with at least one unsupported operation.

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

### -Kind
Filter by protected-content mechanism.

```yaml
Type: OfficeProtectionKind
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: PasswordEncryption, RecipientEncryption, DigitalSignature, FontObfuscation, EditingRestriction, AccessDeterrence

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

- `OfficeIMO.Security.OfficeProtectionCapability`
- `System.String`

## RELATED LINKS

- None
