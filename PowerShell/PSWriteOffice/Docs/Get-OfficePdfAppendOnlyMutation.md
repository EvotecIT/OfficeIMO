---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficePdfAppendOnlyMutation
## SYNOPSIS
Gets append-only PDF mutation support and blockers for an existing PDF.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficePdfAppendOnlyMutation [-Path] <string> [-Password <string>] [-IgnorePermissionRestrictions] [<CommonParameters>]
```

## DESCRIPTION
Gets append-only PDF mutation support and blockers for an existing PDF.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $plan = Get-OfficePdfAppendOnlyMutation -Path .\SignedOrReviewed.pdf
$plan.CanAppendMetadata
$plan.Blockers
```

Returns OfficeIMO.Pdf append-only mutation support and blocker details.

## PARAMETERS

### -IgnorePermissionRestrictions
After successful password authentication, explicitly ignore owner-imposed usage restrictions.

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

### -Password
Password used to authenticate an encrypted PDF.

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

### -Path
PDF file path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String`

## OUTPUTS

- `OfficeIMO.Pdf.PdfAppendOnlyMutationReport`

## RELATED LINKS

- None
