---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficePackageSecurity
## SYNOPSIS
Inspects an Open XML or compound Office package without opening active content.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficePackageSecurity [-Path] <string> [-Options <OfficePackageSecurityOptions>] [-Untrusted] [-ThrowOnViolation] [<CommonParameters>]
```

## DESCRIPTION
Inspects an Open XML or compound Office package without opening active content.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $report = Get-OfficePackageSecurity -Path .\Incoming\Report.xlsm -Untrusted
$report | Select-Object IsValid, MacroPartCount, EmbeddedPayloadPartCount, ExternalRelationshipCount
$report.Findings | Format-Table Severity, Rule, PartName, Message
```

Returns observations and policy violations; it does not execute package content.

## PARAMETERS

### -Options
Custom package size, expansion, and active-content policy.

```yaml
Type: OfficePackageSecurityOptions
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
Path to an Open XML or compound Office package.

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

### -ThrowOnViolation
Throw on the first policy violation instead of returning only the report.

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

### -Untrusted
Use the bounded policy that rejects macros, embedded payloads, ActiveX, and external relationships.

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

- `OfficeIMO.OfficePackageSecurityReport`

## RELATED LINKS

- None
