---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeExcelProtectionOptions
## SYNOPSIS
Creates Excel-specific hash and reservation settings for Protect-OfficeDocument.

## SYNTAX
### Default (Default)
```powershell
New-OfficeExcelProtectionOptions [-LegacyPasswordHash <string>] [-ModifyPasswordHash <string>] [-ReservedBy <string>] [<CommonParameters>]
```

## DESCRIPTION
Creates Excel-specific hash and reservation settings for Protect-OfficeDocument.

## EXAMPLES

### EXAMPLE 1
```powershell
$options = New-OfficeExcelProtectionOptions -ReservedBy 'Finance team'
Protect-OfficeDocument -Path ./Budget.xlsx -Restriction ReadOnlyRecommended -Options $options -OutputPath ./Budget-review.xlsx
```


### EXAMPLE 2
```powershell
$modify = Read-Host -AsSecureString -Prompt 'Password to modify'
Protect-OfficeDocument -Path .\Finance\Budget-2027.xlsx -Restriction ReadOnlyRecommended -RestrictionPassword $modify -Options (New-OfficeExcelProtectionOptions -ReservedBy 'Finance team')
```

Recommends read-only opening and records who reserved the workbook.

### EXAMPLE 3
```powershell
$modify = Read-Host -AsSecureString -Prompt 'Password to modify'
$excel = New-OfficeExcelProtectionOptions -ReservedBy 'Financial Planning'
Get-ChildItem -Path .\Finance\Published -Filter *.xlsx |
    Protect-OfficeDocument -Restriction Structure, ReadOnlyRecommended -RestrictionPassword $modify -Options $excel
```

Locks sheet structure in every published budget and names the reserving team, using one password read at the start.

## PARAMETERS

### -LegacyPasswordHash
Precomputed native workbook protection hash, with at most four hexadecimal digits.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ModifyPasswordHash
Precomputed native write-reservation hash, with at most four hexadecimal digits.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ReservedBy
Name displayed by Excel for the write reservation.

```yaml
Type: String
Parameter Sets: Default
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

- `None`

## OUTPUTS

- `PSWriteOffice.Services.OfficeProtectionExcelOptions`: Shell options for the Excel protection vocabulary in the cross-format command.

## RELATED LINKS

- None
