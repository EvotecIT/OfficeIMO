---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficePdfEncryptionOptions
## SYNOPSIS
Creates PDF encryption settings; Protect-OfficeDocument supplies the actual passwords.

## SYNTAX
### Default (Default)
```powershell
New-OfficePdfEncryptionOptions [-Algorithm <PdfStandardEncryptionAlgorithm>] [-NoMetadataEncryption] [<CommonParameters>]
```

## DESCRIPTION
Creates PDF encryption settings; Protect-OfficeDocument supplies the actual passwords.

## EXAMPLES

### EXAMPLE 1
```powershell
$options = New-OfficePdfEncryptionOptions -Algorithm Aes256 -NoMetadataEncryption
Protect-OfficeDocument -Path ./Report.pdf -Encrypt -Password $password -Options $options
```


## PARAMETERS

### -Algorithm
Native PDF Standard encryption algorithm.

```yaml
Type: PdfStandardEncryptionAlgorithm
Parameter Sets: Default
Aliases: None
Possible values: Aes256, Aes128, LegacyRc4

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoMetadataEncryption
Leave document metadata unencrypted.

```yaml
Type: SwitchParameter
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

- `System.Collections.Hashtable`

## RELATED LINKS

- None
