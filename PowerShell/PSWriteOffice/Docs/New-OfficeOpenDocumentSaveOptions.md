---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeOpenDocumentSaveOptions
## SYNOPSIS
Creates native OdfSaveOptions settings for document operations.

## SYNTAX
### Default (Default)
```powershell
New-OfficeOpenDocumentSaveOptions [-CompatibilityProfile <OdfCompatibilityProfile>] [-SignatureHandling <OdfSignatureHandling>] [-EncryptionHandling <OdfEncryptionHandling>] [-NoDeterministic] [<CommonParameters>]
```

## DESCRIPTION
Creates native OdfSaveOptions settings for document operations.

## EXAMPLES

### EXAMPLE 1
```powershell
$options = New-OfficeOpenDocumentSaveOptions
```


## PARAMETERS

### -CompatibilityProfile
CompatibilityProfile override; omission preserves the native default.

```yaml
Type: OdfCompatibilityProfile
Parameter Sets: Default
Aliases: None
Possible values: Odf14, Odf13, PreserveSource

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -EncryptionHandling
EncryptionHandling override; omission preserves the native default.

```yaml
Type: OdfEncryptionHandling
Parameter Sets: Default
Aliases: None
Possible values: Preserve, Remove

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoDeterministic
NoDeterministic override; omission preserves the native default.

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

### -SignatureHandling
SignatureHandling override; omission preserves the native default.

```yaml
Type: OdfSignatureHandling
Parameter Sets: Default
Aliases: None
Possible values: RejectInvalidation, RemoveInvalidated

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

- `OfficeIMO.OpenDocument.OdfSaveOptions`

## RELATED LINKS

- None
