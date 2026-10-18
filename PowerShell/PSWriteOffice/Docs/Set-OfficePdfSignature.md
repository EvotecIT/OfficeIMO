---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficePdfSignature
## SYNOPSIS
Injects externally produced CMS, CAdES, or timestamp signature bytes into a prepared PDF signature placeholder.

## SYNTAX
### __AllParameterSets
```powershell
Set-OfficePdfSignature [-Path] <string> [-SignaturePath] <string> [-OutputPath] <string> [-Password <string>] [-IgnorePermissionRestrictions] [-PassThruReport] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Injects externally produced CMS, CAdES, or timestamp signature bytes into a prepared PDF signature placeholder.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Set-OfficePdfSignature -Path .\Prepared.pdf -SignaturePath .\signature.der -OutputPath .\Signed.pdf
            Get-OfficePdfSignature -Path .\Signed.pdf
```

Writes a PDF with the reserved /Contents hex slot patched in place.

## PARAMETERS

### -IgnorePermissionRestrictions
After successful password authentication, explicitly ignore owner-imposed signature restrictions.

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

### -OutputPath
Output signed PDF path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the object created or changed by the command.

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

### -PassThruReport
Return a signature validation report for the written PDF instead of only the output file.

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
Password used to authenticate an encrypted prepared PDF.

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
Prepared PDF path.

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

### -SignaturePath
DER/CMS/CAdES/TSA response bytes to inject into the reserved /Contents slot.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String`

## OUTPUTS

- `System.IO.FileInfo`
- `OfficeIMO.Pdf.PdfSignatureValidationReport`

## RELATED LINKS

- None
