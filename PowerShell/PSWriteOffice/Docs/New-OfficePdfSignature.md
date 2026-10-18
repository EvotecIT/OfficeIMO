---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficePdfSignature
## SYNOPSIS
Prepares an existing PDF for external digital signing by appending a signature field, /ByteRange, and reserved /Contents placeholder.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficePdfSignature [-Path] <string> [-OutputPath] <string> [-Password <string>] [-IgnorePermissionRestrictions] [-FieldName <string>] [-Filter <string>] [-SubFilter <PdfExternalSignatureSubFilter>] [-Name <string>] [-Reason <string>] [-Location <string>] [-ContactInfo <string>] [-ReservedBytes <int>] [-PassThruReport] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
The command does not create CMS, CAdES, timestamp, certificate-chain, or revocation data. Use the returned byte range or digest with an external signing service, then inject the produced signature bytes with Set-OfficePdfSignature.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $plan = New-OfficePdfSignature -Path .\Input.pdf -OutputPath .\Prepared.pdf -FieldName Approval -Name 'Alice' -Reason Approval -PassThruReport
$plan.ByteRangeValues
$plan.ComputeSha256Digest()
```

Writes a prepared PDF and returns the OfficeIMO.Pdf external signing preparation report.

## PARAMETERS

### -ContactInfo
Signer contact information stored in the signature dictionary.

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

### -FieldName
Signature field name to append.

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

### -Filter
Signature handler filter name. The default is Adobe.PPKLite.

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

### -IgnorePermissionRestrictions
After successful password authentication, explicitly ignore owner-imposed signature-field restrictions.

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

### -Location
Signing location stored in the signature dictionary.

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

### -Name
Display signer name stored in the signature dictionary.

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

### -OutputPath
Output prepared PDF path.

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

### -PassThruReport
Return the OfficeIMO.Pdf preparation report instead of only the output file.

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
Input PDF path.

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

### -Reason
Signing reason stored in the signature dictionary.

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

### -ReservedBytes
Raw signature bytes to reserve in /Contents before hex encoding.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SubFilter
Signature subfilter that describes the external signature bytes to inject later.

```yaml
Type: PdfExternalSignatureSubFilter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: DetachedCms, CadesDetached, DocumentTimestamp

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

- `System.IO.FileInfo`
- `OfficeIMO.Pdf.PdfExternalSignaturePreparation`

## RELATED LINKS

- None
