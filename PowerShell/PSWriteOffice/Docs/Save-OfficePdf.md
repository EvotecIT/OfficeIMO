---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Save-OfficePdf
## SYNOPSIS
Saves a live PDF to its associated path or a new destination without closing it.

## SYNTAX
### __AllParameterSets
```powershell
Save-OfficePdf [[-Path] <string>] -Document <PdfDocument> [-Show] [-PassThru] [-AllowLoss] [-Password <securestring>] [-OwnerPassword <securestring>] [-Permission <PdfStandardPermissions>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Native serialization rejects content loss unless AllowLoss is supplied. Password and OwnerPassword accept shared secure input. Security options encrypt an unencrypted output; use Protect-OfficeDocument for security changes to an already protected file.

## EXAMPLES

### EXAMPLE 1
```powershell
$pdf = New-OfficePdf -Path ./Review.pdf
Add-OfficePdfParagraph -Document $pdf -Text 'Reviewed and approved'
$pdf | Save-OfficePdf -PassThru
Close-OfficePdf -Document $pdf
```


### EXAMPLE 2
```powershell
$owner = Read-Host -AsSecureString
Save-OfficePdf -Document $pdf -Path ./Restricted.pdf -OwnerPassword $owner -Permission Print,Accessibility
```


## PARAMETERS

### -AllowLoss
Explicitly accept fidelity loss from native PDF serialization.

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

### -Document
Live PDF to save.

```yaml
Type: PdfDocument
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -OwnerPassword
Owner password; without Password the PDF opens without a password and enforces permissions.

```yaml
Type: SecureString
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Return the saved FileInfo.

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
Opening password for an unencrypted output; accepts secure input or warned plaintext.

```yaml
Type: SecureString
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
Optional destination; omission uses the native document's associated path.

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

### -Permission
Operations allowed to readers, expressed as native PDF permission flags.

```yaml
Type: PdfStandardPermissions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: None, Print, ModifyContents, CopyContents, ModifyAnnotations, FillForms, Accessibility, AssembleDocument, HighQualityPrint, All

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Show
Open the saved file in its registered default application.

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

- `OfficeIMO.Pdf.PdfDocument`

## OUTPUTS

- `System.IO.FileInfo`

## RELATED LINKS

- None
