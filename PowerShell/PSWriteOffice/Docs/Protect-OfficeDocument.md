---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Protect-OfficeDocument
## SYNOPSIS
Adds native password encryption or editing restrictions to a document.

## SYNTAX
### Path (Default)
```powershell
Protect-OfficeDocument [-Path] <string[]> [-Password <securestring>] [-Encrypt] [-Restriction <OfficeProtectionRestriction[]>] [-RestrictionPassword <securestring>] [-OutputPath <string>] [-Options <Object[]>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPath
```powershell
Protect-OfficeDocument -LiteralPath <string[]> [-Password <securestring>] [-Encrypt] [-Restriction <OfficeProtectionRestriction[]>] [-RestrictionPassword <securestring>] [-OutputPath <string>] [-Options <Object[]>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Protect-OfficeDocument [-Document] <IDisposable> [-Password <securestring>] [-Encrypt] [-Restriction <OfficeProtectionRestriction[]>] [-RestrictionPassword <securestring>] [-OutputPath <string>] [-Options <Object[]>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Supports PDF and Open XML Word, Excel and PowerPoint packages. Path input is saved in place unless
OutputPath names a protected copy. File inputs must be unencrypted: this command has no CurrentPassword
parameter. Live Word/Excel restrictions may remain in memory; encryption requires OutputPath or a native
associated file path. Encryption is a one-time export, not a stored password setting on the live input.
PDF protection returns a new native document with preservation proof, leaving the original live PDF intact.
Without Encrypt, PDF restrictions use Password or RestrictionPassword as the owner password and allow opening
without a password. With Encrypt, a separate owner password keeps the open password from authorizing permission
changes. Unsupported restrictions produce warnings; an entirely unsupported request writes nothing.
Safe native signature and preservation policies remain in force.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-ChildItem ./Payroll -Filter *.xlsx | Protect-OfficeDocument -Encrypt -Password (Read-Host -AsSecureString) -PassThru
```


### EXAMPLE 2
```powershell
Protect-OfficeDocument -Path ./policy.docx -Restriction ReadOnly -Password (Read-Host -AsSecureString)
```


### EXAMPLE 3
```powershell
Protect-OfficeDocument -Path ./board.pdf -Encrypt -Password $openPassword -Restriction Copy -RestrictionPassword $ownerPassword -OutputPath ./protected.pdf
```


### EXAMPLE 4
```powershell
Protect-OfficeDocument -Path ./board.pdf -Restriction Copy -Password $ownerPassword -OutputPath ./restricted.pdf
```


## PARAMETERS

### -Document
Caller-owned native Word, Excel, PowerPoint or PDF document.

```yaml
Type: IDisposable
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Encrypt
Encrypt the saved artifact using its native engine.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LiteralPath
Literal files, including FileInfo.FullName and provider PSPath properties.

```yaml
Type: String[]
Parameter Sets: LiteralPath
Aliases: PSPath, FullName
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Options
Excel tables accept LegacyPasswordHash, ModifyPasswordHash and ReservedBy; native workbook and reservation
property tables are also accepted. PDF tables accept Algorithm, EncryptMetadata, Permissions or AllowedPermissions,
and AesCryptographyProvider. Supply PDF passwords through Password and RestrictionPassword. Typed native PDF
options retain their owner password, permission mask, algorithm, metadata setting and AES provider.

```yaml
Type: Object[]
Parameter Sets: Path, LiteralPath, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputPath
Protected file destination; accepts exactly one input when supplied.

```yaml
Type: String
Parameter Sets: Path, LiteralPath, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit FileInfo for file input, or the protected native document for Document input.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Open password with Encrypt; otherwise the restriction password. Accepts a credential or warned plain text.

```yaml
Type: SecureString
Parameter Sets: Path, LiteralPath, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Files to protect, in pipeline and provider match order.

```yaml
Type: String[]
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Restriction
Word editing mode, Excel protection/reservation, or PDF permissions to withhold.

```yaml
Type: OfficeProtectionRestriction[]
Parameter Sets: Path, LiteralPath, Document
Aliases: None
Possible values: ReadOnly, Comments, TrackedChanges, Forms, Structure, Windows, ReadOnlyRecommended, Print, Copy, Modify, Annotate, FillForms, Assemble

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RestrictionPassword
Restriction password, separate from the open password when Encrypt is used; PDF owner password.

```yaml
Type: SecureString
Parameter Sets: Path, LiteralPath, Document
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

- `System.String[]`
- `System.IDisposable`

## OUTPUTS

- `System.IO.FileInfo`
- `OfficeIMO.Word.WordDocument`
- `OfficeIMO.Excel.ExcelDocument`
- `OfficeIMO.PowerPoint.PowerPointPresentation`
- `OfficeIMO.Pdf.PdfDocument`

## RELATED LINKS

- None
