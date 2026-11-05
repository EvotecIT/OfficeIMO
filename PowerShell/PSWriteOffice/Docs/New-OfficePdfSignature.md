---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficePdfSignature
## SYNOPSIS
Prepare external PDF signature using native byte-range and signature-placeholder operations.

## SYNTAX
### Path (Default)
```powershell
New-OfficePdfSignature [-Path] <string[]> -OutputPath <string> [-FieldName <string>] [-Filter <string>] [-SubFilter <PdfExternalSignatureSubFilter>] [-Name <string>] [-Reason <string>] [-Location <string>] [-ContactInfo <string>] [-ReservedSize <int>] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### PathDirectory
```powershell
New-OfficePdfSignature [-Path] <string[]> -OutputDirectory <string> [-FieldName <string>] [-Filter <string>] [-SubFilter <PdfExternalSignatureSubFilter>] [-Name <string>] [-Reason <string>] [-Location <string>] [-ContactInfo <string>] [-ReservedSize <int>] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPath
```powershell
New-OfficePdfSignature -OutputPath <string> -LiteralPath <string[]> [-FieldName <string>] [-Filter <string>] [-SubFilter <PdfExternalSignatureSubFilter>] [-Name <string>] [-Reason <string>] [-Location <string>] [-ContactInfo <string>] [-ReservedSize <int>] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPathDirectory
```powershell
New-OfficePdfSignature -OutputDirectory <string> -LiteralPath <string[]> [-FieldName <string>] [-Filter <string>] [-SubFilter <PdfExternalSignatureSubFilter>] [-Name <string>] [-Reason <string>] [-Location <string>] [-ContactInfo <string>] [-ReservedSize <int>] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Confirmation precedes loading, native mutation and output creation. This command does not obtain certificates or private keys.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficePdfSignature -Path @('C:\Path') -OutputPath 'C:\Path'
```


### EXAMPLE 2
```powershell
New-OfficePdfSignature -OutputPath 'C:\Path' -LiteralPath @('C:\Path')
```


### EXAMPLE 3
```powershell
New-OfficePdfSignature -OutputDirectory 'Value' -LiteralPath @('C:\Path')
```


## PARAMETERS

### -ContactInfo
ContactInfo setting.

```yaml
Type: String
Parameter Sets: Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FieldName
FieldName setting.

```yaml
Type: String
Parameter Sets: Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Filter
Filter setting.

```yaml
Type: String
Parameter Sets: Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IgnorePermissionRestrictions
IgnorePermissionRestrictions setting.

```yaml
Type: SwitchParameter
Parameter Sets: Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LiteralPath
Literal input PDF paths; accepts FileInfo.FullName.

```yaml
Type: String[]
Parameter Sets: LiteralPath, LiteralPathDirectory
Aliases: FullName, PSPath
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -LoadOptions
Native parsing, authentication and resource policy.

```yaml
Type: PdfLoadOptions
Parameter Sets: Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Location
Location setting.

```yaml
Type: String
Parameter Sets: Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Name
Name setting.

```yaml
Type: String
Parameter Sets: Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputDirectory
Batch destination retaining source names.

```yaml
Type: String
Parameter Sets: PathDirectory, LiteralPathDirectory
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputPath
Single-file destination.

```yaml
Type: String
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Return native annotation edit reports.

```yaml
Type: SwitchParameter
Parameter Sets: Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Password setting.

```yaml
Type: SecureString
Parameter Sets: Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Input PDF paths with wildcard expansion.

```yaml
Type: String[]
Parameter Sets: Path, PathDirectory
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Reason
Reason setting.

```yaml
Type: String
Parameter Sets: Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ReservedSize
ReservedSize setting.

```yaml
Type: Int32
Parameter Sets: Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SubFilter
SubFilter setting.

```yaml
Type: PdfExternalSignatureSubFilter
Parameter Sets: Path, PathDirectory, LiteralPath, LiteralPathDirectory
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

- `System.String[]`

## OUTPUTS

- `System.IO.FileInfo`
- `OfficeIMO.Pdf.PdfExternalSignaturePreparation`

## RELATED LINKS

- None
