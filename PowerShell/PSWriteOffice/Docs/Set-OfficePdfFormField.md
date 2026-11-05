---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficePdfFormField
## SYNOPSIS
Fills existing AcroForm fields from a hashtable, optionally flattening the result.

## SYNTAX
### Document (Default)
```powershell
Set-OfficePdfFormField -Value <hashtable> [-Flatten] [-KeepNeedAppearances] [-Incremental] [-AppearanceFontPath <string>] [-AppearanceFont <string>] [-PassThru] [-Document <PdfDocument>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
Set-OfficePdfFormField [-Path] <string[]> -OutputPath <string> -Value <hashtable> [-Flatten] [-KeepNeedAppearances] [-Incremental] [-AppearanceFontPath <string>] [-AppearanceFont <string>] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### PathDirectory
```powershell
Set-OfficePdfFormField [-Path] <string[]> -Value <hashtable> -OutputDirectory <string> [-Flatten] [-KeepNeedAppearances] [-Incremental] [-AppearanceFontPath <string>] [-AppearanceFont <string>] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPath
```powershell
Set-OfficePdfFormField -OutputPath <string> -Value <hashtable> -LiteralPath <string[]> [-Flatten] [-KeepNeedAppearances] [-Incremental] [-AppearanceFontPath <string>] [-AppearanceFont <string>] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPathDirectory
```powershell
Set-OfficePdfFormField -Value <hashtable> -OutputDirectory <string> -LiteralPath <string[]> [-Flatten] [-KeepNeedAppearances] [-Incremental] [-AppearanceFontPath <string>] [-AppearanceFont <string>] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Use Open-OfficePdf editing blocks for existing fields. Set initial field values with Add-OfficePdfFormField when composing a new PDF.

## EXAMPLES

### EXAMPLE 1
```powershell
Set-OfficePdfFormField -Path .\Forms\Access-Request.pdf -OutputPath .\Forms\Access-Request-Ada.pdf -Value @{ Requester = 'Ada Lovelace'; System = 'ERP' }
```


## PARAMETERS

### -AppearanceFont
AppearanceFont for this operation.

```yaml
Type: String
Parameter Sets: Document, Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AppearanceFontPath
AppearanceFontPath for this operation.

```yaml
Type: String
Parameter Sets: Document, Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Document for this operation.

```yaml
Type: PdfDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Flatten
Flatten for this operation.

```yaml
Type: SwitchParameter
Parameter Sets: Document, Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IgnorePermissionRestrictions
After successful authentication only; never bypasses a missing password.

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

### -Incremental
Incremental for this operation.

```yaml
Type: SwitchParameter
Parameter Sets: Document, Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -KeepNeedAppearances
KeepNeedAppearances for this operation.

```yaml
Type: SwitchParameter
Parameter Sets: Document, Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LiteralPath
LiteralPath for this operation.

```yaml
Type: String[]
Parameter Sets: LiteralPath, LiteralPathDirectory
Aliases: PSPath, FullName
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -LoadOptions
LoadOptions for this operation.

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

### -OutputDirectory
OutputDirectory for this operation.

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
OutputPath for this operation.

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
PassThru for this operation.

```yaml
Type: SwitchParameter
Parameter Sets: Document, Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Argument transformation also accepts a plain string (with a warning) or a PSCredential (its password).

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
Path for this operation.

```yaml
Type: String[]
Parameter Sets: Path, PathDirectory
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Value
Field name -> value; replaces -Field.

```yaml
Type: Hashtable
Parameter Sets: Document, Path, PathDirectory, LiteralPath, LiteralPathDirectory
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String[]`
- `OfficeIMO.Pdf.PdfDocument`

## OUTPUTS

- `System.IO.FileInfo`
- `OfficeIMO.Pdf.PdfDocument`

## RELATED LINKS

- None
