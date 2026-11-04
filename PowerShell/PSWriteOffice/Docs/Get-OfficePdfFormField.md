---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficePdfFormField
## SYNOPSIS
Gets simple AcroForm fields from a PDF.

## SYNTAX
### Path (Default)
```powershell
Get-OfficePdfFormField [-Path] <string[]> [-LoadOptions <PdfLoadOptions>] [-Name <string>] [-Password <securestring>] [-IgnorePermissionRestrictions] [<CommonParameters>]
```

### LiteralPath
```powershell
Get-OfficePdfFormField -LiteralPath <string[]> [-LoadOptions <PdfLoadOptions>] [-Name <string>] [-Password <securestring>] [-IgnorePermissionRestrictions] [<CommonParameters>]
```

### Document
```powershell
Get-OfficePdfFormField -Document <PdfDocument> [-Name <string>] [<CommonParameters>]
```

## DESCRIPTION
Gets simple AcroForm fields from a PDF.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficePdfFormField -Path .\Examples\Documents\Request.pdf |
    Select-Object Name, FieldType, Value
Set-OfficePdfForm -Path .\Examples\Documents\Request.pdf -OutputPath .\Examples\Documents\Request-Filled.pdf -Field @{ Requester = 'Ada Lovelace' }
```

Reads form field names so the fill hashtable can use the right keys.

## PARAMETERS

### -Document
Caller-owned live PDF; inspection never disposes it.

```yaml
Type: PdfDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -IgnorePermissionRestrictions
After successful password authentication, explicitly ignore owner-imposed usage restrictions.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LiteralPath
Exact PDF paths, including literal brackets.

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

### -LoadOptions
Bounded input load settings.

```yaml
Type: PdfLoadOptions
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Name
Optional field name filter.

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

### -Password
Password used to inspect a Standard password-encrypted PDF.

```yaml
Type: SecureString
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
PDF input paths; supports wildcards.

```yaml
Type: String[]
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String[]`
- `OfficeIMO.Pdf.PdfDocument`

## OUTPUTS

- `OfficeIMO.Pdf.PdfFormField`

## RELATED LINKS

- None
