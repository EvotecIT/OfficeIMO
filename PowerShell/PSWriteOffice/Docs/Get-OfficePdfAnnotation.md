---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficePdfAnnotation
## SYNOPSIS
Gets page annotations and action evidence from a PDF.

## SYNTAX
### Path (Default)
```powershell
Get-OfficePdfAnnotation [-Path] <string[]> [-LoadOptions <PdfLoadOptions>] [-Subtype <string>] [-Page <Int32>] [-WithAction] [-Password <securestring>] [-IgnorePermissionRestrictions] [<CommonParameters>]
```

### LiteralPath
```powershell
Get-OfficePdfAnnotation -LiteralPath <string[]> [-LoadOptions <PdfLoadOptions>] [-Subtype <string>] [-Page <Int32>] [-WithAction] [-Password <securestring>] [-IgnorePermissionRestrictions] [<CommonParameters>]
```

### Document
```powershell
Get-OfficePdfAnnotation -Document <PdfDocument> [-Subtype <string>] [-Page <Int32>] [-WithAction] [<CommonParameters>]
```

## DESCRIPTION
Gets page annotations and action evidence from a PDF.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-OfficePdfAnnotation -Path @('C:\Path')
```


### EXAMPLE 2
```powershell
Get-OfficePdfAnnotation -Document 'Value'
```


### EXAMPLE 3
```powershell
Get-OfficePdfAnnotation -LiteralPath @('C:\Path')
```


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

### -Page
Optional one-based page number filter.

```yaml
Type: Int32
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

### -Subtype
Optional annotation subtype filter such as Text, Link, Widget, or FreeText.

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

### -WithAction
Return only annotations with actions, additional actions, or chained actions.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String[]`
- `OfficeIMO.Pdf.PdfDocument`

## OUTPUTS

- `OfficeIMO.Pdf.PdfAnnotation`

## RELATED LINKS

- None
