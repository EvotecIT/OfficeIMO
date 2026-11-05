---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficePdfDiagnostic
## SYNOPSIS
Returns native PDF structure, preflight, incremental, optimization or generated-text diagnostics.

## SYNTAX
### Path (Default)
```powershell
Get-OfficePdfDiagnostic [-Path] <string[]> [-Type <OfficePdfDiagnosticType>] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [<CommonParameters>]
```

### LiteralPath
```powershell
Get-OfficePdfDiagnostic -LiteralPath <string[]> [-Type <OfficePdfDiagnosticType>] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [<CommonParameters>]
```

### Text
```powershell
Get-OfficePdfDiagnostic -Text <string> [-Source <string>] [-FontPath <string>] [-Encoding] [-AdvancedLayout] [<CommonParameters>]
```

## DESCRIPTION
Returns native PDF structure, preflight, incremental, optimization or generated-text diagnostics.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-OfficePdfDiagnostic -Path @('C:\Path')
```


### EXAMPLE 2
```powershell
Get-OfficePdfDiagnostic -LiteralPath @('C:\Path')
```


### EXAMPLE 3
```powershell
Get-OfficePdfDiagnostic -Text 'Value'
```


## PARAMETERS

### -AdvancedLayout
Return script, direction, mark positioning and line-breaking findings.

```yaml
Type: SwitchParameter
Parameter Sets: Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Encoding
Return encoding and glyph coverage findings.

```yaml
Type: SwitchParameter
Parameter Sets: Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FontPath
Literal font file for glyph coverage and shaping diagnostics.

```yaml
Type: String
Parameter Sets: Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IgnorePermissionRestrictions
After authentication, explicitly ignore owner-imposed restrictions.

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
Literal input files.

```yaml
Type: String[]
Parameter Sets: LiteralPath
Aliases: FullName, PSPath
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -LoadOptions
Native source loading policy.

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

### -Password
Source password.

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
Input PDF files with wildcard expansion.

```yaml
Type: String[]
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Source
Source label included in text findings.

```yaml
Type: String
Parameter Sets: Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Generated text to diagnose before rendering.

```yaml
Type: String
Parameter Sets: Text
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Type
Native report to return for each PDF.

```yaml
Type: OfficePdfDiagnosticType
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values: Structure, Preflight, Incremental, Optimization

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
- `System.String`

## OUTPUTS

- `None`

## RELATED LINKS

- None
