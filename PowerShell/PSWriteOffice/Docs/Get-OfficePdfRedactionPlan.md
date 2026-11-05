---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficePdfRedactionPlan
## SYNOPSIS
Previews explicit, literal-text or bounded regular-expression redaction selections without changing PDF content.

## SYNTAX
### Text (Default)
```powershell
Get-OfficePdfRedactionPlan [-Path] <string[]> -Text <string[]> [-CaseSensitive] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [<CommonParameters>]
```

### Pattern
```powershell
Get-OfficePdfRedactionPlan [-Path] <string[]> -Pattern <string[]> [-CaseSensitive] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [<CommonParameters>]
```

### Area
```powershell
Get-OfficePdfRedactionPlan [-Path] <string[]> -Area <PdfRedactionArea[]> [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [<CommonParameters>]
```

### Rectangle
```powershell
Get-OfficePdfRedactionPlan [-Path] <string[]> -Page <int> -X <OfficeLength> -Y <OfficeLength> -Width <OfficeLength> -Height <OfficeLength> [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-Label <string>] [<CommonParameters>]
```

### TextLiteralPath
```powershell
Get-OfficePdfRedactionPlan -Text <string[]> -LiteralPath <string[]> [-CaseSensitive] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [<CommonParameters>]
```

### PatternLiteralPath
```powershell
Get-OfficePdfRedactionPlan -LiteralPath <string[]> -Pattern <string[]> [-CaseSensitive] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [<CommonParameters>]
```

### AreaLiteralPath
```powershell
Get-OfficePdfRedactionPlan -LiteralPath <string[]> -Area <PdfRedactionArea[]> [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [<CommonParameters>]
```

### RectangleLiteralPath
```powershell
Get-OfficePdfRedactionPlan -LiteralPath <string[]> -Page <int> -X <OfficeLength> -Y <OfficeLength> -Width <OfficeLength> -Height <OfficeLength> [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-Label <string>] [<CommonParameters>]
```

## DESCRIPTION
Previews explicit, literal-text or bounded regular-expression redaction selections without changing PDF content.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-OfficePdfRedactionPlan -LiteralPath './Report[1].pdf' -Pattern 'Account: \d+'
```


## PARAMETERS

### -Area
Area setting.

```yaml
Type: PdfRedactionArea[]
Parameter Sets: Area, AreaLiteralPath
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CaseSensitive
CaseSensitive setting.

```yaml
Type: SwitchParameter
Parameter Sets: Text, Pattern, TextLiteralPath, PatternLiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Height
Height setting.

```yaml
Type: OfficeLength
Parameter Sets: Rectangle, RectangleLiteralPath
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IgnorePermissionRestrictions
IgnorePermissionRestrictions setting.

```yaml
Type: SwitchParameter
Parameter Sets: Text, Pattern, Area, Rectangle, TextLiteralPath, PatternLiteralPath, AreaLiteralPath, RectangleLiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Label
Label setting.

```yaml
Type: String
Parameter Sets: Rectangle, RectangleLiteralPath
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
Parameter Sets: TextLiteralPath, PatternLiteralPath, AreaLiteralPath, RectangleLiteralPath
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
Parameter Sets: Text, Pattern, Area, Rectangle, TextLiteralPath, PatternLiteralPath, AreaLiteralPath, RectangleLiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Page
Page setting.

```yaml
Type: Int32
Parameter Sets: Rectangle, RectangleLiteralPath
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Password setting.

```yaml
Type: SecureString
Parameter Sets: Text, Pattern, Area, Rectangle, TextLiteralPath, PatternLiteralPath, AreaLiteralPath, RectangleLiteralPath
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
Parameter Sets: Text, Pattern, Area, Rectangle
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Pattern
Pattern setting.

```yaml
Type: String[]
Parameter Sets: Pattern, PatternLiteralPath
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Text setting.

```yaml
Type: String[]
Parameter Sets: Text, TextLiteralPath
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Width setting.

```yaml
Type: OfficeLength
Parameter Sets: Rectangle, RectangleLiteralPath
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -X
X setting.

```yaml
Type: OfficeLength
Parameter Sets: Rectangle, RectangleLiteralPath
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Y
Y setting.

```yaml
Type: OfficeLength
Parameter Sets: Rectangle, RectangleLiteralPath
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

## OUTPUTS

- `OfficeIMO.Pdf.PdfRedactionPlan`

## RELATED LINKS

- None
