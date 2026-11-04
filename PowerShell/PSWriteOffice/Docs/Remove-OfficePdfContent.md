---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Remove-OfficePdfContent
## SYNOPSIS
Permanently removes (redacts) text, pattern matches or areas from a PDF and writes a new PDF.

## SYNTAX
### Text (Default)
```powershell
Remove-OfficePdfContent [-Path] <string[]> -OutputPath <string> -Text <string[]> [-CaseSensitive] [-FillColor <OfficeColor>] [-OnlyPaintMatches] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-NoClobber] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### TextDirectory
```powershell
Remove-OfficePdfContent [-Path] <string[]> -Text <string[]> -OutputDirectory <string> [-CaseSensitive] [-FillColor <OfficeColor>] [-OnlyPaintMatches] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-NoClobber] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Pattern
```powershell
Remove-OfficePdfContent [-Path] <string[]> -OutputPath <string> -Pattern <string[]> [-CaseSensitive] [-FillColor <OfficeColor>] [-OnlyPaintMatches] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-NoClobber] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### PatternDirectory
```powershell
Remove-OfficePdfContent [-Path] <string[]> -OutputDirectory <string> -Pattern <string[]> [-CaseSensitive] [-FillColor <OfficeColor>] [-OnlyPaintMatches] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-NoClobber] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Area
```powershell
Remove-OfficePdfContent [-Path] <string[]> -OutputPath <string> -Area <PdfRedactionArea[]> [-FillColor <OfficeColor>] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-NoClobber] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### AreaDirectory
```powershell
Remove-OfficePdfContent [-Path] <string[]> -OutputDirectory <string> -Area <PdfRedactionArea[]> [-FillColor <OfficeColor>] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-NoClobber] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Rectangle
```powershell
Remove-OfficePdfContent [-Path] <string[]> -OutputPath <string> -Page <Int32> -X <OfficeLength> -Y <OfficeLength> -Width <OfficeLength> -Height <OfficeLength> [-FillColor <OfficeColor>] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-NoClobber] [-PassThru] [-Label <string>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### RectangleDirectory
```powershell
Remove-OfficePdfContent [-Path] <string[]> -OutputDirectory <string> -Page <Int32> -X <OfficeLength> -Y <OfficeLength> -Width <OfficeLength> -Height <OfficeLength> [-FillColor <OfficeColor>] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-NoClobber] [-PassThru] [-Label <string>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### TextLiteralPath
```powershell
Remove-OfficePdfContent -OutputPath <string> -Text <string[]> -LiteralPath <string[]> [-CaseSensitive] [-FillColor <OfficeColor>] [-OnlyPaintMatches] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-NoClobber] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### PatternLiteralPath
```powershell
Remove-OfficePdfContent -OutputPath <string> -LiteralPath <string[]> -Pattern <string[]> [-CaseSensitive] [-FillColor <OfficeColor>] [-OnlyPaintMatches] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-NoClobber] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### AreaLiteralPath
```powershell
Remove-OfficePdfContent -OutputPath <string> -LiteralPath <string[]> -Area <PdfRedactionArea[]> [-FillColor <OfficeColor>] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-NoClobber] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### RectangleLiteralPath
```powershell
Remove-OfficePdfContent -OutputPath <string> -LiteralPath <string[]> -Page <Int32> -X <OfficeLength> -Y <OfficeLength> -Width <OfficeLength> -Height <OfficeLength> [-FillColor <OfficeColor>] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-NoClobber] [-PassThru] [-Label <string>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### TextLiteralPathDirectory
```powershell
Remove-OfficePdfContent -Text <string[]> -OutputDirectory <string> -LiteralPath <string[]> [-CaseSensitive] [-FillColor <OfficeColor>] [-OnlyPaintMatches] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-NoClobber] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### PatternLiteralPathDirectory
```powershell
Remove-OfficePdfContent -OutputDirectory <string> -LiteralPath <string[]> -Pattern <string[]> [-CaseSensitive] [-FillColor <OfficeColor>] [-OnlyPaintMatches] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-NoClobber] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### AreaLiteralPathDirectory
```powershell
Remove-OfficePdfContent -OutputDirectory <string> -LiteralPath <string[]> -Area <PdfRedactionArea[]> [-FillColor <OfficeColor>] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-NoClobber] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### RectangleLiteralPathDirectory
```powershell
Remove-OfficePdfContent -OutputDirectory <string> -LiteralPath <string[]> -Page <Int32> -X <OfficeLength> -Y <OfficeLength> -Width <OfficeLength> -Height <OfficeLength> [-FillColor <OfficeColor>] [-Password <securestring>] [-IgnorePermissionRestrictions] [-LoadOptions <PdfLoadOptions>] [-NoClobber] [-PassThru] [-Label <string>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Permanently removes (redacts) text, pattern matches or areas from a PDF and writes a new PDF.

## EXAMPLES

### EXAMPLE 1
```powershell
Remove-OfficePdfContent -Path .\Contracts\MSA-Contoso.pdf -OutputPath .\Contracts\MSA-Redacted.pdf -Text 'Contoso Ltd'
```


## PARAMETERS

### -Area
Area for this operation.

```yaml
Type: PdfRedactionArea[]
Parameter Sets: Area, AreaDirectory, AreaLiteralPath, AreaLiteralPathDirectory
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CaseSensitive
Replaces -MatchCase.

```yaml
Type: SwitchParameter
Parameter Sets: Text, TextDirectory, Pattern, PatternDirectory, TextLiteralPath, PatternLiteralPath, TextLiteralPathDirectory, PatternLiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FillColor
FillColor for this operation.

```yaml
Type: OfficeColor
Parameter Sets: Text, TextDirectory, Pattern, PatternDirectory, Area, AreaDirectory, Rectangle, RectangleDirectory, TextLiteralPath, PatternLiteralPath, AreaLiteralPath, RectangleLiteralPath, TextLiteralPathDirectory, PatternLiteralPathDirectory, AreaLiteralPathDirectory, RectangleLiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Height
Height for this operation.

```yaml
Type: OfficeLength
Parameter Sets: Rectangle, RectangleDirectory, RectangleLiteralPath, RectangleLiteralPathDirectory
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IgnorePermissionRestrictions
After successful authentication only; never bypasses a missing password.

```yaml
Type: SwitchParameter
Parameter Sets: Text, TextDirectory, Pattern, PatternDirectory, Area, AreaDirectory, Rectangle, RectangleDirectory, TextLiteralPath, PatternLiteralPath, AreaLiteralPath, RectangleLiteralPath, TextLiteralPathDirectory, PatternLiteralPathDirectory, AreaLiteralPathDirectory, RectangleLiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Label
Label for this operation.

```yaml
Type: String
Parameter Sets: Rectangle, RectangleDirectory, RectangleLiteralPath, RectangleLiteralPathDirectory
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
Parameter Sets: TextLiteralPath, PatternLiteralPath, AreaLiteralPath, RectangleLiteralPath, TextLiteralPathDirectory, PatternLiteralPathDirectory, AreaLiteralPathDirectory, RectangleLiteralPathDirectory
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
Parameter Sets: Text, TextDirectory, Pattern, PatternDirectory, Area, AreaDirectory, Rectangle, RectangleDirectory, TextLiteralPath, PatternLiteralPath, AreaLiteralPath, RectangleLiteralPath, TextLiteralPathDirectory, PatternLiteralPathDirectory, AreaLiteralPathDirectory, RectangleLiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoClobber
NoClobber for this operation.

```yaml
Type: SwitchParameter
Parameter Sets: Text, TextDirectory, Pattern, PatternDirectory, Area, AreaDirectory, Rectangle, RectangleDirectory, TextLiteralPath, PatternLiteralPath, AreaLiteralPath, RectangleLiteralPath, TextLiteralPathDirectory, PatternLiteralPathDirectory, AreaLiteralPathDirectory, RectangleLiteralPathDirectory
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OnlyPaintMatches
OnlyPaintMatches for this operation.

```yaml
Type: SwitchParameter
Parameter Sets: Text, TextDirectory, Pattern, PatternDirectory, TextLiteralPath, PatternLiteralPath, TextLiteralPathDirectory, PatternLiteralPathDirectory
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
Parameter Sets: TextDirectory, PatternDirectory, AreaDirectory, RectangleDirectory, TextLiteralPathDirectory, PatternLiteralPathDirectory, AreaLiteralPathDirectory, RectangleLiteralPathDirectory
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
Parameter Sets: Text, Pattern, Area, Rectangle, TextLiteralPath, PatternLiteralPath, AreaLiteralPath, RectangleLiteralPath
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Page
1-based.

```yaml
Type: Int32
Parameter Sets: Rectangle, RectangleDirectory, RectangleLiteralPath, RectangleLiteralPathDirectory
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
Parameter Sets: Text, TextDirectory, Pattern, PatternDirectory, Area, AreaDirectory, Rectangle, RectangleDirectory, TextLiteralPath, PatternLiteralPath, AreaLiteralPath, RectangleLiteralPath, TextLiteralPathDirectory, PatternLiteralPathDirectory, AreaLiteralPathDirectory, RectangleLiteralPathDirectory
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
Parameter Sets: Text, TextDirectory, Pattern, PatternDirectory, Area, AreaDirectory, Rectangle, RectangleDirectory, TextLiteralPath, PatternLiteralPath, AreaLiteralPath, RectangleLiteralPath, TextLiteralPathDirectory, PatternLiteralPathDirectory, AreaLiteralPathDirectory, RectangleLiteralPathDirectory
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
Parameter Sets: Text, TextDirectory, Pattern, PatternDirectory, Area, AreaDirectory, Rectangle, RectangleDirectory
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Pattern
Bounded regular expressions matched with the same case policy as Get-OfficePdfRedactionPlan.

```yaml
Type: String[]
Parameter Sets: Pattern, PatternDirectory, PatternLiteralPath, PatternLiteralPathDirectory
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Text for this operation.

```yaml
Type: String[]
Parameter Sets: Text, TextDirectory, TextLiteralPath, TextLiteralPathDirectory
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Width for this operation.

```yaml
Type: OfficeLength
Parameter Sets: Rectangle, RectangleDirectory, RectangleLiteralPath, RectangleLiteralPathDirectory
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -X
X for this operation.

```yaml
Type: OfficeLength
Parameter Sets: Rectangle, RectangleDirectory, RectangleLiteralPath, RectangleLiteralPathDirectory
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Y
Y for this operation.

```yaml
Type: OfficeLength
Parameter Sets: Rectangle, RectangleDirectory, RectangleLiteralPath, RectangleLiteralPathDirectory
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

- `System.Object`

## RELATED LINKS

- None
