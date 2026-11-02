---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficePowerPointTheme
## SYNOPSIS
Restyles a deck or changes the selected masters' colors, fonts and theme name.

## SYNTAX
### Document (Default)
```powershell
Set-OfficePowerPointTheme [-Theme <OfficeThemeInput>] [-Colors <hashtable>] [-HeadingFont <string>] [-BodyFont <string>] [-HeadingFontEastAsian <string>] [-BodyFontEastAsian <string>] [-HeadingFontComplexScript <string>] [-BodyFontComplexScript <string>] [-ClearMissing] [-Name <string>] [-Master <int>] [-AllMasters] [-PassThru] [-Document <PowerPointDocument>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Restyles a deck or changes the selected masters' colors, fonts and theme name.

## EXAMPLES

### EXAMPLE 1
```powershell
Set-OfficePowerPointTheme -Document $deck -Theme Boardroom -AllMasters
```


### EXAMPLE 2
```powershell
Set-OfficePowerPointTheme -Document $deck -Colors @{ Accent1 = '#0F766E' } -HeadingFont 'Segoe UI Semibold' -BodyFont 'Segoe UI' -Name 'Service Brief'
```


## PARAMETERS

### -AllMasters
Gets or sets whether every master is updated.

```yaml
Type: SwitchParameter
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BodyFont
Gets or sets the Latin body font.

```yaml
Type: String
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BodyFontComplexScript
Gets or sets the complex script body font.

```yaml
Type: String
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BodyFontEastAsian
Gets or sets the East Asian body font.

```yaml
Type: String
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ClearMissing
Gets or sets whether unspecified font slots are cleared. Requires at least one font override.

```yaml
Type: SwitchParameter
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Colors
Gets or sets theme slots mapped to OfficeColor values, CSS colors or semantic roles.

```yaml
Type: Hashtable
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Gets or sets an explicit live document; omission uses the enclosing content block.

```yaml
Type: PowerPointDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -HeadingFont
Gets or sets the Latin heading font.

```yaml
Type: String
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HeadingFontComplexScript
Gets or sets the complex script heading font.

```yaml
Type: String
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HeadingFontEastAsian
Gets or sets the East Asian heading font.

```yaml
Type: String
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Master
Gets or sets the one-based master position; defaults to the first.

```yaml
Type: Int32
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Name
Gets or sets the theme name.

```yaml
Type: String
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Gets or sets whether the updated native document is returned.

```yaml
Type: SwitchParameter
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Theme
Gets or sets a built-in theme, OfficeTheme object or document/template path. Applying it restyles masters and layouts before overrides.

```yaml
Type: OfficeThemeInput
Parameter Sets: Document
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

- `OfficeIMO.PowerPoint.PowerPointDocument`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointDocument`

## RELATED LINKS

- None
