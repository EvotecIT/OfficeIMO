---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficePowerPointTextBox
## SYNOPSIS
Adds a native text box with deliberate placement and plain or rich content.

## SYNTAX
### Text (Default)
```powershell
Add-OfficePowerPointTextBox [-Text] <string> [-Slide <PowerPointSlideSelector>] [-Name <string>] [-X <OfficeLength>] [-Y <OfficeLength>] [-Width <OfficeLength>] [-Height <OfficeLength>] [-Column <Int32>] [-Span <Int32>] [-PassThru] [<CommonParameters>]
```

### Content
```powershell
Add-OfficePowerPointTextBox [-Content] <scriptblock> [-Slide <PowerPointSlideSelector>] [-Name <string>] [-X <OfficeLength>] [-Y <OfficeLength>] [-Width <OfficeLength>] [-Height <OfficeLength>] [-Column <Int32>] [-Span <Int32>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a native text box with deliberate placement and plain or rich content.

## EXAMPLES

### EXAMPLE 1
```powershell
PowerPointSlide -Title 'Service review' {
    PowerPointTextBox -X 2cm -Y 4cm -Width 12cm -Height 3cm {
        PowerPointParagraph { PowerPointText 'Status: ' -Bold; PowerPointText 'Healthy' }
    }
}
```


## PARAMETERS

### -Column
Gets or sets the one-based theme grid column; explicit X overrides its position.

```yaml
Type: Int32
Parameter Sets: Text, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Content
Gets or sets native paragraph and run content.

```yaml
Type: ScriptBlock
Parameter Sets: Content
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Height
Gets or sets the height; omission uses the native default.

```yaml
Type: OfficeLength
Parameter Sets: Text, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Name
Gets or sets the native shape name.

```yaml
Type: String
Parameter Sets: Text, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Gets or sets whether the live native text box is returned.

```yaml
Type: SwitchParameter
Parameter Sets: Text, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Slide
Gets or sets a native slide or one-based position; omission uses the current slide block.

```yaml
Type: PowerPointSlideSelector
Parameter Sets: Text, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Span
Gets or sets the number of grid columns; the default is twelve, and explicit Width overrides it.

```yaml
Type: Int32
Parameter Sets: Text, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Gets or sets plain text; an empty string creates an empty box.

```yaml
Type: String
Parameter Sets: Text
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Gets or sets the width; omission uses the native default.

```yaml
Type: OfficeLength
Parameter Sets: Text, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -X
Gets or sets the left coordinate; numbers mean points.

```yaml
Type: OfficeLength
Parameter Sets: Text, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Y
Gets or sets the top coordinate; numbers mean points.

```yaml
Type: OfficeLength
Parameter Sets: Text, Content
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

- `PSWriteOffice.Foundation.PowerPointSlideSelector`: A live slide or a one-based slide position in a document.

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointTextBox`

## RELATED LINKS

- None
