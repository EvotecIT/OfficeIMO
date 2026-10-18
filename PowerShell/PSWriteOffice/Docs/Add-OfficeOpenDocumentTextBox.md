---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeOpenDocumentTextBox
## SYNOPSIS
Adds a positioned text box to an OpenDocument presentation slide.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeOpenDocumentTextBox [-Text] <string> [-Slide <OdpSlide>] [-X <double>] [-Y <double>] [-Width <double>] [-Height <double>] [-Name <string>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a positioned text box to an OpenDocument presentation slide.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Add-OfficeOpenDocumentTextBox -Text 'Approved' -X 18 -Y 12 -Width 6 -Height 2
```


## PARAMETERS

### -Height
Height in centimeters.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Name
Optional shape name.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the created text box.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Slide
Slide target. Omit inside Add-OfficeOpenDocumentSlide -Content.

```yaml
Type: OdpSlide
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Text
Text box content.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Width in centimeters.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -X
Horizontal position in centimeters.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Y
Vertical position in centimeters.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
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

- `OfficeIMO.OpenDocument.OdpSlide`

## OUTPUTS

- `OfficeIMO.OpenDocument.OdpTextBox`

## RELATED LINKS

- None
