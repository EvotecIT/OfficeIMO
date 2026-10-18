---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficePdfTextDiagnostic
## SYNOPSIS
Gets PDF text encoding and advanced-layout diagnostics for generated text before rendering.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficePdfTextDiagnostic [-Text] <string> [-Source <string>] [-FontPath <string>] [-Encoding] [-AdvancedLayout] [<CommonParameters>]
```

## DESCRIPTION
Gets PDF text encoding and advanced-layout diagnostics for generated text before rendering.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficePdfTextDiagnostic -Text 'مرحبا' -AdvancedLayout
```

Returns OfficeIMO.Pdf diagnostics describing right-to-left and complex-script layout requirements.

## PARAMETERS

### -AdvancedLayout
Emit only advanced-layout diagnostics such as right-to-left, complex-script shaping, mark positioning, and script line breaking.

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

### -Encoding
Emit only encoding/glyph coverage diagnostics.

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

### -FontPath
Optional TrueType or OpenType/CFF font used for embedded-font glyph coverage and layout diagnostics.

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

### -Source
Optional source label included in diagnostic objects.

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

### -Text
Text to inspect.

```yaml
Type: String
Parameter Sets: __AllParameterSets
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

- `System.String`

## OUTPUTS

- `OfficeIMO.Pdf.PdfTextEncodingDiagnostic`
- `OfficeIMO.Pdf.PdfTextShapingDiagnostic`

## RELATED LINKS

- None
