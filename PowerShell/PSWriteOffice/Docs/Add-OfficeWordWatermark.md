---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeWordWatermark
## SYNOPSIS
Adds a watermark to the current section or header.

## SYNTAX
### Text (Default)
```powershell
Add-OfficeWordWatermark [-Text] <string> [-X <OfficeLength>] [-Y <OfficeLength>] [-Scale <double>] [-PassThru] [<CommonParameters>]
```

### Image
```powershell
Add-OfficeWordWatermark [-ImagePath] <string> [-X <OfficeLength>] [-Y <OfficeLength>] [-Scale <double>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Supports text or image watermarks using OfficeIMO.Word.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficeWord -Path .\ProtectedReport.docx {
    Add-OfficeWordParagraph -Text 'Confidential report'
    Add-OfficeWordWatermark -Text 'CONFIDENTIAL' -Scale 1.2
    Protect-OfficeDocument -Restriction ReadOnly -RestrictionPassword 'secret'
}
```

Applies a text watermark to the current section and then protects the document through OfficeIMO settings.

### EXAMPLE 2
```powershell
PS> New-OfficeWord -Path .\DraftReport.docx {
    Add-OfficeWordParagraph -Text 'Draft report'
    Add-OfficeWordWatermark -ImagePath .\Assets\Draft.png -Scale 0.6 -X 20 -Y 40
}
```

Uses the image watermark path and placement parameters exposed by OfficeIMO.Word.

## PARAMETERS

### -ImagePath
Path to an image watermark, resolved literally against the current PowerShell location.

```yaml
Type: String
Parameter Sets: Image
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the created watermark.

```yaml
Type: SwitchParameter
Parameter Sets: Text, Image
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Scale
Scale factor for the watermark.

```yaml
Type: Double
Parameter Sets: Text, Image
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Watermark text.

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

### -X
Horizontal offset for the watermark.

```yaml
Type: OfficeLength
Parameter Sets: Text, Image
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Y
Vertical offset for the watermark.

```yaml
Type: OfficeLength
Parameter Sets: Text, Image
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

- `None`

## OUTPUTS

- `None`

## RELATED LINKS

- None
