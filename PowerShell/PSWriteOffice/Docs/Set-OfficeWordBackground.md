---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeWordBackground
## SYNOPSIS
Sets the background for a Word document.

## SYNTAX
### Color (Default)
```powershell
Set-OfficeWordBackground [-Color] <string> [-Document <WordDocument>] [-PassThru] [<CommonParameters>]
```

### Image
```powershell
Set-OfficeWordBackground [-ImagePath] <string> [-Document <WordDocument>] [-Width <Double>] [-Height <Double>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Sets the background for a Word document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficeWord -Path .\BrandedReport.docx {
    Set-OfficeWordBackground -Color '#f4f7fb'
    Add-OfficeWordParagraph -Text 'Executive summary'
}
```

Sets the document background to the provided hex color and continues normal document composition.

### EXAMPLE 2
```powershell
PS> $doc = Get-OfficeWord -Path .\Draft.docx
$doc | Set-OfficeWordBackground -ImagePath .\Assets\Background.png -Width 600 -Height 800 -PassThru |
    Save-OfficeWord -Path .\Draft-Branded.docx
```

Uses OfficeIMO.Word background image support and saves the updated document through the standard save cmdlet.

## PARAMETERS

### -Color
Background color. Named colors and hexadecimal values are accepted.

```yaml
Type: String
Parameter Sets: Color
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Document to update when provided explicitly.

```yaml
Type: WordDocument
Parameter Sets: Color, Image
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Height
Optional background image height in pixels.

```yaml
Type: Double
Parameter Sets: Image
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ImagePath
Path to the background image.

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
Emit the updated document.

```yaml
Type: SwitchParameter
Parameter Sets: Color, Image
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Optional background image width in pixels.

```yaml
Type: Double
Parameter Sets: Image
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

- `OfficeIMO.Word.WordDocument`

## OUTPUTS

- `OfficeIMO.Word.WordDocument`

## RELATED LINKS

- None
