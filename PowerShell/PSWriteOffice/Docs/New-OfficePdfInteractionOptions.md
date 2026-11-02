---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficePdfInteractionOptions
## SYNOPSIS
Creates native region limits and invisible-text policy for Get-OfficePdfInteractionMap.

## SYNTAX
### Default (Default)
```powershell
New-OfficePdfInteractionOptions [-MaxTextRegions <Int32>] [-MaxImageRegions <Int32>] [-IncludeInvisibleText] [<CommonParameters>]
```

## DESCRIPTION
Creates native region limits and invisible-text policy for Get-OfficePdfInteractionMap.

## EXAMPLES

### EXAMPLE 1
```powershell
$options = New-OfficePdfInteractionOptions -IncludeInvisibleText -MaxTextRegions 50000
Get-OfficePdfInteractionMap -Path ./Scanned.pdf -Options $options
```


## PARAMETERS

### -IncludeInvisibleText
Include invisible extracted text such as an OCR layer.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxImageRegions
Maximum image-placement regions emitted for a page.

```yaml
Type: Int32
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxTextRegions
Maximum text-element regions emitted for a page.

```yaml
Type: Int32
Parameter Sets: Default
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

- `OfficeIMO.Pdf.PdfPageInteractionOptions`

## RELATED LINKS

- None
