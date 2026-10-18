---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficePdfTableCellImage
## SYNOPSIS
Creates a typed image for a PDF table cell.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficePdfTableCellImage [-Path] <string> -Width <double> -Height <double> [-LinkUri <string>] [-LinkContents <string>] [<CommonParameters>]
```

## DESCRIPTION
Creates a typed image for a PDF table cell.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $logo = New-OfficePdfTableCellImage -Path .\logo.png -Width 28 -Height 28 -LinkUri 'https://example.com'
$cell = New-OfficePdfTableCell -Text 'Portal' -Image $logo
```

The image remains a native PDF table-cell visual and may carry its own link.

## PARAMETERS

### -Height
Rendered height in PDF points.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LinkContents
Accessible annotation text for the image link.

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

### -LinkUri
Optional absolute or catalog-base-relative URI linked from the image.

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

### -Path
Raster image path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: ImagePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Rendered width in PDF points.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
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

- `None`

## OUTPUTS

- `OfficeIMO.Pdf.PdfTableCellImage`

## RELATED LINKS

- None
