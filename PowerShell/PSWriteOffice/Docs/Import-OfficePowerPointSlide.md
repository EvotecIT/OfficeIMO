---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Import-OfficePowerPointSlide
## SYNOPSIS
Imports selected slides from a file or live document, preserving native relationships.

## SYNTAX
### Path (Default)
```powershell
Import-OfficePowerPointSlide [-Path] <string> [-Document <PowerPointDocument>] [-SourceSlide <int[]>] [-Position <Int32>] [-PassThru] [<CommonParameters>]
```

### LiteralPath
```powershell
Import-OfficePowerPointSlide -LiteralPath <string> [-Document <PowerPointDocument>] [-SourceSlide <int[]>] [-Position <Int32>] [-PassThru] [<CommonParameters>]
```

### SourceDocument
```powershell
Import-OfficePowerPointSlide -SourceDocument <PowerPointDocument> [-Document <PowerPointDocument>] [-SourceSlide <int[]>] [-Position <Int32>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
SourceSlide and Position are one-based. Omit SourceSlide to import every slide in order; omit Position to append. The target remains open until explicitly saved and closed.

## EXAMPLES

### EXAMPLE 1
```powershell
Import-OfficePowerPointSlide -Document $deck -SourceDocument $library -SourceSlide 3, 4 -Position 2 -PassThru
```


### EXAMPLE 2
```powershell
Get-Item -LiteralPath './Appendix[final].pptx' | Import-OfficePowerPointSlide -Document $deck
```


## PARAMETERS

### -Document
Gets or sets the live target; omission uses the current presentation block.

```yaml
Type: PowerPointDocument
Parameter Sets: Path, LiteralPath, SourceDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -LiteralPath
Gets or sets a source file without wildcard expansion, including FileInfo.FullName.

```yaml
Type: String
Parameter Sets: LiteralPath
Aliases: FullName, PSPath
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -PassThru
Gets or sets whether to return the live imported slides in selection order.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath, SourceDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Gets or sets the source file path with FileSystem wildcard expansion.

```yaml
Type: String
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Position
Gets or sets the one-based insertion position; omission appends.

```yaml
Type: Int32
Parameter Sets: Path, LiteralPath, SourceDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SourceDocument
Gets or sets the live source, which remains open after import.

```yaml
Type: PowerPointDocument
Parameter Sets: SourceDocument
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SourceSlide
Gets or sets one-based source positions in selection order; omission imports every slide.

```yaml
Type: Int32[]
Parameter Sets: Path, LiteralPath, SourceDocument
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
- `System.String`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointSlide`

## RELATED LINKS

- None
