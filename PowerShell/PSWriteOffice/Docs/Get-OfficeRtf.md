---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeRtf
## SYNOPSIS
Reads RTF into OfficeIMO's semantic and lossless syntax models.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeRtf [-Path] <string> [<CommonParameters>]
```

### Text
```powershell
Get-OfficeRtf -Text <string> [<CommonParameters>]
```

## DESCRIPTION
Reads RTF into OfficeIMO's semantic and lossless syntax models.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $rtf = Get-OfficeRtf -Path .\Report.rtf
$rtf.Document.Paragraphs[0].ToPlainText()
```

Reads an RTF file and returns the OfficeIMO RTF read result.

## PARAMETERS

### -Path
RTF file path.

```yaml
Type: String
Parameter Sets: Path
Aliases: FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Text
Raw RTF text.

```yaml
Type: String
Parameter Sets: Text
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

- `System.String`

## OUTPUTS

- `OfficeIMO.Rtf.RtfReadResult`

## RELATED LINKS

- None
