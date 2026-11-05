---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Open-OfficeRtf
## SYNOPSIS
Reads an RTF file or RTF text into a read result with the document, embedded payloads and diagnostics.

## SYNTAX
### Path (Default)
```powershell
Open-OfficeRtf [-Path] <string> [-AsResult] [<CommonParameters>]
```

### LiteralPath
```powershell
Open-OfficeRtf -LiteralPath <string> [-AsResult] [<CommonParameters>]
```

### Text
```powershell
Open-OfficeRtf -Text <string> [-AsResult] [<CommonParameters>]
```

## DESCRIPTION
Reads an RTF file or RTF text into a read result with the document, embedded payloads and diagnostics.

## EXAMPLES

### EXAMPLE 1
```powershell
$rtf = Open-OfficeRtf -Path .\Contract.rtf
```


## PARAMETERS

### -AsResult
AsResult for this operation.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LiteralPath
LiteralPath for this operation.

```yaml
Type: String
Parameter Sets: LiteralPath
Aliases: PSPath, FullName
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Path
Path for this operation.

```yaml
Type: String
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Text
Text for this operation.

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

- `OfficeIMO.Rtf.RtfDocument`

## RELATED LINKS

- None
