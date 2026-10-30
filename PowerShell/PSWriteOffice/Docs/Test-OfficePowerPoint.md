---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Test-OfficePowerPoint
## SYNOPSIS
Inspects a live or saved presentation without changing or reopening a supplied document.

## SYNTAX
### Document (Default)
```powershell
Test-OfficePowerPoint [-Document <PowerPointPresentation>] [-Options <Object[]>] [<CommonParameters>]
```

### Path
```powershell
Test-OfficePowerPoint [-Path] <string[]> [-Options <Object[]>] [-Password <securestring>] [-ThrottleLimit <int>] [<CommonParameters>]
```

### LiteralPath
```powershell
Test-OfficePowerPoint -LiteralPath <string[]> [-Options <Object[]>] [-Password <securestring>] [-ThrottleLimit <int>] [<CommonParameters>]
```

## DESCRIPTION
Inspects a live or saved presentation without changing or reopening a supplied document.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-ChildItem ./Publish -Filter *.pptx | Test-OfficePowerPoint
```


### EXAMPLE 2
```powershell
$policy = New-OfficePowerPointInspectionOptions -Features -Visuals
Test-OfficePowerPoint -Document $deck -Options $policy
```


## PARAMETERS

### -Document
Gets or sets a live document; omission uses the enclosing presentation block.

```yaml
Type: PowerPointPresentation
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -LiteralPath
Gets or sets input paths without wildcard expansion, binding FileInfo.FullName and provider PSPath.

```yaml
Type: String[]
Parameter Sets: LiteralPath
Aliases: PSPath, FullName
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Options
Gets or sets inspection, native format or Reader policy objects, including hashtables.

```yaml
Type: Object[]
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Gets or sets the password for opening saved presentations.

```yaml
Type: SecureString
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Gets or sets input paths with provider wildcard expansion.

```yaml
Type: String[]
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -ThrottleLimit
Gets or sets the maximum number of concurrent background operations.

```yaml
Type: Int32
Parameter Sets: Path, LiteralPath
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

- `OfficeIMO.PowerPoint.PowerPointPresentation`
- `System.String[]`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointInspectionReport`

## RELATED LINKS

- None
