---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Import-OfficeConversionReport
## SYNOPSIS
Reads saved conversion run reports without executing jobs or changing outputs.

## SYNTAX
### Path (Default)
```powershell
Import-OfficeConversionReport [-Path] <string[]> [<CommonParameters>]
```

### LiteralPath
```powershell
Import-OfficeConversionReport -LiteralPath <string[]> [<CommonParameters>]
```

## DESCRIPTION
Reports describe an earlier attempt; checkpoint receipts and current source/output validation determine whether artifacts can be reused.

## EXAMPLES

### EXAMPLE 1
```powershell
Import-OfficeConversionReport -Path ./runs/*.json | Select-Object Status, Summary
```


### EXAMPLE 2
```powershell
Get-ChildItem ./runs -Filter *.json -File | Import-OfficeConversionReport
```


## PARAMETERS

### -LiteralPath
Literal report paths, including FileInfo.FullName and provider PSPath pipeline properties.

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

### -Path
Saved run-report JSON paths, including provider wildcards.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String[]`

## OUTPUTS

- `OfficeIMO.OfficeBatchReport`

## RELATED LINKS

- None
