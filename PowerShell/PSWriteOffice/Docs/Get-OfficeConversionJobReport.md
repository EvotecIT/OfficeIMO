---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeConversionJobReport
## SYNOPSIS
Reads a bounded saved conversion run report without executing its job or changing outputs.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeConversionJobReport [-Path] <string> [<CommonParameters>]
```

## DESCRIPTION
Reports describe an earlier attempt; checkpoint receipts and current source/output validation determine whether artifacts can be reused.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-OfficeConversionJobReport -Path ./runs/monthly.json
            Invoke-OfficeConversionJob -Path ./monthly.officejob.json -Plan
```


## PARAMETERS

### -Path
Literal run-report JSON path. Reading does not open referenced sources, outputs or jobs.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: LiteralPath, FullName
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String`

## OUTPUTS

- `OfficeIMO.OfficeBatchReport`

## RELATED LINKS

- None
