---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Test-OfficeToml
## SYNOPSIS
Tests whether TOML files or text load, without changing them.

## SYNTAX
### Path (Default)
```powershell
Test-OfficeToml [-Path] <string[]> [-Untrusted] [-Detailed] [<CommonParameters>]
```

### LiteralPath
```powershell
Test-OfficeToml -LiteralPath <string[]> [-Untrusted] [-Detailed] [<CommonParameters>]
```

### Text
```powershell
Test-OfficeToml -Text <string> [-Untrusted] [-Detailed] [<CommonParameters>]
```

## DESCRIPTION
Returns $true or $false for each input, as Test-Json does. Detailed returns a result with the error's
diagnostic code, message, and 1-based line and column. Loading checks all of TOML 1.1, including duplicate keys and
the rules for defining tables; Untrusted applies the engine's limits for files from outside.

## EXAMPLES

### EXAMPLE 1
```powershell
if (-not (Test-OfficeToml -Path ./pyproject.toml)) { throw 'pyproject.toml is not valid TOML.' }
```


### EXAMPLE 2
```powershell
Get-ChildItem ./crates -Recurse -Filter *.toml | Test-OfficeToml -Detailed |
                Where-Object { -not $_.IsValid } |
                Select-Object -Property Path, Line, Column, Code, Message
```


## PARAMETERS

### -Detailed
Return a result with the error code and position instead of a Boolean.

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
Literal TOML files, including FileInfo.FullName and provider PSPath.

```yaml
Type: String[]
Parameter Sets: LiteralPath
Aliases: FullName, PSPath
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Path
TOML files with FileSystem wildcard expansion.

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

### -Text
TOML text to test.

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

### -Untrusted
Caps every limit at the engine's untrusted defaults, for files from outside.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String[]`

## OUTPUTS

- `System.Boolean`
- `PSWriteOffice.Models.Toml.OfficeTomlTestResult`: The outcome of Test-OfficeToml -Detailed for one input.

## RELATED LINKS

- None
