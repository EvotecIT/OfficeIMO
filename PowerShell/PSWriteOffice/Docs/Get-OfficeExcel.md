---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeExcel
## SYNOPSIS
Opens an existing Excel workbook.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeExcel [-Path] <string> [-ReadOnly] [-Password <string>] [<CommonParameters>]
```

### Uri
```powershell
Get-OfficeExcel [-Uri] <uri> [-AllowHttp] [-ReadOnly] [-Password <string>] [<CommonParameters>]
```

## DESCRIPTION
Returns the underlying ExcelDocument so callers can inspect or reuse it in DSL pipelines.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $workbook = Get-OfficeExcel -Path .\report.xlsx -ReadOnly
```

Loads report.xlsx for inspection without enabling writes.

## PARAMETERS

### -AllowHttp
Allow HTTP workbook downloads in addition to HTTPS.

```yaml
Type: SwitchParameter
Parameter Sets: Uri
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Password used to open an encrypted workbook package.

```yaml
Type: String
Parameter Sets: Path, Uri
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Path to the workbook to load.

```yaml
Type: String
Parameter Sets: Path
Aliases: InputPath, FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ReadOnly
Open the file in read-only mode.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Uri
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Uri
Remote workbook URI to load.

```yaml
Type: Uri
Parameter Sets: Uri
Aliases: Url
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `None`

## RELATED LINKS

- None
