---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Repair-OfficeDocument
## SYNOPSIS
Runs native workbook repairs in place, into a copy, or on a caller-owned model.

## SYNTAX
### Path (Default)
```powershell
Repair-OfficeDocument [-Path] <string[]> [-OutputPath <string>] [-Options <Object[]>] [-Password <securestring>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPath
```powershell
Repair-OfficeDocument -LiteralPath <string[]> [-OutputPath <string>] [-Options <Object[]>] [-Password <securestring>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Repair-OfficeDocument [-Document] <Object> [-OutputPath <string>] [-Options <Object[]>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Runs native workbook repairs in place, into a copy, or on a caller-owned model.

## EXAMPLES

### EXAMPLE 1
```powershell
Repair-OfficeDocument -Path @('C:\Path')
```


### EXAMPLE 2
```powershell
Repair-OfficeDocument -Document 'Value'
```


### EXAMPLE 3
```powershell
Repair-OfficeDocument -LiteralPath @('C:\Path')
```


## PARAMETERS

### -Document
Caller-owned native document to repair without taking ownership.

```yaml
Type: Object
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -LiteralPath
Exact input paths.

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
Native format repair settings and property overrides.

```yaml
Type: Object[]
Parameter Sets: Path, LiteralPath, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputPath
Write one repaired copy; file inputs otherwise save in place.

```yaml
Type: String
Parameter Sets: Path, LiteralPath, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit native repair findings and publication details.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Password used to open an encrypted file.

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
Input paths; supports wildcards.

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
- `System.Object`

## OUTPUTS

- `PSWriteOffice.Models.OfficeRepairReport`: Native workbook repair findings and the destination actually published by Repair-OfficeDocument.

## RELATED LINKS

- None
