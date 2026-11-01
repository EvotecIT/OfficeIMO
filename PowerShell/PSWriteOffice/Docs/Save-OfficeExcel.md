---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Save-OfficeExcel
## SYNOPSIS
Saves a live Excel workbook without closing it.

## SYNTAX
### Default (Default)
```powershell
Save-OfficeExcel [-Document] <ExcelDocument> [-Path <string>] [-Password <securestring>] [-SaveOptions <ExcelSaveOptions>] [-AllowLoss] [-Show] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Omitted Password preserves loaded encryption; an explicit null removes it. SaveOptions accepts native options or a hashtable. PassThru returns the saved FileInfo. The workbook must use explicit persistence.

## EXAMPLES

### EXAMPLE 1
```powershell
Save-OfficeExcel -Document $workbook -Path ./Reviewed.xlsx -SaveOptions @{ Preflight = $true; ValidateOpenXml = $true } -PassThru
            Close-OfficeExcel -Document $workbook
```


## PARAMETERS

### -AllowLoss
Explicitly accept reported native-save reductions; overrides SaveOptions when supplied.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Live workbook to save.

```yaml
Type: ExcelDocument
Parameter Sets: Default
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -PassThru
Return the saved FileInfo.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Override the output password; omission retains loaded encryption.

```yaml
Type: SecureString
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Optional new destination; omission uses the associated path.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SaveOptions
Native save settings, or a hashtable using the save-option factory's names.

```yaml
Type: ExcelSaveOptions
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Show
Open the saved file in its registered default application.

```yaml
Type: SwitchParameter
Parameter Sets: Default
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

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `System.IO.FileInfo`

## RELATED LINKS

- None
