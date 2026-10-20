---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeAccessRow
## SYNOPSIS
Streams native or modeled table rows as PowerShell objects.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeAccessRow [-TableName] <string> -Document <AccessDocument> [-MaxRows <long>] [<CommonParameters>]
```

## DESCRIPTION
Streams native or modeled table rows as PowerShell objects.

## EXAMPLES

### EXAMPLE 1
```powershell
$database | Get-OfficeAccessRow -TableName Contacts | Export-Csv ./contacts.csv -NoTypeInformation
```


## PARAMETERS

### -Document
Live database document.

```yaml
Type: AccessDocument
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -MaxRows
Maximum rows emitted; a partial result produces a warning.

```yaml
Type: Int64
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TableName
Table name.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
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

- `OfficeIMO.Access.AccessDocument`

## OUTPUTS

- `System.Management.Automation.PSObject`

## RELATED LINKS

- None
