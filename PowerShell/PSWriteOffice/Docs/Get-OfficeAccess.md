---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeAccess
## SYNOPSIS
Reads a native database without running queries, macros, or linked connections.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeAccess [-Path] <string> [-TableName <string[]>] [-MaxInputBytes <long>] [-MaxRows <long>] [<CommonParameters>]
```

## DESCRIPTION
Reads a native database without running queries, macros, or linked connections.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-OfficeAccess -Path ./contacts.accdb -TableName Contacts
```


## PARAMETERS

### -MaxInputBytes
Maximum source size in bytes.

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

### -MaxRows
Maximum rows traversed by each native reader.

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

### -Path
Literal database filename.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: FullName, LiteralPath, FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue, ByPropertyName)
Accept wildcard characters: False
```

### -TableName
Optional user-table selection.

```yaml
Type: String[]
Parameter Sets: __AllParameterSets
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

- `System.String`

## OUTPUTS

- `OfficeIMO.Access.AccessDocument`

## RELATED LINKS

- None
