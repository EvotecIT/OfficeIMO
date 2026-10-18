---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeIWork
## SYNOPSIS
Reads a modern Apple Pages, Numbers, or Keynote package without launching iWork.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeIWork [-Path] <string> [-Kind <IWorkDocumentKind>] [-Options <IWorkReadOptions>] [<CommonParameters>]
```

## DESCRIPTION
Reads a modern Apple Pages, Numbers, or Keynote package without launching iWork.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $source = Get-OfficeIWork -Path .\Quarterly.numbers
$source | Select-Object Kind, ContainerKind, BuildVersions
$source.ReadNumbers().Sheets | Select-Object Name
```

Returns OfficeIMO's bounded, loss-aware source model.

## PARAMETERS

### -Kind
Optional expected application kind; a mismatch is rejected.

```yaml
Type: IWorkDocumentKind
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Pages, Numbers, Keynote

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Options
Bounded OfficeIMO iWork read options.

```yaml
Type: IWorkReadOptions
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
Path to a modern Pages, Numbers, or Keynote package or directory bundle.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String`

## OUTPUTS

- `OfficeIMO.IWork.IWorkSourceDocument`

## RELATED LINKS

- None
