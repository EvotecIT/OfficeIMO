---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Open-OfficeIWork
## SYNOPSIS
Reads a supported Apple Pages, Numbers, or Keynote package without launching iWork.

## SYNTAX
### Path (Default)
```powershell
Open-OfficeIWork [-Path] <string> [-Kind <IWorkDocumentKind>] [-Options <IWorkLoadOptions>] [<CommonParameters>]
```

### LiteralPath
```powershell
Open-OfficeIWork -LiteralPath <string> [-Kind <IWorkDocumentKind>] [-Options <IWorkLoadOptions>] [<CommonParameters>]
```

## DESCRIPTION
Reads supported iWork packages, including the qualified iWork '09 XML profile, into OfficeIMO's bounded source model.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $source = Open-OfficeIWork -Path .\Quarterly.numbers
$source | Select-Object Kind, ContainerKind, BuildVersions
$source.ReadNumbers().Sheets | Select-Object Name
```

Returns OfficeIMO's bounded, loss-aware source model.

## PARAMETERS

### -Kind
Optional expected application kind; a mismatch is rejected.

```yaml
Type: IWorkDocumentKind
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values: Pages, Numbers, Keynote

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LiteralPath
Exact input path without wildcard expansion.

```yaml
Type: String
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
Bounded OfficeIMO iWork read options.

```yaml
Type: IWorkLoadOptions
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
Path to a supported Pages, Numbers, or Keynote package or directory bundle, including the qualified iWork '09 XML profile.

```yaml
Type: String
Parameter Sets: Path
Aliases: None
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
