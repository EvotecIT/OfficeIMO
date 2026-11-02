---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Search-OfficeDocument
## SYNOPSIS
Finds regular-expression or literal matches across supported files and folders.

## SYNTAX
### Path (Default)
```powershell
Search-OfficeDocument [-Path] <string[]> [-Pattern] <string[]> [-SimpleMatch] [-CaseSensitive] [-WholeWord] [-MaximumResults <int>] [-Recurse] [-Extension <string[]>] [-MaximumFiles <int>] [-MaximumStoreItems <int>] [-IncludePageLocations] [-Reader <OfficeDocumentReader>] [-Password <securestring>] [-ThrottleLimit <int>] [<CommonParameters>]
```

### InputObject
```powershell
Search-OfficeDocument [-InputObject] <OfficeDocumentReadResult> [-Pattern] <string[]> [-SimpleMatch] [-CaseSensitive] [-WholeWord] [-MaximumResults <int>] [<CommonParameters>]
```

### LiteralPath
```powershell
Search-OfficeDocument [-Pattern] <string[]> -LiteralPath <string[]> [-SimpleMatch] [-CaseSensitive] [-WholeWord] [-MaximumResults <int>] [-Recurse] [-Extension <string[]>] [-MaximumFiles <int>] [-MaximumStoreItems <int>] [-IncludePageLocations] [-Reader <OfficeDocumentReader>] [-Password <securestring>] [-ThrottleLimit <int>] [<CommonParameters>]
```

## DESCRIPTION
Finds regular-expression or literal matches across supported files and folders.

## EXAMPLES

### EXAMPLE 1
```powershell
Search-OfficeDocument ./Reports -Recurse -Pattern 'INV-\d+'
```


### EXAMPLE 2
```powershell
Read-OfficeDocument ./Report.docx | Search-OfficeDocument -Pattern 'retention period' -SimpleMatch
```


## PARAMETERS

### -CaseSensitive
Use case-sensitive matching.

```yaml
Type: SwitchParameter
Parameter Sets: Path, InputObject, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Extension
Selected source extensions, with or without leading dots.

```yaml
Type: String[]
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludePageLocations
Compute Word and RTF page locations when supported.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputObject
Normalized document to search without rereading a file.

```yaml
Type: OfficeDocumentReadResult
Parameter Sets: InputObject
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -LiteralPath
Gets or sets input paths without wildcard expansion, binding FileInfo.FullName and provider PSPath.

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

### -MaximumFiles
Maximum selected files; zero means no limit.

```yaml
Type: Int32
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumResults
Maximum total matches per document across all patterns; zero means no limit.

```yaml
Type: Int32
Parameter Sets: Path, InputObject, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumStoreItems
Maximum items projected from each mail store; zero means all.

```yaml
Type: Int32
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Open password for encrypted sources.

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
Gets or sets input paths with provider wildcard expansion.

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

### -Pattern
One or more regular expressions, or literal strings with SimpleMatch.

```yaml
Type: String[]
Parameter Sets: Path, InputObject, LiteralPath
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Reader
Use an immutable, caller-configured Reader.

```yaml
Type: OfficeDocumentReader
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Recurse
Include subfolders.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SimpleMatch
Treat patterns as literal text.

```yaml
Type: SwitchParameter
Parameter Sets: Path, InputObject, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ThrottleLimit
Gets or sets the maximum number of concurrent background operations.

```yaml
Type: Int32
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WholeWord
Accept matches only at word boundaries.

```yaml
Type: SwitchParameter
Parameter Sets: Path, InputObject, LiteralPath
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

- `OfficeIMO.Reader.OfficeDocumentReadResult`
- `System.String[]`

## OUTPUTS

- `PSWriteOffice.Models.Reader.OfficeDocumentSearchMatch`: PowerShell-friendly occurrence returned by a path-based document search.

## RELATED LINKS

- None
