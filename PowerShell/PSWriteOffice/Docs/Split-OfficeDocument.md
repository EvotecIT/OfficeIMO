---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Split-OfficeDocument
## SYNOPSIS
Splits PDF files into consecutive groups, selected ranges or top-level bookmarked chapters.

## SYNTAX
### PathByPageCount (Default)
```powershell
Split-OfficeDocument [-Path] <string[]> [-OutputDirectory] <string> [-PageCount <int>] [-Options <Object[]>] [-Password <securestring>] [-Prefix <string>] [-PadIndex] [-IndexWidth <int>] [-NoClobber] [-ThrottleLimit <int>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### PathByRange
```powershell
Split-OfficeDocument [-Path] <string[]> [-OutputDirectory] <string> -Page <string[]> [-Options <Object[]>] [-Password <securestring>] [-Prefix <string>] [-PadIndex] [-IndexWidth <int>] [-NoClobber] [-ThrottleLimit <int>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### PathByBookmark
```powershell
Split-OfficeDocument [-Path] <string[]> [-OutputDirectory] <string> -ByBookmark [-Bookmark <string[]>] [-Options <Object[]>] [-Password <securestring>] [-Prefix <string>] [-PadIndex] [-IndexWidth <int>] [-NoClobber] [-ThrottleLimit <int>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPathByPageCount
```powershell
Split-OfficeDocument [-OutputDirectory] <string> -LiteralPath <string[]> [-PageCount <int>] [-Options <Object[]>] [-Password <securestring>] [-Prefix <string>] [-PadIndex] [-IndexWidth <int>] [-NoClobber] [-ThrottleLimit <int>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPathByRange
```powershell
Split-OfficeDocument [-OutputDirectory] <string> -LiteralPath <string[]> -Page <string[]> [-Options <Object[]>] [-Password <securestring>] [-Prefix <string>] [-PadIndex] [-IndexWidth <int>] [-NoClobber] [-ThrottleLimit <int>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPathByBookmark
```powershell
Split-OfficeDocument [-OutputDirectory] <string> -LiteralPath <string[]> -ByBookmark [-Bookmark <string[]>] [-Options <Object[]>] [-Password <securestring>] [-Prefix <string>] [-PadIndex] [-IndexWidth <int>] [-NoClobber] [-ThrottleLimit <int>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Each completed file is returned immediately. Native page extraction retains its permission,
signature and preservation policies. Only the selected files are written; other folder contents remain intact.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-ChildItem ./Scans -Filter *.pdf | Split-OfficeDocument -OutputDirectory ./Records -PageCount 2 -NoClobber
```


### EXAMPLE 2
```powershell
Split-OfficeDocument -LiteralPath './Manual [final].pdf' -OutputDirectory ./Chapters -ByBookmark
```


## PARAMETERS

### -Bookmark
Only these exact top-level bookmark titles; omission selects every chapter.

```yaml
Type: String[]
Parameter Sets: PathByBookmark, LiteralPathByBookmark
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ByBookmark
Split at top-level bookmarks; nested pages remain within their chapter.

```yaml
Type: SwitchParameter
Parameter Sets: PathByBookmark, LiteralPathByBookmark
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IndexWidth
Explicit numeric index width, up to 32 characters.

```yaml
Type: Int32
Parameter Sets: PathByPageCount, PathByRange, PathByBookmark, LiteralPathByPageCount, LiteralPathByRange, LiteralPathByBookmark
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LiteralPath
Literal files; also binds FileInfo.FullName and provider PSPath properties.

```yaml
Type: String[]
Parameter Sets: LiteralPathByPageCount, LiteralPathByRange, LiteralPathByBookmark
Aliases: PSPath, FullName
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -NoClobber
Refuse to replace an existing part, including a file created by another writer.

```yaml
Type: SwitchParameter
Parameter Sets: PathByPageCount, PathByRange, PathByBookmark, LiteralPathByPageCount, LiteralPathByRange, LiteralPathByBookmark
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Options
Native PDF load settings or property overrides; includes resource and permission policies.

```yaml
Type: Object[]
Parameter Sets: PathByPageCount, PathByRange, PathByBookmark, LiteralPathByPageCount, LiteralPathByRange, LiteralPathByBookmark
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputDirectory
Destination folder. File names begin with Prefix or the source base name.

```yaml
Type: String
Parameter Sets: PathByPageCount, PathByRange, PathByBookmark, LiteralPathByPageCount, LiteralPathByRange, LiteralPathByBookmark
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PadIndex
Pad numeric part indexes to the width of the source page count.

```yaml
Type: SwitchParameter
Parameter Sets: PathByPageCount, PathByRange, PathByBookmark, LiteralPathByPageCount, LiteralPathByRange, LiteralPathByBookmark
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Page
One output per range string, for example '1-3','4,6'.

```yaml
Type: String[]
Parameter Sets: PathByRange, LiteralPathByRange
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PageCount
Consecutive source pages per output; defaults to one.

```yaml
Type: Int32
Parameter Sets: PathByPageCount, LiteralPathByPageCount
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Input password; accepts a credential or warned plain text.

```yaml
Type: SecureString
Parameter Sets: PathByPageCount, PathByRange, PathByBookmark, LiteralPathByPageCount, LiteralPathByRange, LiteralPathByBookmark
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Input files in provider and pipeline order.

```yaml
Type: String[]
Parameter Sets: PathByPageCount, PathByRange, PathByBookmark
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Prefix
Portable output file-name prefix; defaults to the source base name.

```yaml
Type: String
Parameter Sets: PathByPageCount, PathByRange, PathByBookmark, LiteralPathByPageCount, LiteralPathByRange, LiteralPathByBookmark
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ThrottleLimit
Maximum concurrent input files.

```yaml
Type: Int32
Parameter Sets: PathByPageCount, PathByRange, PathByBookmark, LiteralPathByPageCount, LiteralPathByRange, LiteralPathByBookmark
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

- `System.IO.FileInfo`

## RELATED LINKS

- None
