---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeConfluencePage
## SYNOPSIS
Reads one page or streams a filtered Confluence Cloud page listing.

## SYNTAX
### List (Default)
```powershell
Get-OfficeConfluencePage -Session <ConfluenceSession> [-SpaceId <string>] [-Title <string>] [-Cursor <string>] [-Limit <int>] [-AsPage] [-BodyFormat <ConfluenceBodyFormat>] [-AsMarkdown] [-AsHtml] [-FailOnLoss] [<CommonParameters>]
```

### ById
```powershell
Get-OfficeConfluencePage [-PageId] <string> -Session <ConfluenceSession> [-BodyFormat <ConfluenceBodyFormat>] [-AsMarkdown] [-AsHtml] [-FailOnLoss] [<CommonParameters>]
```

## DESCRIPTION
Reads one page or streams a filtered Confluence Cloud page listing.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficeConfluencePage -Session $session -PageId 12345 -AsMarkdown
```

Returns the converted Markdown value and its ADF conversion report.

## PARAMETERS

### -AsHtml
Project each page body to HTML and include conversion evidence.

```yaml
Type: SwitchParameter
Parameter Sets: List, ById
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AsMarkdown
Project each page body to Markdown and include conversion evidence.

```yaml
Type: SwitchParameter
Parameter Sets: List, ById
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AsPage
Return listing batches rather than enumerating their pages.

```yaml
Type: SwitchParameter
Parameter Sets: List
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BodyFormat
Body representation requested from Confluence.

```yaml
Type: ConfluenceBodyFormat
Parameter Sets: List, ById
Aliases: None
Possible values: Storage, AtlasDocFormat

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Cursor
Optional cursor at which to resume a page listing.

```yaml
Type: String
Parameter Sets: List
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FailOnLoss
Throw when a requested projection reports reduced fidelity.

```yaml
Type: SwitchParameter
Parameter Sets: List, ById
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Limit
Maximum pages requested in each listing batch.

```yaml
Type: Int32
Parameter Sets: List
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PageId
Page identifier.

```yaml
Type: String
Parameter Sets: ById
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Session
Configured Confluence session.

```yaml
Type: ConfluenceSession
Parameter Sets: List, ById
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SpaceId
Optional space identifier used when listing pages.

```yaml
Type: String
Parameter Sets: List
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Title
Optional exact title used when listing pages.

```yaml
Type: String
Parameter Sets: List
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

- `None`

## OUTPUTS

- `OfficeIMO.Confluence.ConfluencePage`
- `OfficeIMO.Confluence.ConfluencePageBatch`
- `OfficeIMO.Confluence.ConfluenceContentConversionResult[System.String]`

## RELATED LINKS

- None
