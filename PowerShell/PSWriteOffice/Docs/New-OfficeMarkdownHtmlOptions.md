---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeMarkdownHtmlOptions
## SYNOPSIS
Creates native Markdown HTML document, styling and resource-link settings.

## SYNTAX
### Default (Default)
```powershell
New-OfficeMarkdownHtmlOptions [-ReaderOptions <MarkdownLoadOptions>] [-Theme <OfficeThemeInput>] [-DocumentMode] [-Style <MarkdownHtmlStyle>] [-CssDelivery <MarkdownCssDelivery>] [-AssetMode <MarkdownAssetMode>] [-Title <string>] [-RawHtmlHandling <MarkdownRawHtmlHandling>] [-ExternalLinksRel <string>] [-ExternalLinksReferrerPolicy <string>] [-ImagesReferrerPolicy <string>] [-IncludeAnchorLinks] [-GitHubTaskListHtml] [-GitHubFootnoteHtml] [-ExternalLinksTargetBlank] [-RestrictHttpLinksToBaseOrigin] [-RestrictHttpImagesToBaseOrigin] [-BlockExternalHttpImages] [-ImagesLoadingLazy] [-ImagesDecodingAsync] [-AllowedHttpLinkHost <string[]>] [-AllowedHttpImageHost <string[]>] [<CommonParameters>]
```

## DESCRIPTION
Creates native Markdown HTML document, styling and resource-link settings.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficeMarkdownHtmlOptions -AllowedHttpImageHost @('Value')
```


## PARAMETERS

### -AllowedHttpImageHost
Additional allowed HTTP hosts.

```yaml
Type: String[]
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AllowedHttpLinkHost
Additional allowed HTTP hosts.

```yaml
Type: String[]
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AssetMode
Native AssetMode setting.

```yaml
Type: MarkdownAssetMode
Parameter Sets: Default
Aliases: None
Possible values: Online, Offline

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BlockExternalHttpImages
Enable or explicitly disable BlockExternalHttpImages.

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

### -CssDelivery
Native CssDelivery setting.

```yaml
Type: MarkdownCssDelivery
Parameter Sets: Default
Aliases: None
Possible values: Inline, ExternalFile, LinkHref, None

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DocumentMode
Render a complete HTML document; explicit false renders a fragment.

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

### -ExternalLinksReferrerPolicy
Native ExternalLinksReferrerPolicy setting.

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

### -ExternalLinksRel
Native ExternalLinksRel setting.

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

### -ExternalLinksTargetBlank
Enable or explicitly disable ExternalLinksTargetBlank.

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

### -GitHubFootnoteHtml
Enable or explicitly disable GitHubFootnoteHtml.

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

### -GitHubTaskListHtml
Enable or explicitly disable GitHubTaskListHtml.

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

### -ImagesDecodingAsync
Enable or explicitly disable ImagesDecodingAsync.

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

### -ImagesLoadingLazy
Enable or explicitly disable ImagesLoadingLazy.

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

### -ImagesReferrerPolicy
Native ImagesReferrerPolicy setting.

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

### -IncludeAnchorLinks
Enable or explicitly disable IncludeAnchorLinks.

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

### -RawHtmlHandling
Native RawHtmlHandling setting.

```yaml
Type: MarkdownRawHtmlHandling
Parameter Sets: Default
Aliases: None
Possible values: Allow, Escape, Sanitize, Strip

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ReaderOptions
Native syntax, normalization, URL and input-limit policy, captured independently.

```yaml
Type: MarkdownLoadOptions
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RestrictHttpImagesToBaseOrigin
Enable or explicitly disable RestrictHttpImagesToBaseOrigin.

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

### -RestrictHttpLinksToBaseOrigin
Enable or explicitly disable RestrictHttpLinksToBaseOrigin.

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

### -Style
Native Style setting.

```yaml
Type: MarkdownHtmlStyle
Parameter Sets: Default
Aliases: None
Possible values: Plain, Clean, GithubLight, GithubDark, GithubAuto, ChatLight, ChatDark, ChatAuto, Word

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Theme
Built-in theme, shared theme object or theme template.

```yaml
Type: OfficeThemeInput
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Title
Native Title setting.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `OfficeIMO.Markdown.MarkdownHtmlExportOptions`

## RELATED LINKS

- None
