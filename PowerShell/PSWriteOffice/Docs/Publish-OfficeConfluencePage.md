---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Publish-OfficeConfluencePage
## SYNOPSIS
Plans, creates, or updates a Confluence Cloud page.

## SYNTAX
### Create (Default)
```powershell
Publish-OfficeConfluencePage -SpaceId <string> -Title <string> -Content <string> [-Session <ConfluenceSession>] [-ParentId <string>] [-ContentFormat <OfficeConfluenceContentFormat>] [-BodyFormat <ConfluenceBodyFormat>] [-PlanOnly] [-FailOnLoss] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Update
```powershell
Publish-OfficeConfluencePage -PageId <string> -VersionNumber <int> -Title <string> -Content <string> [-Session <ConfluenceSession>] [-VersionMessage <string>] [-ContentFormat <OfficeConfluenceContentFormat>] [-BodyFormat <ConfluenceBodyFormat>] [-PlanOnly] [-FailOnLoss] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Plans, creates, or updates a Confluence Cloud page.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Publish-OfficeConfluencePage -SpaceId 42 -Title 'Daily status' -Content $markdown -PlanOnly
```

Returns a serializable request plan and performs no network operation.

### EXAMPLE 2
```powershell
PS> Publish-OfficeConfluencePage -Session $session -PageId 123 -Title 'Daily status' -VersionNumber 8 -Content $markdown -VersionMessage 'automation refresh'
```

The version number must be the next Confluence page version.

## PARAMETERS

### -BodyFormat
Confluence representation to publish.

```yaml
Type: ConfluenceBodyFormat
Parameter Sets: Create, Update
Aliases: None
Possible values: Storage, AtlasDocFormat

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Content
Page content in the representation selected by ContentFormat.

```yaml
Type: String
Parameter Sets: Create, Update
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -ContentFormat
Representation of Content.

```yaml
Type: OfficeConfluenceContentFormat
Parameter Sets: Create, Update
Aliases: None
Possible values: Markdown, Html, AtlasDocFormat, Storage

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FailOnLoss
Throw when conversion reports reduced fidelity.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Update
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PageId
Page identifier for an update.

```yaml
Type: String
Parameter Sets: Update
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ParentId
Optional parent page identifier for a new page.

```yaml
Type: String
Parameter Sets: Create
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PlanOnly
Return the exact write plan without contacting Confluence.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Update
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Session
Configured session required for live create or update operations.

```yaml
Type: ConfluenceSession
Parameter Sets: Create, Update
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SpaceId
Space identifier for a new page.

```yaml
Type: String
Parameter Sets: Create
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Title
Page title.

```yaml
Type: String
Parameter Sets: Create, Update
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -VersionMessage
Optional version message for an update.

```yaml
Type: String
Parameter Sets: Update
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -VersionNumber
Next positive page version number for an update.

```yaml
Type: Int32
Parameter Sets: Update
Aliases: None
Possible values:

Required: True
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

- `OfficeIMO.Confluence.ConfluencePageWritePlan`
- `OfficeIMO.Confluence.ConfluencePage`

## RELATED LINKS

- None
