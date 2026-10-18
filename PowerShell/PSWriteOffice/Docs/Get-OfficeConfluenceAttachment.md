---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeConfluenceAttachment
## SYNOPSIS
Lists or downloads Confluence page attachments.

## SYNTAX
### List (Default)
```powershell
Get-OfficeConfluenceAttachment -Session <ConfluenceSession> -PageId <string> [-Cursor <string>] [-Limit <int>] [-AsPage] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Download
```powershell
Get-OfficeConfluenceAttachment -Session <ConfluenceSession> -PageId <string> -AttachmentId <string> [-OutFile <string>] [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Lists or downloads Confluence page attachments.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficeConfluenceAttachment -Session $session -PageId 12345
```

Follows attachment cursors and streams attachment metadata.

## PARAMETERS

### -AsPage
Return attachment batches rather than individual metadata objects.

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

### -AttachmentId
Attachment identifier to download.

```yaml
Type: String
Parameter Sets: Download
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Cursor
Optional cursor at which to resume attachment listing.

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

### -Force
Overwrite an existing destination file.

```yaml
Type: SwitchParameter
Parameter Sets: Download
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Limit
Maximum attachments requested per listing batch.

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

### -OutFile
Optional destination path. Without this parameter, the download is returned as one byte array.

```yaml
Type: String
Parameter Sets: Download
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
Parameter Sets: List, Download
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Session
Configured Confluence session.

```yaml
Type: ConfluenceSession
Parameter Sets: List, Download
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

- `None`

## OUTPUTS

- `OfficeIMO.Confluence.ConfluenceAttachment`
- `OfficeIMO.Confluence.ConfluenceAttachmentBatch`
- `System.Byte[]`
- `System.IO.FileInfo`

## RELATED LINKS

- None
