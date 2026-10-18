---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Send-OfficeConfluenceAttachment
## SYNOPSIS
Uploads or versions a Confluence page attachment.

## SYNTAX
### __AllParameterSets
```powershell
Send-OfficeConfluenceAttachment [-Path] <string> -Session <ConfluenceSession> -PageId <string> [-FileName <string>] [-ContentType <string>] [-Comment <string>] [-MinorEdit <bool>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Uploads or versions a Confluence page attachment.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Send-OfficeConfluenceAttachment -Session $session -PageId 12345 -Path .\report.xlsx -Comment 'Daily refresh'
```

Uses Confluence's multipart attachment endpoint without automatically retrying the write.

## PARAMETERS

### -Comment
Optional attachment version comment.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ContentType
MIME content type.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FileName
Optional attachment file name. Defaults to the local file name.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MinorEdit
Whether the attachment update is a minor edit.

```yaml
Type: Boolean
Parameter Sets: __AllParameterSets
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
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Local file to upload.

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

### -Session
Configured Confluence session.

```yaml
Type: ConfluenceSession
Parameter Sets: __AllParameterSets
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

## RELATED LINKS

- None
