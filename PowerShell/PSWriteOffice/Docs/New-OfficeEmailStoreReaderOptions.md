---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficeEmailStoreReaderOptions
## SYNOPSIS
Creates bounded email-store reader settings without requiring .NET constructor syntax.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeEmailStoreReaderOptions [-MaxInputBytes <long>] [-MaxNodeCount <int>] [-MaxBTreeDepth <int>] [-MaxCachedBTreePages <int>] [-MaxFolderCount <int>] [-MaxItemCount <int>] [-MaxPropertiesPerItem <int>] [-MaxDecodedPropertyBytesPerItem <long>] [-MaxAttachmentsPerItem <int>] [-MaxAttachmentBytes <long>] [-MaxTotalAttachmentBytes <long>] [-ExcludeAttachmentContent] [-PstPassword <string>] [-PstPasswordEncoding <string>] [-IncludeAssociatedItems] [-IncludeOrphanedItems] [-MaxNestedMessageDepth <int>] [-MaxArchiveEntries <int>] [-MaxArchiveEntryBytes <long>] [-MaxArchiveDecodedBytes <long>] [-MaxXmlCharactersPerItem <long>] [-MaxMessageBytes <long>] [-MaxDirectoryDepth <int>] [-MaxDirectoryFileCount <int>] [-MaxDecodedTableBytes <long>] [<CommonParameters>]
```

## DESCRIPTION
Creates bounded email-store reader settings without requiring .NET constructor syntax.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $options = New-OfficeEmailStoreReaderOptions -ExcludeAttachmentContent -MaxAttachmentsPerItem 100
Get-OfficeEmail -Path .\Message.emlx -StoreOptions $options -AsResult
```


## PARAMETERS

### -ExcludeAttachmentContent
Do not retain decoded attachment payloads in memory.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeAssociatedItems
Materialize folder-associated information items.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeOrphanedItems
Recover item nodes absent from folder contents tables.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxArchiveDecodedBytes
Maximum total decoded size declared by archive entries.

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

### -MaxArchiveEntries
Maximum entries accepted from a compressed email-store archive.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxArchiveEntryBytes
Maximum decoded size declared by one archive entry.

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

### -MaxAttachmentBytes
Maximum decoded bytes in one attachment.

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

### -MaxAttachmentsPerItem
Maximum attachments per item.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxBTreeDepth
Maximum tree traversal depth.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxCachedBTreePages
Maximum PST/OST B-tree pages retained by the cache.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxDecodedPropertyBytesPerItem
Maximum decoded property bytes per item.

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

### -MaxDecodedTableBytes
Maximum decoded bytes traversed from one PST/OST table data tree.

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

### -MaxDirectoryDepth
Maximum directory depth traversed by mailbox-directory sessions.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxDirectoryFileCount
Maximum EML, EMLX, and Maildir files indexed by one directory session.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxFolderCount
Maximum folders materialized.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxInputBytes
Maximum seekable source length.

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

### -MaxItemCount
Maximum items materialized.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxMessageBytes
Maximum RFC 5322/MIME message bytes accepted from one item.

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

### -MaxNestedMessageDepth
Maximum embedded-message recursion depth.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxNodeCount
Maximum NDB nodes and blocks visited.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxPropertiesPerItem
Maximum MAPI properties decoded per item.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxTotalAttachmentBytes
Maximum decoded attachment bytes across the read.

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

### -MaxXmlCharactersPerItem
Maximum XML characters parsed from one archive item.

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

### -PstPassword
Password used to validate legacy protected PST files.

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

### -PstPasswordEncoding
Encoding name used for the legacy PST password checksum.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `OfficeIMO.Email.Store.EmailStoreReaderOptions`

## RELATED LINKS

- None
