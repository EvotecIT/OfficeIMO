---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficeEmailReaderOptions
## SYNOPSIS
Creates bounded EML, MSG, and TNEF reader settings through ordinary PowerShell parameters.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeEmailReaderOptions [-MaxInputBytes <long>] [-MaxHeaderBytes <int>] [-MaxHeaderCount <int>] [-MaxPartCount <int>] [-MaxMimeDepth <int>] [-MaxAttachmentBytes <long>] [-MaxTotalAttachmentBytes <long>] [-MaxNestedMessageDepth <int>] [-ExcludeAttachmentContent] [-PreserveRawSource] [-MaxCompoundDirectoryEntries <int>] [-MaxMapiPropertyCount <int>] [-MaxDecodedPropertyBytes <long>] [-MaxTnefAttributeCount <int>] [-MaxAttachmentCount <int>] [<CommonParameters>]
```

## DESCRIPTION
Creates bounded EML, MSG, and TNEF reader settings through ordinary PowerShell parameters.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $options = New-OfficeEmailReaderOptions -ExcludeAttachmentContent -MaxAttachmentBytes 25MB
Get-OfficeEmail -Path .\Message.msg -Options $options -AsResult
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

### -MaxAttachmentBytes
Maximum decoded bytes for one attachment.

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

### -MaxAttachmentCount
Maximum aggregate attachment count.

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

### -MaxCompoundDirectoryEntries
Maximum CFB directory entries accepted while reading MSG.

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

### -MaxDecodedPropertyBytes
Maximum aggregate bytes represented by decoded MSG property streams.

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

### -MaxHeaderBytes
Maximum bytes allowed in one MIME header section.

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

### -MaxHeaderCount
Maximum number of header fields in one entity.

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
Maximum artifact size accepted by the reader.

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

### -MaxMapiPropertyCount
Maximum aggregate MAPI properties across a message and embedded messages.

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

### -MaxMimeDepth
Maximum nested MIME depth.

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

### -MaxPartCount
Maximum MIME entity count.

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

### -MaxTnefAttributeCount
Maximum number of TNEF attributes.

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
Maximum aggregate decoded attachment bytes.

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

### -PreserveRawSource
Retain original artifact bytes for an explicit lossless write.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `OfficeIMO.Email.EmailReaderOptions`

## RELATED LINKS

- None
