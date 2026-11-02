---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeDocumentCapability
## SYNOPSIS
Lists installed reading, conversion and protection capabilities from their canonical catalogs.

## SYNTAX
### Read (Default)
```powershell
Get-OfficeDocumentCapability [-Reader <OfficeDocumentReader>] [-Origin <ReaderHandlerOrigin>] [-Manifest] [<CommonParameters>]
```

### Conversion
```powershell
Get-OfficeDocumentCapability -Conversion [-From <string>] [-To <string>] [-Id <string[]>] [-Executable] [<CommonParameters>]
```

### Protection
```powershell
Get-OfficeDocumentCapability [[-Id] <string[]>] -Protection [-Format <string>] [-Kind <OfficeProtectionKind>] [-IncompleteOnly] [-AsJson] [-AsMarkdown] [<CommonParameters>]
```

## DESCRIPTION
Lists installed reading, conversion and protection capabilities from their canonical catalogs.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-OfficeDocumentCapability -Conversion -From Word -To Pdf -Executable | Format-List *
```


### EXAMPLE 2
```powershell
Get-OfficeDocumentCapability -Protection -Format PDF
```


## PARAMETERS

### -AsJson
Return the complete protection catalog as deterministic JSON, without row filters.

```yaml
Type: SwitchParameter
Parameter Sets: Protection
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AsMarkdown
Return the complete protection catalog as Markdown, without row filters.

```yaml
Type: SwitchParameter
Parameter Sets: Protection
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Conversion
List conversion routes and their fidelity and support evidence.

```yaml
Type: SwitchParameter
Parameter Sets: Conversion
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Executable
Return only conversion routes implemented by this installed host.

```yaml
Type: SwitchParameter
Parameter Sets: Conversion
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Format
Protection format identifier or family, such as PDF or DOCX.

```yaml
Type: String
Parameter Sets: Protection
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -From
Source format family or exact extension, with or without a leading dot.

```yaml
Type: String
Parameter Sets: Conversion
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Id
Exact stable route or protection capability identifiers.

```yaml
Type: String[]
Parameter Sets: Conversion, Protection
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -IncompleteOnly
Return protection rows with unsupported operations.

```yaml
Type: SwitchParameter
Parameter Sets: Protection
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Kind
Protected-content mechanism.

```yaml
Type: OfficeProtectionKind
Parameter Sets: Protection
Aliases: None
Possible values: PasswordEncryption, RecipientEncryption, DigitalSignature, FontObfuscation, EditingRestriction, AccessDeterrence

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Manifest
Return the Reader capability manifest instead of individual handlers.

```yaml
Type: SwitchParameter
Parameter Sets: Read
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Origin
Return handlers supplied by OfficeIMO or registered by a caller.

```yaml
Type: ReaderHandlerOrigin
Parameter Sets: Read
Aliases: None
Possible values: Custom, OfficeIMO

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Protection
List password, encryption, signature and restriction support.

```yaml
Type: SwitchParameter
Parameter Sets: Protection
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Reader
Immutable caller-configured Reader registry.

```yaml
Type: OfficeDocumentReader
Parameter Sets: Read
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -To
Destination format family or exact extension, with or without a leading dot.

```yaml
Type: String
Parameter Sets: Conversion
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

- `OfficeIMO.Reader.ReaderHandlerCapability`
- `OfficeIMO.Reader.ReaderCapabilityManifest`
- `OfficeIMO.OfficeConversionCapability`
- `OfficeIMO.Security.OfficeProtectionCapability`
- `System.String`

## RELATED LINKS

- None
