---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficePowerPointOpenDocumentOptions
## SYNOPSIS
Creates PowerPoint/OpenDocument conversion settings.

## SYNTAX
### Default (Default)
```powershell
New-OfficePowerPointOpenDocumentOptions [-LossPolicy <OfficeConversionLossPolicy>] [-NoImages] [-NoSpeakerNotes] [-NoBasicFormatting] [-MaxTableRows <Int32>] [-MaxTableColumns <Int32>] [<CommonParameters>]
```

## DESCRIPTION
Creates PowerPoint/OpenDocument conversion settings.

## EXAMPLES

### EXAMPLE 1
```powershell
$options = New-OfficePowerPointOpenDocumentOptions -NoSpeakerNotes
Convert-OfficeDocument -Path ./Review.pptx -To OpenDocument -Options $options
```


## PARAMETERS

### -LossPolicy
Gets or sets whether known omitted content is rejected or reported; omission retains the adapter default.

```yaml
Type: OfficeConversionLossPolicy
Parameter Sets: Default
Aliases: None
Possible values: Block, Allow

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxTableColumns
Gets or sets the maximum number of columns in a converted table.

```yaml
Type: Int32
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxTableRows
Gets or sets the maximum number of rows in a converted table.

```yaml
Type: Int32
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoBasicFormatting
Gets or sets whether common fill, outline and run formatting is omitted.

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

### -NoImages
Gets or sets whether embedded images are omitted.

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

### -NoSpeakerNotes
Gets or sets whether speaker notes are omitted.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `OfficeIMO.PowerPoint.OpenDocument.PowerPointOpenDocumentConversionOptions`

## RELATED LINKS

- None
