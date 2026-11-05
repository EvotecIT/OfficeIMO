---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeOneNoteNotebookOptions
## SYNOPSIS
Creates native OneNoteNotebookReaderOptions settings for document operations.

## SYNTAX
### Default (Default)
```powershell
New-OfficeOneNoteNotebookOptions [-NoSectionContent] [-StopOnSectionError] [-NoSectionGroupRecursion] [-IncludeRecycleBin] [-MaxSectionGroupDepth <Int32>] [-MaxNotebookEntries <Int32>] [<CommonParameters>]
```

## DESCRIPTION
Creates native OneNoteNotebookReaderOptions settings for document operations.

## EXAMPLES

### EXAMPLE 1
```powershell
$options = New-OfficeOneNoteNotebookOptions
```


## PARAMETERS

### -IncludeRecycleBin
IncludeRecycleBin override; omission preserves the native default.

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

### -MaxNotebookEntries
MaxNotebookEntries override; omission preserves the native default.

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

### -MaxSectionGroupDepth
MaxSectionGroupDepth override; omission preserves the native default.

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

### -NoSectionContent
NoSectionContent override; omission preserves the native default.

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

### -NoSectionGroupRecursion
NoSectionGroupRecursion override; omission preserves the native default.

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

### -StopOnSectionError
StopOnSectionError override; omission preserves the native default.

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

- `OfficeIMO.OneNote.OneNoteNotebookReaderOptions`

## RELATED LINKS

- None
