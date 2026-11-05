---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeMarkdownListItem
## SYNOPSIS
Adds a plain or rich item to a MarkdownList block.

## SYNTAX
### Text (Default)
```powershell
Add-OfficeMarkdownListItem [-Text] <string> [<CommonParameters>]
```

### Content
```powershell
Add-OfficeMarkdownListItem [-Content] <scriptblock> [<CommonParameters>]
```

## DESCRIPTION
Adds a plain or rich item to a MarkdownList block.

## EXAMPLES

### EXAMPLE 1
```powershell
MarkdownList { MarkdownListItem { MarkdownText 'Review' -Bold } }
```


## PARAMETERS

### -Content
Styled MarkdownText content.

```yaml
Type: ScriptBlock
Parameter Sets: Content
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Plain item text.

```yaml
Type: String
Parameter Sets: Text
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `System.Void`

## RELATED LINKS

- None
