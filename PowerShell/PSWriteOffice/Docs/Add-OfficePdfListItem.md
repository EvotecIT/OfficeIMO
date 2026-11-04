---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficePdfListItem
## SYNOPSIS
Adds one list item, from -Text or from PdfText runs.

## SYNTAX
### Text (Default)
```powershell
Add-OfficePdfListItem [-Text] <string> [<CommonParameters>]
```

### Content
```powershell
Add-OfficePdfListItem [-Content] <scriptblock> [<CommonParameters>]
```

## DESCRIPTION
Adds one list item, from -Text or from PdfText runs.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficePdf -Path .\Checklist.pdf {
    PdfList {
        PdfListItem 'Backups verified'
        PdfListItem 'Monitoring silenced'
    }
}
```


## PARAMETERS

### -Content
Content for this operation.

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
Text for this operation.

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
