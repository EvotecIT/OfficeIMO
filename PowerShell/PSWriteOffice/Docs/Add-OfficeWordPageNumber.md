---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeWordPageNumber
## SYNOPSIS
Adds a PAGE field to the current header/footer.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeWordPageNumber [-IncludeTotalPages] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Typically used inside WordFooter to render “Page X of Y”.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> WordFooter { Add-OfficeWordPageNumber -IncludeTotalPages }
```

Outputs “Page # of #” in the footer.

## PARAMETERS

### -IncludeTotalPages
Include “of N” when true.

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

### -PassThru
Emit the object created or changed by the command.

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

- `None`

## RELATED LINKS

- None
