---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeAccess
## SYNOPSIS
Creates a native Access database without a provider, seed file or execution.

## SYNTAX
### Live (Default)
```powershell
New-OfficeAccess [[-Path] <string>] [-Force] [-Title <string>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Content
```powershell
New-OfficeAccess [-Path] <string> [-Content] <scriptblock> [-AllowLoss] [-PassThru] [-Force] [-Title <string>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
A path associates a live model without writing it; a content block atomically saves and closes it. The .mdb or .accdb extension selects the native file family.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficeAccess -Path ./Contacts.accdb {
    AccessTable Contacts { AccessColumn ContactId -Type AutoNumber; AccessColumn DisplayName -MaxLength 100 }
}
```


## PARAMETERS

### -AllowLoss
Accept reductions reported by native save.

```yaml
Type: SwitchParameter
Parameter Sets: Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Content
Authoring block completed before atomic publication.

```yaml
Type: ScriptBlock
Parameter Sets: Content
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Force
Permit replacement of the associated destination.

```yaml
Type: SwitchParameter
Parameter Sets: Live, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Return the saved FileInfo after the block completes.

```yaml
Type: SwitchParameter
Parameter Sets: Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Optional live association; required for a completed content block.

```yaml
Type: String
Parameter Sets: Live, Content
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Title
Database title.

```yaml
Type: String
Parameter Sets: Live, Content
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

- `OfficeIMO.Access.AccessDocument`
- `System.IO.FileInfo`

## RELATED LINKS

- None
