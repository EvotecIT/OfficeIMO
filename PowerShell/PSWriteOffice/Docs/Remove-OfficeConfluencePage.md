---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Remove-OfficeConfluencePage
## SYNOPSIS
Plans or deletes a Confluence Cloud page.

## SYNTAX
### __AllParameterSets
```powershell
Remove-OfficeConfluencePage [-PageId] <string> [-Session <ConfluenceSession>] [-Purge] [-Draft] [-PlanOnly] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Plans or deletes a Confluence Cloud page.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Remove-OfficeConfluencePage -PageId 12345 -Purge -PlanOnly
```

Returns the exact DELETE request plan without using a session.

### EXAMPLE 2
```powershell
PS> Remove-OfficeConfluencePage -Session $session -PageId 12345
```

Uses PowerShell ShouldProcess before sending the non-retried delete request.

## PARAMETERS

### -Draft
Delete a draft page.

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

### -PageId
Page identifier.

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

### -PassThru
Return the completed delete plan after a live operation.

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

### -PlanOnly
Return the exact delete plan without contacting Confluence.

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

### -Purge
Permanently delete a page that is already in the trash.

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

### -Session
Configured session required for a live delete operation.

```yaml
Type: ConfluenceSession
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

- `OfficeIMO.Confluence.ConfluencePageWritePlan`

## RELATED LINKS

- None
