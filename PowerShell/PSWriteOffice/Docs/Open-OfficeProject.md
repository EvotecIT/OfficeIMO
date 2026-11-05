---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Open-OfficeProject
## SYNOPSIS
Loads a project without calculating or executing its content.

## SYNTAX
### Path (Default)
```powershell
Open-OfficeProject [-Path] <string> [[-Content] <scriptblock>] [-PassThru] [-AllowLoss] [-ReadOnly] [-MaxInputBytes <long>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPath
```powershell
Open-OfficeProject [[-Content] <scriptblock>] -LiteralPath <string> [-PassThru] [-AllowLoss] [-ReadOnly] [-MaxInputBytes <long>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Loads a project without calculating or executing its content.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeProject -Path ./delivery.mpp -ReadOnly
```


## PARAMETERS

### -AllowLoss
Explicitly accept native-save loss for an authoring block.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Content
An optional authoring block; omission returns the live document.

```yaml
Type: ScriptBlock
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LiteralPath
Exact source path without wildcard expansion.

```yaml
Type: String
Parameter Sets: LiteralPath
Aliases: PSPath, FullName
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -MaxInputBytes
Maximum source size in bytes.

```yaml
Type: Int64
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the FileInfo after a completed authoring block.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Literal filesystem path.

```yaml
Type: String
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -ReadOnly
Open without allowing edits.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath
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

- `System.String`

## OUTPUTS

- `OfficeIMO.Project.ProjectDocument`

## RELATED LINKS

- None
