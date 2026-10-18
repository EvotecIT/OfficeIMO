---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Test-OfficePdfRewrite
## SYNOPSIS
Creates a proof report for user-visible signals preserved by a PDF rewrite.

## SYNTAX
### __AllParameterSets
```powershell
Test-OfficePdfRewrite [-ReferencePath] <string> [-DifferencePath] <string> [-Options <PdfRewritePreservationOptions>] [-FailOnLoss] [-ReferencePassword <string>] [-IgnoreReferencePermissionRestrictions] [-DifferencePassword <string>] [-IgnoreDifferencePermissionRestrictions] [<CommonParameters>]
```

## DESCRIPTION
Creates a proof report for user-visible signals preserved by a PDF rewrite.

## EXAMPLES

### EXAMPLE 1
```powershell
Test-OfficePdfRewrite -ReferencePath 'C:\Path' -DifferencePath 'C:\Path'
```


## PARAMETERS

### -DifferencePassword
Password used to authenticate the rewritten PDF.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DifferencePath
Rewritten PDF path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FailOnLoss
Throw when preservation checks find a mismatch.

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

### -IgnoreDifferencePermissionRestrictions
After authentication, explicitly ignore restrictions on the rewritten PDF.

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

### -IgnoreReferencePermissionRestrictions
After authentication, explicitly ignore restrictions on the original PDF.

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

### -Options
Optional required preservation signals and limits.

```yaml
Type: PdfRewritePreservationOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ReferencePassword
Password used to authenticate the original PDF.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ReferencePath
Original PDF path.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `OfficeIMO.Pdf.PdfRewritePreservationReport`

## RELATED LINKS

- None
