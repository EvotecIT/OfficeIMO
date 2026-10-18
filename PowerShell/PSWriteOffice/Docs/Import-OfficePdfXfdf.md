---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Import-OfficePdfXfdf
## SYNOPSIS
Imports bounded DTD-free XFDF through the validated PDF form filler.

## SYNTAX
### Text (Default)
```powershell
Import-OfficePdfXfdf [-Path] <string> -Xfdf <string> -OutputPath <string> [-Options <PdfFormFillerOptions>] [-ReadOptions <PdfLoadOptions>] [-Password <string>] [-IgnorePermissionRestrictions] [-MaxXfdfBytes <long>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### File
```powershell
Import-OfficePdfXfdf [-Path] <string> -XfdfPath <string> -OutputPath <string> [-Options <PdfFormFillerOptions>] [-ReadOptions <PdfLoadOptions>] [-Password <string>] [-IgnorePermissionRestrictions] [-MaxXfdfBytes <long>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Imports bounded DTD-free XFDF through the validated PDF form filler.

## EXAMPLES

### EXAMPLE 1
```powershell
Import-OfficePdfXfdf -Path 'C:\Path' -Xfdf 'Value' -OutputPath 'C:\Path'
```


### EXAMPLE 2
```powershell
Import-OfficePdfXfdf -Path 'C:\Path' -XfdfPath 'C:\Path' -OutputPath 'C:\Path'
```


## PARAMETERS

### -IgnorePermissionRestrictions
After successful password authentication, explicitly ignore owner-imposed form-modification restrictions.

```yaml
Type: SwitchParameter
Parameter Sets: Text, File
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxXfdfBytes
Maximum UTF-8 byte count accepted from an XFDF file or pipeline. Default: 4 MiB.

```yaml
Type: Int64
Parameter Sets: Text, File
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Options
Optional validated form filling behavior.

```yaml
Type: PdfFormFillerOptions
Parameter Sets: Text, File
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputPath
Destination PDF path.

```yaml
Type: String
Parameter Sets: Text, File
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Return the rewritten fluent PDF document.

```yaml
Type: SwitchParameter
Parameter Sets: Text, File
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Password used to authenticate the source form.

```yaml
Type: String
Parameter Sets: Text, File
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Source PDF path.

```yaml
Type: String
Parameter Sets: Text, File
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ReadOptions
Optional bounded PDF parsing and password settings for the source form.

```yaml
Type: PdfLoadOptions
Parameter Sets: Text, File
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Xfdf
XFDF XML.

```yaml
Type: String
Parameter Sets: Text
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -XfdfPath
Path to an XFDF file.

```yaml
Type: String
Parameter Sets: File
Aliases: None
Possible values:

Required: True
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

- `OfficeIMO.Pdf.PdfDocument`

## RELATED LINKS

- None
