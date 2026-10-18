---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeVisioConnector
## SYNOPSIS
Adds a connector between two Visio shapes.

## SYNTAX
### ByKey (Default)
```powershell
Add-OfficeVisioConnector -From <string> -To <string> [-Page <VisioPage>] [-Kind <ConnectorKind>] [-FromSide <VisioSide>] [-ToSide <VisioSide>] [-Label <string>] [-LineColor <string>] [-LineWeight <Double>] [-LinePattern <Int32>] [-BeginArrow <EndArrow>] [-EndArrow <EndArrow>] [-PassThru] [<CommonParameters>]
```

### ByShape
```powershell
Add-OfficeVisioConnector -FromShape <VisioShape> -ToShape <VisioShape> [-Page <VisioPage>] [-Kind <ConnectorKind>] [-FromSide <VisioSide>] [-ToSide <VisioSide>] [-Label <string>] [-LineColor <string>] [-LineWeight <Double>] [-LinePattern <Int32>] [-BeginArrow <EndArrow>] [-EndArrow <EndArrow>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a connector between two Visio shapes.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> OfficeVisio -Path .\Flow.vsdx {
    VisioRectangle -Key source -Text 'Source' -X 1 -Y 4
    VisioRectangle -Key target -Text 'Target' -X 4 -Y 4
    VisioConnector -From source -To target -Kind RightAngle -EndArrow Triangle -Label 'sync'
}
```

Adds a routed connector between shapes registered in the current DSL scope.

## PARAMETERS

### -BeginArrow
Begin arrow style.

```yaml
Type: EndArrow
Parameter Sets: ByKey, ByShape
Aliases: None
Possible values: None, Arrow, Triangle

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -EndArrow
End arrow style.

```yaml
Type: EndArrow
Parameter Sets: ByKey, ByShape
Aliases: None
Possible values: None, Arrow, Triangle

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -From
Source shape key, id, or name.

```yaml
Type: String
Parameter Sets: ByKey
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FromShape
Source shape object.

```yaml
Type: VisioShape
Parameter Sets: ByShape
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FromSide
Preferred source shape side.

```yaml
Type: VisioSide
Parameter Sets: ByKey, ByShape
Aliases: None
Possible values: Auto, Left, Right, Bottom, Top

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Kind
Connector kind.

```yaml
Type: ConnectorKind
Parameter Sets: ByKey, ByShape
Aliases: None
Possible values: Straight, RightAngle, Curved, Dynamic

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Label
Connector label.

```yaml
Type: String
Parameter Sets: ByKey, ByShape
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LineColor
Line color name or hex value.

```yaml
Type: String
Parameter Sets: ByKey, ByShape
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LinePattern
Native Visio line-pattern index: 0 hides the line, 1 is solid, and 2 through 23 select built-in patterns.

```yaml
Type: Int32
Parameter Sets: ByKey, ByShape
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LineWeight
Line weight.

```yaml
Type: Double
Parameter Sets: ByKey, ByShape
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Page
Target page. Optional inside VisioPage or OfficeVisio.

```yaml
Type: VisioPage
Parameter Sets: ByKey, ByShape
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -PassThru
Emit the object created or changed by the command.

```yaml
Type: SwitchParameter
Parameter Sets: ByKey, ByShape
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -To
Target shape key, id, or name.

```yaml
Type: String
Parameter Sets: ByKey
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ToShape
Target shape object.

```yaml
Type: VisioShape
Parameter Sets: ByShape
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ToSide
Preferred target shape side.

```yaml
Type: VisioSide
Parameter Sets: ByKey, ByShape
Aliases: None
Possible values: Auto, Left, Right, Bottom, Top

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Visio.VisioPage`

## OUTPUTS

- `OfficeIMO.Visio.VisioConnector`

## RELATED LINKS

- None
