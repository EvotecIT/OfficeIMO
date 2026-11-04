---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeVisioConnector
## SYNOPSIS
Adds a connector between two Visio shapes.

## SYNTAX
### Page (Default)
```powershell
Add-OfficeVisioConnector -From <Object> -To <Object> [-Page <VisioPage>] [-Kind <ConnectorKind>] [-FromSide <VisioSide>] [-ToSide <VisioSide>] [-Label <string>] [-LineColor <OfficeColor>] [-LineWeight <OfficeLength>] [-LinePattern <Int32>] [-BeginArrow <EndArrow>] [-EndArrow <EndArrow>] [-PassThru] [<CommonParameters>]
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
Parameter Sets: Page
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
Parameter Sets: Page
Aliases: None
Possible values: None, Arrow, Triangle

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -From
Source shape object or a registered key, id or name.

```yaml
Type: Object
Parameter Sets: Page
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
Parameter Sets: Page
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
Parameter Sets: Page
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
Parameter Sets: Page
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
Type: OfficeColor
Parameter Sets: Page
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
Parameter Sets: Page
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
Type: OfficeLength
Parameter Sets: Page
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
Parameter Sets: Page
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
Parameter Sets: Page
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -To
Target shape object or a registered key, id or name.

```yaml
Type: Object
Parameter Sets: Page
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
Parameter Sets: Page
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
