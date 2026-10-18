---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeVisioStencilShape
## SYNOPSIS
Adds a stencil shape to the current Visio page.

## SYNTAX
### CatalogName (Default)
```powershell
Add-OfficeVisioStencilShape [-Stencil] <string> [[-Text] <string>] [-Page <VisioPage>] [-Catalog <string>] [-Key <string>] [-X <double>] [-Y <double>] [-Width <Double>] [-Height <Double>] [-ShapeName <string>] [-NameU <string>] [-FillColor <string>] [-LineColor <string>] [-LineWeight <Double>] [-LinePattern <Int32>] [-FillPattern <Int32>] [-Angle <Double>] [-PassThru] [<CommonParameters>]
```

### CatalogObject
```powershell
Add-OfficeVisioStencilShape [-Stencil] <string> [[-Text] <string>] -CatalogObject <VisioStencilCatalog> [-Page <VisioPage>] [-Key <string>] [-X <double>] [-Y <double>] [-Width <Double>] [-Height <Double>] [-ShapeName <string>] [-NameU <string>] [-FillColor <string>] [-LineColor <string>] [-LineWeight <Double>] [-LinePattern <Int32>] [-FillPattern <Int32>] [-Angle <Double>] [-PassThru] [<CommonParameters>]
```

### BuiltIn
```powershell
Add-OfficeVisioStencilShape [-Stencil] <string> [[-Text] <string>] [-Page <VisioPage>] [-BuiltIn <OfficeVisioBuiltInStencilCatalog>] [-Key <string>] [-X <double>] [-Y <double>] [-Width <Double>] [-Height <Double>] [-ShapeName <string>] [-NameU <string>] [-FillColor <string>] [-LineColor <string>] [-LineWeight <Double>] [-LinePattern <Int32>] [-FillPattern <Int32>] [-Angle <Double>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a stencil shape to the current Visio page.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> OfficeVisio -Path .\StencilFlow.vsdx -UseMastersByDefault {
    VisioStencilImport -BuiltIn Flowchart -Name Flow -Default
    VisioStencil -Catalog Flow -Stencil process -Key intake -Text 'Intake' -X 1.5 -Y 4
}
```

Registers a built-in catalog and places a stencil shape on the active page.

## PARAMETERS

### -Angle
Shape angle in radians.

```yaml
Type: Double
Parameter Sets: CatalogName, CatalogObject, BuiltIn
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BuiltIn
Built-in OfficeIMO stencil catalog containing the shape.

```yaml
Type: OfficeVisioBuiltInStencilCatalog
Parameter Sets: BuiltIn
Aliases: None
Possible values: All, BasicShapes, Flowchart, BlockDiagram, Architecture, Network, Infrastructure, Cloud, SecurityIdentity, ContainersKubernetes, DataPlatform, CollaborationBusiness, Swimlane, OrgChart, Timeline, Sequence

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Catalog
Catalog previously registered in the active Visio DSL scope.

```yaml
Type: String
Parameter Sets: CatalogName
Aliases: CatalogName
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CatalogObject
Catalog object containing the stencil shape.

```yaml
Type: VisioStencilCatalog
Parameter Sets: CatalogObject
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FillColor
Fill color name or hex value.

```yaml
Type: String
Parameter Sets: CatalogName, CatalogObject, BuiltIn
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FillPattern
Native Visio fill-pattern index: 0 has no fill, 1 is solid, and 2 through 40 select built-in patterns.

```yaml
Type: Int32
Parameter Sets: CatalogName, CatalogObject, BuiltIn
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Height
Optional shape height. Omit to use the stencil default height.

```yaml
Type: Double
Parameter Sets: CatalogName, CatalogObject, BuiltIn
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Key
DSL key used by connector commands.

```yaml
Type: String
Parameter Sets: CatalogName, CatalogObject, BuiltIn
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
Parameter Sets: CatalogName, CatalogObject, BuiltIn
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
Parameter Sets: CatalogName, CatalogObject, BuiltIn
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
Parameter Sets: CatalogName, CatalogObject, BuiltIn
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NameU
Optional universal shape name.

```yaml
Type: String
Parameter Sets: CatalogName, CatalogObject, BuiltIn
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
Parameter Sets: CatalogName, CatalogObject, BuiltIn
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
Parameter Sets: CatalogName, CatalogObject, BuiltIn
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ShapeName
Optional shape name.

```yaml
Type: String
Parameter Sets: CatalogName, CatalogObject, BuiltIn
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Stencil
Stencil id, name, master name, keyword, alias, or tag.

```yaml
Type: String
Parameter Sets: CatalogName, CatalogObject, BuiltIn
Aliases: Name
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Text placed inside the shape. Omit to use the stencil display name.

```yaml
Type: String
Parameter Sets: CatalogName, CatalogObject, BuiltIn
Aliases: None
Possible values:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Optional shape width. Omit to use the stencil default width.

```yaml
Type: Double
Parameter Sets: CatalogName, CatalogObject, BuiltIn
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -X
X coordinate of the stencil shape center.

```yaml
Type: Double
Parameter Sets: CatalogName, CatalogObject, BuiltIn
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Y
Y coordinate of the stencil shape center.

```yaml
Type: Double
Parameter Sets: CatalogName, CatalogObject, BuiltIn
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

- `OfficeIMO.Visio.VisioPage`

## OUTPUTS

- `OfficeIMO.Visio.VisioShape`

## RELATED LINKS

- None
