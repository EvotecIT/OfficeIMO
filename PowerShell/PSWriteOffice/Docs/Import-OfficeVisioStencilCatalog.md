---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Import-OfficeVisioStencilCatalog
## SYNOPSIS
Loads a stencil catalog (built-in, from .vssx/.vsdx packages or installed Visio); inside a diagram block it also registers it.

## SYNTAX
### BuiltIn (Default)
```powershell
Import-OfficeVisioStencilCatalog [-BuiltIn <OfficeVisioBuiltInStencilCatalog>] [-Name <string>] [-Default] [<CommonParameters>]
```

### Path
```powershell
Import-OfficeVisioStencilCatalog [-Path] <string[]> [-Name <string>] [-Default] [-Recurse] [-CatalogName <string>] [-Category <string>] [-IdPrefix <string>] [-MasterName <string[]>] [-IncludeUnsupportedMasters] [-NoLearnMasterDimensions] [-NoPreviewImageMetadata] [-NoConnectionPointMetadata] [-DefaultWidth <OfficeLength>] [-DefaultHeight <OfficeLength>] [<CommonParameters>]
```

### LiteralPath
```powershell
Import-OfficeVisioStencilCatalog -LiteralPath <string[]> [-Name <string>] [-Default] [-Recurse] [-CatalogName <string>] [-Category <string>] [-IdPrefix <string>] [-MasterName <string[]>] [-IncludeUnsupportedMasters] [-NoLearnMasterDimensions] [-NoPreviewImageMetadata] [-NoConnectionPointMetadata] [-DefaultWidth <OfficeLength>] [-DefaultHeight <OfficeLength>] [<CommonParameters>]
```

### Installed
```powershell
Import-OfficeVisioStencilCatalog -Installed [-Name <string>] [-Default] [-CatalogName <string>] [-Category <string>] [-IdPrefix <string>] [-MasterName <string[]>] [-IncludeUnsupportedMasters] [-NoLearnMasterDimensions] [-NoPreviewImageMetadata] [-NoConnectionPointMetadata] [-DefaultWidth <OfficeLength>] [-DefaultHeight <OfficeLength>] [<CommonParameters>]
```

### Catalog
```powershell
Import-OfficeVisioStencilCatalog -Catalog <VisioStencilCatalog> [-Name <string>] [-Default] [<CommonParameters>]
```

## DESCRIPTION
Loads a stencil catalog (built-in, from .vssx/.vsdx packages or installed Visio); inside a diagram block it also registers it.

## EXAMPLES

### EXAMPLE 1
```powershell
$network = Import-OfficeVisioStencilCatalog -BuiltIn Network
```


## PARAMETERS

### -BuiltIn
Default All.

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
Catalog for this operation.

```yaml
Type: VisioStencilCatalog
Parameter Sets: Catalog
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -CatalogName
Name given to the loaded catalog.

```yaml
Type: String
Parameter Sets: Path, LiteralPath, Installed
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Category
Category for this operation.

```yaml
Type: String
Parameter Sets: Path, LiteralPath, Installed
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Default
Inside a diagram block: make it the default catalog.

```yaml
Type: SwitchParameter
Parameter Sets: BuiltIn, Path, LiteralPath, Installed, Catalog
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DefaultHeight
Default 0.9in.

```yaml
Type: OfficeLength
Parameter Sets: Path, LiteralPath, Installed
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DefaultWidth
Default 1.8in.

```yaml
Type: OfficeLength
Parameter Sets: Path, LiteralPath, Installed
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IdPrefix
IdPrefix for this operation.

```yaml
Type: String
Parameter Sets: Path, LiteralPath, Installed
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeUnsupportedMasters
IncludeUnsupportedMasters for this operation.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath, Installed
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Installed
Installed for this operation.

```yaml
Type: SwitchParameter
Parameter Sets: Installed
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LiteralPath
LiteralPath for this operation.

```yaml
Type: String[]
Parameter Sets: LiteralPath
Aliases: PSPath, FullName
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -MasterName
MasterName for this operation.

```yaml
Type: String[]
Parameter Sets: Path, LiteralPath, Installed
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Name
Inside a diagram block: key under which the catalog is registered for VisioStencil -Catalog.

```yaml
Type: String
Parameter Sets: BuiltIn, Path, LiteralPath, Installed, Catalog
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoConnectionPointMetadata
NoConnectionPointMetadata for this operation.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath, Installed
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoLearnMasterDimensions
NoLearnMasterDimensions for this operation.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath, Installed
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoPreviewImageMetadata
NoPreviewImageMetadata for this operation.

```yaml
Type: SwitchParameter
Parameter Sets: Path, LiteralPath, Installed
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Path for this operation.

```yaml
Type: String[]
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Recurse
Recurse for this operation.

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

- `System.String[]`
- `OfficeIMO.Visio.Stencils.VisioStencilCatalog`

## OUTPUTS

- `OfficeIMO.Visio.Stencils.VisioStencilCatalog`

## RELATED LINKS

- None
