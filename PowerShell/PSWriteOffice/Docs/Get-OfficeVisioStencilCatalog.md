---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeVisioStencilCatalog
## SYNOPSIS
Gets built-in or package-backed OfficeIMO Visio stencil catalogs.

## SYNTAX
### BuiltIn (Default)
```powershell
Get-OfficeVisioStencilCatalog [-BuiltIn <OfficeVisioBuiltInStencilCatalog>] [<CommonParameters>]
```

### Path
```powershell
Get-OfficeVisioStencilCatalog [-Path] <string[]> [-Recurse] [-CatalogName <string>] [-Category <string>] [-IdPrefix <string>] [-MasterName <string[]>] [-IncludeUnsupportedMasters] [-NoLearnMasterDimensions] [-NoPreviewImageMetadata] [-NoConnectionPointMetadata] [-DefaultWidth <double>] [-DefaultHeight <double>] [<CommonParameters>]
```

### Installed
```powershell
Get-OfficeVisioStencilCatalog -Installed [-CatalogName <string>] [-Category <string>] [-IdPrefix <string>] [-MasterName <string[]>] [-IncludeUnsupportedMasters] [-NoLearnMasterDimensions] [-NoPreviewImageMetadata] [-NoConnectionPointMetadata] [-DefaultWidth <double>] [-DefaultHeight <double>] [<CommonParameters>]
```

## DESCRIPTION
Gets built-in or package-backed OfficeIMO Visio stencil catalogs.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $catalog = Get-OfficeVisioStencilCatalog -BuiltIn Flowchart
Find-OfficeVisioStencil -Catalog $catalog -Query decision -First 3
```

Gets a built-in catalog and searches it before using stencils in a diagram.

## PARAMETERS

### -BuiltIn
Built-in OfficeIMO stencil catalog to return.

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

### -CatalogName
Catalog display name when loading package metadata.

```yaml
Type: String
Parameter Sets: Path, Installed
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Category
Category assigned to package-backed stencil shapes.

```yaml
Type: String
Parameter Sets: Path, Installed
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DefaultHeight
Default height for package-backed stencils when dimensions cannot be learned.

```yaml
Type: Double
Parameter Sets: Path, Installed
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DefaultWidth
Default width for package-backed stencils when dimensions cannot be learned.

```yaml
Type: Double
Parameter Sets: Path, Installed
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IdPrefix
Stable id prefix for package-backed stencil shapes.

```yaml
Type: String
Parameter Sets: Path, Installed
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeUnsupportedMasters
Include unsupported package masters as generic generated masters.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Installed
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Installed
Discover installed Microsoft Visio stencils and templates.

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

### -MasterName
Optional master filters for package-backed catalogs.

```yaml
Type: String[]
Parameter Sets: Path, Installed
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoConnectionPointMetadata
Skip reading source connection point metadata from package master parts.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Installed
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoLearnMasterDimensions
Skip reading master dimensions from package master parts.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Installed
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoPreviewImageMetadata
Skip reading preview image relationship metadata from package master parts.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Installed
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Visio package, package directory, or OfficeIMO native stencil manifest path.

```yaml
Type: String[]
Parameter Sets: Path
Aliases: FilePath, LiteralPath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Recurse
Search directories recursively.

```yaml
Type: SwitchParameter
Parameter Sets: Path
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

- `OfficeIMO.Visio.Stencils.VisioStencilCatalog`

## RELATED LINKS

- None
