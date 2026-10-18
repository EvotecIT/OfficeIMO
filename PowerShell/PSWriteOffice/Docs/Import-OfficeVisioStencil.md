---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Import-OfficeVisioStencil
## SYNOPSIS
Registers a stencil catalog with the active Visio DSL scope.

## SYNTAX
### Catalog (Default)
```powershell
Import-OfficeVisioStencil -Catalog <VisioStencilCatalog> [-Name <string>] [-Default] [<CommonParameters>]
```

### BuiltIn
```powershell
Import-OfficeVisioStencil -BuiltIn <OfficeVisioBuiltInStencilCatalog> [-Name <string>] [-Default] [<CommonParameters>]
```

### Path
```powershell
Import-OfficeVisioStencil [-Path] <string[]> [-Name <string>] [-Default] [-Recurse] [-LoadCatalogName <string>] [-Category <string>] [-IdPrefix <string>] [-MasterName <string[]>] [-IncludeUnsupportedMasters] [-NoLearnMasterDimensions] [-NoPreviewImageMetadata] [-NoConnectionPointMetadata] [-DefaultWidth <double>] [-DefaultHeight <double>] [<CommonParameters>]
```

### Installed
```powershell
Import-OfficeVisioStencil -Installed [-Name <string>] [-Default] [-LoadCatalogName <string>] [-Category <string>] [-IdPrefix <string>] [-MasterName <string[]>] [-IncludeUnsupportedMasters] [-NoLearnMasterDimensions] [-NoPreviewImageMetadata] [-NoConnectionPointMetadata] [-DefaultWidth <double>] [-DefaultHeight <double>] [<CommonParameters>]
```

## DESCRIPTION
Registers a stencil catalog with the active Visio DSL scope.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> OfficeVisio -Path .\Flow.vsdx -UseMastersByDefault {
    VisioStencilImport -BuiltIn Flowchart -Name Flow -Default
    VisioStencil -Stencil process -Key step -Text 'Step' -X 2 -Y 4
}
```

Registers the flowchart catalog and makes it the default for later VisioStencil calls.

## PARAMETERS

### -BuiltIn
Built-in OfficeIMO stencil catalog to register.

```yaml
Type: OfficeVisioBuiltInStencilCatalog
Parameter Sets: BuiltIn
Aliases: None
Possible values: All, BasicShapes, Flowchart, BlockDiagram, Architecture, Network, Infrastructure, Cloud, SecurityIdentity, ContainersKubernetes, DataPlatform, CollaborationBusiness, Swimlane, OrgChart, Timeline, Sequence

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Catalog
Catalog object to register.

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

### -Default
Make this catalog the default for later VisioStencil calls.

```yaml
Type: SwitchParameter
Parameter Sets: Catalog, BuiltIn, Path, Installed
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
Discover installed Microsoft Visio stencils and templates, then register the combined catalog.

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

### -LoadCatalogName
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

### -Name
Name used by VisioStencil -Catalog in the DSL.

```yaml
Type: String
Parameter Sets: Catalog, BuiltIn, Path, Installed
Aliases: CatalogName, Key
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
Visio package, package directory, or OfficeIMO native stencil manifest path to load and register.

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

- `OfficeIMO.Visio.Stencils.VisioStencilCatalog`

## OUTPUTS

- `OfficeIMO.Visio.Stencils.VisioStencilCatalog`

## RELATED LINKS

- None
