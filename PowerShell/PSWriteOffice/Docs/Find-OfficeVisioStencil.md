---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Find-OfficeVisioStencil
## SYNOPSIS
Searches OfficeIMO Visio stencil catalogs.

## SYNTAX
### Query (Default)
```powershell
Find-OfficeVisioStencil [[-Query] <string>] [-Catalog <VisioStencilCatalog>] [-First <int>] [<CommonParameters>]
```

### CatalogName
```powershell
Find-OfficeVisioStencil [[-Query] <string>] [-Catalog <VisioStencilCatalog>] [-CatalogName <string>] [-First <int>] [<CommonParameters>]
```

### BuiltIn
```powershell
Find-OfficeVisioStencil [[-Query] <string>] [-Catalog <VisioStencilCatalog>] [-BuiltIn <OfficeVisioBuiltInStencilCatalog>] [-First <int>] [<CommonParameters>]
```

## DESCRIPTION
Searches OfficeIMO Visio stencil catalogs.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $flow = Get-OfficeVisioStencilCatalog -BuiltIn Flowchart
Find-OfficeVisioStencil -Catalog $flow -Query process -First 5
```

Returns matching stencil definitions that can be used by VisioStencil.

## PARAMETERS

### -BuiltIn
Built-in OfficeIMO stencil catalog to search.

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
Catalog object to search. Defaults to the combined built-in catalog.

```yaml
Type: VisioStencilCatalog
Parameter Sets: Query, CatalogName, BuiltIn
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -CatalogName
Catalog previously registered in the active Visio DSL scope.

```yaml
Type: String
Parameter Sets: CatalogName
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -First
Maximum number of shapes to return.

```yaml
Type: Int32
Parameter Sets: Query, CatalogName, BuiltIn
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Query
Search text. Empty search returns catalog contents.

```yaml
Type: String
Parameter Sets: Query, CatalogName, BuiltIn
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Visio.Stencils.VisioStencilCatalog`

## OUTPUTS

- `OfficeIMO.Visio.Stencils.VisioStencilShape`

## RELATED LINKS

- None
