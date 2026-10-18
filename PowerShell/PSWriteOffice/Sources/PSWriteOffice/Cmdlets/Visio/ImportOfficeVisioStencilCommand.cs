using System.Management.Automation;
using OfficeIMO.Visio.Stencils;
using PSWriteOffice.Services.Visio;

namespace PSWriteOffice.Cmdlets.Visio;

/// <summary>Registers a stencil catalog with the active Visio DSL scope.</summary>
/// <example>
///   <summary>Register a catalog for DSL use.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>OfficeVisio -Path .\Flow.vsdx -UseMastersByDefault {
///     VisioStencilImport -BuiltIn Flowchart -Name Flow -Default
///     VisioStencil -Stencil process -Key step -Text 'Step' -X 2 -Y 4
/// }</code>
///   <para>Registers the flowchart catalog and makes it the default for later VisioStencil calls.</para>
/// </example>
[Cmdlet(VerbsData.Import, "OfficeVisioStencil", DefaultParameterSetName = CatalogParameterSet)]
[Alias("VisioStencilImport", "Import-VisioStencil")]
[OutputType(typeof(VisioStencilCatalog))]
public sealed class ImportOfficeVisioStencilCommand : PSCmdlet
{
    private const string CatalogParameterSet = "Catalog";
    private const string BuiltInParameterSet = "BuiltIn";
    private const string PathParameterSet = "Path";
    private const string InstalledParameterSet = "Installed";

    /// <summary>Catalog object to register.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = CatalogParameterSet)]
    public VisioStencilCatalog? Catalog { get; set; }

    /// <summary>Built-in OfficeIMO stencil catalog to register.</summary>
    [Parameter(Mandatory = true, ParameterSetName = BuiltInParameterSet)]
    public OfficeVisioBuiltInStencilCatalog BuiltIn { get; set; }

    /// <summary>Visio package, package directory, or OfficeIMO native stencil manifest path to load and register.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = PathParameterSet)]
    [Alias("FilePath", "LiteralPath")]
    public string[] Path { get; set; } = [];

    /// <summary>Discover installed Microsoft Visio stencils and templates, then register the combined catalog.</summary>
    [Parameter(Mandatory = true, ParameterSetName = InstalledParameterSet)]
    public SwitchParameter Installed { get; set; }

    /// <summary>Name used by <c>VisioStencil -Catalog</c> in the DSL.</summary>
    [Parameter]
    [Alias("CatalogName", "Key")]
    public string? Name { get; set; }

    /// <summary>Make this catalog the default for later <c>VisioStencil</c> calls.</summary>
    [Parameter]
    public SwitchParameter Default { get; set; }

    /// <summary>Search directories recursively.</summary>
    [Parameter(ParameterSetName = PathParameterSet)]
    public SwitchParameter Recurse { get; set; }

    /// <summary>Catalog display name when loading package metadata.</summary>
    [Parameter(ParameterSetName = PathParameterSet)]
    [Parameter(ParameterSetName = InstalledParameterSet)]
    public string? LoadCatalogName { get; set; }

    /// <summary>Category assigned to package-backed stencil shapes.</summary>
    [Parameter(ParameterSetName = PathParameterSet)]
    [Parameter(ParameterSetName = InstalledParameterSet)]
    public string? Category { get; set; }

    /// <summary>Stable id prefix for package-backed stencil shapes.</summary>
    [Parameter(ParameterSetName = PathParameterSet)]
    [Parameter(ParameterSetName = InstalledParameterSet)]
    public string? IdPrefix { get; set; }

    /// <summary>Optional master filters for package-backed catalogs.</summary>
    [Parameter(ParameterSetName = PathParameterSet)]
    [Parameter(ParameterSetName = InstalledParameterSet)]
    public string[]? MasterName { get; set; }

    /// <summary>Include unsupported package masters as generic generated masters.</summary>
    [Parameter(ParameterSetName = PathParameterSet)]
    [Parameter(ParameterSetName = InstalledParameterSet)]
    public SwitchParameter IncludeUnsupportedMasters { get; set; }

    /// <summary>Skip reading master dimensions from package master parts.</summary>
    [Parameter(ParameterSetName = PathParameterSet)]
    [Parameter(ParameterSetName = InstalledParameterSet)]
    public SwitchParameter NoLearnMasterDimensions { get; set; }

    /// <summary>Skip reading preview image relationship metadata from package master parts.</summary>
    [Parameter(ParameterSetName = PathParameterSet)]
    [Parameter(ParameterSetName = InstalledParameterSet)]
    public SwitchParameter NoPreviewImageMetadata { get; set; }

    /// <summary>Skip reading source connection point metadata from package master parts.</summary>
    [Parameter(ParameterSetName = PathParameterSet)]
    [Parameter(ParameterSetName = InstalledParameterSet)]
    public SwitchParameter NoConnectionPointMetadata { get; set; }

    /// <summary>Default width for package-backed stencils when dimensions cannot be learned.</summary>
    [Parameter(ParameterSetName = PathParameterSet)]
    [Parameter(ParameterSetName = InstalledParameterSet)]
    public double DefaultWidth { get; set; } = 1.8;

    /// <summary>Default height for package-backed stencils when dimensions cannot be learned.</summary>
    [Parameter(ParameterSetName = PathParameterSet)]
    [Parameter(ParameterSetName = InstalledParameterSet)]
    public double DefaultHeight { get; set; } = 0.9;

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var context = VisioDslContext.Require(this);
        var catalog = ResolveCatalog();
        context.RegisterStencilCatalog(Name, catalog, Default.IsPresent);
        WriteObject(catalog);
    }

    private VisioStencilCatalog ResolveCatalog()
    {
        if (ParameterSetName == CatalogParameterSet)
        {
            return Catalog!;
        }

        if (ParameterSetName == BuiltInParameterSet)
        {
            return VisioStencilCommandUtilities.GetBuiltInCatalog(BuiltIn);
        }

        var options = VisioStencilCommandUtilities.BuildPackageLoadOptions(
            LoadCatalogName,
            Category,
            IdPrefix,
            MasterName,
            IncludeUnsupportedMasters.IsPresent,
            NoLearnMasterDimensions.IsPresent,
            NoPreviewImageMetadata.IsPresent,
            NoConnectionPointMetadata.IsPresent,
            DefaultWidth,
            DefaultHeight);

        if (ParameterSetName == InstalledParameterSet)
        {
            return VisioStencilPackageCatalog.LoadMany(VisioStencilPackageCatalog.DiscoverInstalledVisioPackages(), options);
        }

        return VisioStencilCommandUtilities.LoadCatalog(this, Path, options, Recurse.IsPresent);
    }
}
