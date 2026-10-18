using System.IO;
using System.Management.Automation;
using OfficeIMO.Visio.Stencils;
using PSWriteOffice.Services.Visio;

namespace PSWriteOffice.Cmdlets.Visio;

/// <summary>Gets built-in or package-backed OfficeIMO Visio stencil catalogs.</summary>
/// <example>
///   <summary>Load and inspect the flowchart stencil catalog.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$catalog = Get-OfficeVisioStencilCatalog -BuiltIn Flowchart
/// Find-OfficeVisioStencil -Catalog $catalog -Query decision -First 3</code>
///   <para>Gets a built-in catalog and searches it before using stencils in a diagram.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeVisioStencilCatalog", DefaultParameterSetName = BuiltInParameterSet)]
[Alias("VisioStencilCatalog")]
[OutputType(typeof(VisioStencilCatalog))]
public sealed class GetOfficeVisioStencilCatalogCommand : PSCmdlet
{
    private const string BuiltInParameterSet = "BuiltIn";
    private const string PathParameterSet = "Path";
    private const string InstalledParameterSet = "Installed";

    /// <summary>Built-in OfficeIMO stencil catalog to return.</summary>
    [Parameter(ParameterSetName = BuiltInParameterSet)]
    public OfficeVisioBuiltInStencilCatalog BuiltIn { get; set; } = OfficeVisioBuiltInStencilCatalog.All;

    /// <summary>Visio package, package directory, or OfficeIMO native stencil manifest path.</summary>
    [Parameter(Mandatory = true, ParameterSetName = PathParameterSet, Position = 0)]
    [Alias("FilePath", "LiteralPath")]
    public string[] Path { get; set; } = [];

    /// <summary>Discover installed Microsoft Visio stencils and templates.</summary>
    [Parameter(Mandatory = true, ParameterSetName = InstalledParameterSet)]
    public SwitchParameter Installed { get; set; }

    /// <summary>Search directories recursively.</summary>
    [Parameter(ParameterSetName = PathParameterSet)]
    public SwitchParameter Recurse { get; set; }

    /// <summary>Catalog display name when loading package metadata.</summary>
    [Parameter(ParameterSetName = PathParameterSet)]
    [Parameter(ParameterSetName = InstalledParameterSet)]
    public string? CatalogName { get; set; }

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
        if (ParameterSetName == BuiltInParameterSet)
        {
            WriteObject(VisioStencilCommandUtilities.GetBuiltInCatalog(BuiltIn));
            return;
        }

        var options = VisioStencilCommandUtilities.BuildPackageLoadOptions(
            CatalogName,
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
            var installedPackages = VisioStencilPackageCatalog.DiscoverInstalledVisioPackages();
            WriteObject(VisioStencilPackageCatalog.LoadMany(installedPackages, options));
            return;
        }

        WriteObject(VisioStencilCommandUtilities.LoadCatalog(this, Path, options, Recurse.IsPresent));
    }
}
