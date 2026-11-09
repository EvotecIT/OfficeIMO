using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;
using OfficeIMO.Visio.Stencils;
using PSWriteOffice.Services.Visio;
namespace PSWriteOffice.Cmdlets.Visio;

/// <summary>Loads a stencil catalog (built-in, from .vssx/.vsdx packages or installed Visio); inside a diagram block it also registers it.</summary>
/// <example><summary>Use the native OfficeVisioStencilCatalog command.</summary><code>$network = Import-OfficeVisioStencilCatalog -BuiltIn Network</code></example>
[Cmdlet("Import", "OfficeVisioStencilCatalog", DefaultParameterSetName = "BuiltIn")]
[OutputType(typeof(VisioStencilCatalog))]
[Alias("VisioStencilCatalog") ]
public sealed class ImportOfficeVisioStencilCatalogCommand : OfficeCmdlet {
    /// <summary>Default All.</summary>
    [Parameter(ParameterSetName = "BuiltIn")]
    public OfficeVisioBuiltInStencilCatalog? BuiltIn { get; set; }
    /// <summary>Inside a diagram block: key under which the catalog is registered for VisioStencil -Catalog.</summary>
    [Parameter(ParameterSetName = "BuiltIn")]
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Installed")]
    [Parameter(ParameterSetName = "Catalog")]
    public string? Name { get; set; }
    /// <summary>Inside a diagram block: make it the default catalog.</summary>
    [Parameter(ParameterSetName = "BuiltIn")]
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Installed")]
    [Parameter(ParameterSetName = "Catalog")]
    public SwitchParameter Default { get; set; }
    /// <summary>Path for this operation.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true, Position = 0)]
    public string[]? Path { get; set; }
    /// <summary>Recurse for this operation.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    public SwitchParameter Recurse { get; set; }
    /// <summary>Name given to the loaded catalog.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Installed")]
    public string? CatalogName { get; set; }
    /// <summary>Category for this operation.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Installed")]
    public string? Category { get; set; }
    /// <summary>IdPrefix for this operation.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Installed")]
    public string? IdPrefix { get; set; }
    /// <summary>MasterName for this operation.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Installed")]
    public string[]? MasterName { get; set; }
    /// <summary>IncludeUnsupportedMasters for this operation.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Installed")]
    public SwitchParameter IncludeUnsupportedMasters { get; set; }
    /// <summary>NoLearnMasterDimensions for this operation.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Installed")]
    public SwitchParameter NoLearnMasterDimensions { get; set; }
    /// <summary>NoPreviewImageMetadata for this operation.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Installed")]
    public SwitchParameter NoPreviewImageMetadata { get; set; }
    /// <summary>NoConnectionPointMetadata for this operation.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Installed")]
    public SwitchParameter NoConnectionPointMetadata { get; set; }
    /// <summary>Default 1.8in.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Installed")]
    [OfficeLengthTransformation]
    public OfficeLength? DefaultWidth { get; set; }
    /// <summary>Default 0.9in.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Installed")]
    [OfficeLengthTransformation]
    public OfficeLength? DefaultHeight { get; set; }
    /// <summary>LiteralPath for this operation.</summary>
    [Parameter(ParameterSetName = "LiteralPath", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Alias("PSPath", "FullName")]
    public string[]? LiteralPath { get; set; }
    /// <summary>Installed for this operation.</summary>
    [Parameter(ParameterSetName = "Installed", Mandatory = true)]
    public SwitchParameter Installed { get; set; }
    /// <summary>Catalog for this operation.</summary>
    [Parameter(ParameterSetName = "Catalog", Mandatory = true, ValueFromPipeline = true)]
    public VisioStencilCatalog? Catalog { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var catalog = Catalog;
        if (catalog == null && ParameterSetName == "BuiltIn") catalog = VisioStencilCommandUtilities.GetBuiltInCatalog(BuiltIn ?? OfficeVisioBuiltInStencilCatalog.All);
        if (catalog == null) {
            var options = VisioStencilCommandUtilities.BuildPackageLoadOptions(CatalogName, Category, IdPrefix, MasterName,
                IncludeUnsupportedMasters, NoLearnMasterDimensions, NoPreviewImageMetadata, NoConnectionPointMetadata,
                (DefaultWidth ?? OfficeLength.Inches(1.8)).ToInches(), (DefaultHeight ?? OfficeLength.Inches(0.9)).ToInches());
            if (Installed) catalog = VisioStencilPackageCatalog.LoadMany(VisioStencilPackageCatalog.DiscoverInstalledVisioPackages(), options);
            else {
                var paths = new List<string>();
                foreach (var path in LiteralPath ?? Path!) paths.AddRange(OfficePath.ResolveInput(SessionState, path, LiteralPath != null, allowDirectory: true));
                catalog = VisioStencilCommandUtilities.LoadCatalog(this, paths.ToArray(), options, Recurse);
            }
        }
        VisioDslContext.Current?.RegisterStencilCatalog(Name, catalog, Default);
        WriteObject(catalog);
    }
}
