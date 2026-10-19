using System.Linq;
using System.Management.Automation;
using OfficeIMO.Visio.Stencils;
using PSWriteOffice.Services.Visio;

namespace PSWriteOffice.Cmdlets.Visio;

/// <summary>Searches OfficeIMO Visio stencil catalogs.</summary>
/// <example>
///   <summary>Search the built-in flowchart catalog.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$flow = Get-OfficeVisioStencilCatalog -BuiltIn Flowchart
/// Find-OfficeVisioStencil -Catalog $flow -Query process -First 5</code>
///   <para>Returns matching stencil definitions that can be used by VisioStencil.</para>
/// </example>
[Cmdlet(VerbsCommon.Find, "OfficeVisioStencil", DefaultParameterSetName = QueryParameterSet)]
[Alias("Find-VisioStencil")]
[OutputType(typeof(VisioStencilShape))]
public sealed class FindOfficeVisioStencilCommand : PSCmdlet
{
    private const string QueryParameterSet = "Query";
    private const string BuiltInParameterSet = "BuiltIn";
    private const string CatalogNameParameterSet = "CatalogName";

    /// <summary>Catalog object to search. Defaults to the combined built-in catalog.</summary>
    [Parameter(ValueFromPipeline = true)]
    public VisioStencilCatalog? Catalog { get; set; }

    /// <summary>Catalog previously registered in the active Visio DSL scope.</summary>
    [Parameter(ParameterSetName = CatalogNameParameterSet)]
    public string? CatalogName { get; set; }

    /// <summary>Built-in OfficeIMO stencil catalog to search.</summary>
    [Parameter(ParameterSetName = BuiltInParameterSet)]
    public OfficeVisioBuiltInStencilCatalog BuiltIn { get; set; } = OfficeVisioBuiltInStencilCatalog.All;

    /// <summary>Search text. Empty search returns catalog contents.</summary>
    [Parameter(Position = 0)]
    public string? Query { get; set; }

    /// <summary>Maximum number of shapes to return.</summary>
    [Parameter]
    public int First { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var catalog = ResolveCatalog();
        var matches = catalog.Search(Query ?? string.Empty);
        if (First > 0)
        {
            matches = matches.Take(First).ToList().AsReadOnly();
        }

        WriteObject(matches, enumerateCollection: true);
    }

    private VisioStencilCatalog ResolveCatalog()
    {
        if (Catalog != null)
        {
            return Catalog;
        }

        if (ParameterSetName == BuiltInParameterSet)
        {
            return VisioStencilCommandUtilities.GetBuiltInCatalog(BuiltIn);
        }

        if (ParameterSetName == CatalogNameParameterSet)
        {
            return VisioDslContext.Require(this).ResolveStencilCatalog(CatalogName);
        }

        return VisioDslContext.Current?.DefaultStencilCatalog ?? VisioStencils.All;
    }
}
