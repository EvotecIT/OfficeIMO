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
[Cmdlet(VerbsCommon.Find, "OfficeVisioStencil", DefaultParameterSetName = "Catalog")]
[OutputType(typeof(VisioStencilShape))]
public sealed class FindOfficeVisioStencilCommand : PSCmdlet
{
    private const string BuiltInParameterSet = "BuiltIn";

    /// <summary>A stencil catalog object or a name registered in the current Visio block.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = "Catalog")]
    public object? Catalog { get; set; }

    /// <summary>Built-in OfficeIMO stencil catalog to search.</summary>
    [Parameter(Mandatory = true, ParameterSetName = BuiltInParameterSet)]
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
        object? value = Catalog is PSObject wrapped ? wrapped.BaseObject : Catalog;
        if (value is VisioStencilCatalog supplied) return supplied;
        if (value is string name) return VisioDslContext.Require(this).ResolveStencilCatalog(name);
        if (value != null) throw new PSArgumentException("Catalog accepts a VisioStencilCatalog or a registered name.");

        if (ParameterSetName == BuiltInParameterSet)
        {
            return VisioStencilCommandUtilities.GetBuiltInCatalog(BuiltIn);
        }

        return VisioDslContext.Current?.DefaultStencilCatalog ?? VisioStencils.All;
    }
}
