using System.Management.Automation;
using OfficeIMO.Visio;
using PSWriteOffice.Cmdlets.Imaging;

namespace PSWriteOffice.Cmdlets.Visio;

/// <summary>Creates discoverable page and rendering settings for Export-OfficeVisioImage.</summary>
/// <example>
///   <summary>Render the first Visio page with text and connector labels.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficeVisioImageOptions -PageIndex 0 -PageCount 1 -RenderText -RenderConnectorLabels
/// Export-OfficeVisioImage -Path .\Diagram.vsdx -OutputPath .\Preview -Format Svg -Options $options</code>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeVisioImageOptions")]
[OutputType(typeof(VisioImageExportOptions))]
public sealed class NewOfficeVisioImageOptionsCommand : OfficeImageOptionsCommandBase<VisioImageExportOptions> {
    /// <summary>Zero-based first page index.</summary>
    [Parameter] [ValidateRange(0, int.MaxValue)] public int? PageIndex { get; set; }
    /// <summary>Maximum pages exported.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int? PageCount { get; set; }
    /// <summary>Render page text.</summary>
    [Parameter] public SwitchParameter RenderText { get; set; }
    /// <summary>Render supported stencil artwork.</summary>
    [Parameter] public SwitchParameter RenderStencilArtwork { get; set; }
    /// <summary>Render connector labels.</summary>
    [Parameter] public SwitchParameter RenderConnectorLabels { get; set; }
    /// <summary>Resolve connector-label overlaps.</summary>
    [Parameter] public SwitchParameter ResolveConnectorLabelOverlaps { get; set; }
    /// <summary>Raster supersampling factor.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int? Supersampling { get; set; }
    /// <summary>Include an XML declaration in SVG output.</summary>
    [Parameter] public SwitchParameter IncludeSvgXmlDeclaration { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new VisioImageExportOptions();
        ApplyCommon(options);
        if (PageIndex.HasValue) options.PageIndex = PageIndex.Value;
        if (PageCount.HasValue) options.PageCount = PageCount.Value;
        Apply(nameof(RenderText), value => options.RenderText = value);
        Apply(nameof(RenderStencilArtwork), value => options.RenderStencilArtwork = value);
        Apply(nameof(RenderConnectorLabels), value => options.RenderConnectorLabels = value);
        Apply(nameof(ResolveConnectorLabelOverlaps), value => options.ResolveConnectorLabelOverlaps = value);
        if (Supersampling.HasValue) options.Supersampling = Supersampling.Value;
        Apply(nameof(IncludeSvgXmlDeclaration), value => options.IncludeSvgXmlDeclaration = value);
        WriteObject(options);
    }
    private void Apply(string name, System.Action<bool> setter) { if (IsBound(name)) setter(((SwitchParameter)MyInvocation.BoundParameters[name]).IsPresent); }
}
