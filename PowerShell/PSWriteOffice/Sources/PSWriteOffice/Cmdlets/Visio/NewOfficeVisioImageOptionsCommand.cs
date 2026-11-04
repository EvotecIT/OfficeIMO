using PSWriteOffice.Foundation;
using System.Management.Automation;
using OfficeIMO.Visio;
using PSWriteOffice.Cmdlets.Imaging;

namespace PSWriteOffice.Cmdlets.Visio;

/// <summary>Creates discoverable page and rendering settings for Export-OfficeVisioImage.</summary>
/// <example><summary>Render a Visio page with a selected font.</summary>
/// <code>$options = New-OfficeVisioImageOptions -Font Arial -NoConnectorLabels
/// Export-OfficeDocumentImage -Path ./Diagram.vsdx -Page 1 -Format Svg -OutputDirectory ./Preview -Options $options</code></example>
[Cmdlet(VerbsCommon.New, "OfficeVisioImageOptions", DefaultParameterSetName = "Options")]
[OutputType(typeof(VisioImageExportOptions))]
public sealed class NewOfficeVisioImageOptionsCommand : OfficeImageOptionsCommandBase<VisioImageExportOptions> {
    /// <summary>Render page text.</summary>
    [Parameter(ParameterSetName = "Options")] public SwitchParameter NoText { get; set; }
    /// <summary>Render supported stencil artwork.</summary>
    [Parameter] public SwitchParameter NoStencilArtwork { get; set; }
    /// <summary>Render connector labels.</summary>
    [Parameter] public SwitchParameter NoConnectorLabels { get; set; }
    /// <summary>Resolve connector-label overlaps.</summary>
    [Parameter] public SwitchParameter NoConnectorLabelOverlapResolution { get; set; }
    /// <summary>Raster supersampling factor.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int? Supersampling { get; set; }
    /// <summary>Include an XML declaration in SVG output.</summary>
    [Parameter] public SwitchParameter IncludeSvgXmlDeclaration { get; set; }

    /// <summary>Literal font file used for text rendering.</summary>
    [Parameter] public string? FontPath { get; set; }
    /// <summary>Font face name used for text rendering.</summary>
    [Parameter] public string? Font { get; set; }
    /// <summary>Face index within a font collection.</summary>
    [Parameter, ValidateRange(0, int.MaxValue)] public int? FontCollectionIndex { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new VisioImageExportOptions();
        ApplyCommon(options);
        Apply(nameof(NoText), value => options.RenderText = !value);
        Apply(nameof(NoStencilArtwork), value => options.RenderStencilArtwork = !value);
        Apply(nameof(NoConnectorLabels), value => options.RenderConnectorLabels = !value);
        Apply(nameof(NoConnectorLabelOverlapResolution), value => options.ResolveConnectorLabelOverlaps = !value);
        if (Supersampling.HasValue) options.Supersampling = Supersampling.Value;
        Apply(nameof(IncludeSvgXmlDeclaration), value => options.IncludeSvgXmlDeclaration = value);
        if (FontPath != null) options.FontFilePath = OfficePath.ResolveInput(SessionState, FontPath, literal: true)[0];
        if (Font != null) options.FontFaceName = Font;
        if (FontCollectionIndex.HasValue) options.FontCollectionIndex = FontCollectionIndex;
        WriteObject(options);
    }
    private void Apply(string name, System.Action<bool> setter) { if (IsBound(name)) setter(((SwitchParameter)MyInvocation.BoundParameters[name]).IsPresent); }
}
