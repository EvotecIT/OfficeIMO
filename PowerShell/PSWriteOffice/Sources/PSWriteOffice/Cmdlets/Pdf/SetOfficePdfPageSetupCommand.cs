using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Sets physical dimensions, orientation and margins in the enclosing PDF composition.</summary>
[Cmdlet(VerbsCommon.Set, "OfficePdfPageSetup", DefaultParameterSetName = "Context")]
[Alias("PdfPageSetup")]
public sealed class SetOfficePdfPageSetupCommand : PSCmdlet {
    /// <summary>Named paper size or physical page dimensions.</summary>
    [Parameter(ParameterSetName = "Context"), OfficePageSizeTransformation(typeof(OfficePageSize))]
    [ArgumentCompleter(typeof(OfficePdfPageSizeArgumentCompleter))]
    public OfficePageSize? PageSize { get; set; }
    /// <summary>Custom page width.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? Width { get; set; }
    /// <summary>Custom page height.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? Height { get; set; }
    /// <summary>Page orientation.</summary>
    [Parameter]
    public OfficePageOrientation Orientation { get; set; }
    /// <summary>Uniform margin.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? Margin { get; set; }
    /// <summary>Left margin.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? MarginLeft { get; set; }
    /// <summary>Top margin.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? MarginTop { get; set; }
    /// <summary>Right margin.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? MarginRight { get; set; }
    /// <summary>Bottom margin.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? MarginBottom { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        var orientationBound = MyInvocation.BoundParameters.ContainsKey(nameof(Orientation));
        PdfDslContext.Require(this).ConfigurePage(page => {
            if (PageSize.HasValue) page.Size(PageSize.Value.Width, PageSize.Value.Height);
            if (Width.HasValue || Height.HasValue)
                page.Size(Width ?? page.PageSetup.Size.Width, Height ?? page.PageSetup.Size.Height);
            if (orientationBound) page.Orientation(Orientation);
            if (Margin.HasValue) page.Margin(Margin.Value);
            if (MarginLeft.HasValue) page.PageSetup.Margins.Left = MarginLeft.Value;
            if (MarginTop.HasValue) page.PageSetup.Margins.Top = MarginTop.Value;
            if (MarginRight.HasValue) page.PageSetup.Margins.Right = MarginRight.Value;
            if (MarginBottom.HasValue) page.PageSetup.Margins.Bottom = MarginBottom.Value;
        });
    }
}
