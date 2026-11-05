using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Creates native PDF annotation-flattening and page-normalization settings.</summary>
[Cmdlet(VerbsCommon.New, "OfficePdfMergeOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(PdfMergeOptions))]
public sealed class NewOfficePdfMergeOptionsCommand : OfficeCmdlet {
    /// <summary>Paint supported visual annotations into page content before merging.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter FlattenVisualAnnotations { get; set; }
    /// <summary>Target paper preset or physical dimensions.</summary>
    [Parameter, OfficePageSizeTransformation(typeof(OfficePageSize))]
    [ArgumentCompleter(typeof(OfficePdfPageSizeArgumentCompleter))]
    public OfficePageSize? PageSize { get; set; }
    /// <summary>Custom target page width.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? Width { get; set; }
    /// <summary>Custom target page height.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? Height { get; set; }
    /// <summary>Target page orientation.</summary>
    [Parameter]
    public OfficePageOrientation? Orientation { get; set; }
    /// <summary>Fit behavior for source page content.</summary>
    [Parameter]
    public PdfPageResizeMode? ResizeMode { get; set; }
    /// <summary>Margin reserved around fitted content.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? ResizeMargin { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new PdfMergeOptions();
        if (BoundSwitch(nameof(FlattenVisualAnnotations)) is bool flatten) options.FlattenVisualAnnotations = flatten;
        if (PageSize.HasValue || Width.HasValue || Height.HasValue || Orientation.HasValue || ResizeMode.HasValue || ResizeMargin.HasValue) {
            if (Width.HasValue != Height.HasValue && !PageSize.HasValue)
                throw new PSArgumentException("Supply both Width and Height, or PageSize for the other dimension.");
            var size = new PdfPageSize(Width ?? PageSize?.Width ?? PdfPageSizes.Letter.Width,
                Height ?? PageSize?.Height ?? PdfPageSizes.Letter.Height);
            if (Orientation.HasValue) size = size.WithOrientation(Orientation.Value);
            options.ResizePages = new PdfPageResizeOptions(size) {
                Mode = ResizeMode ?? PdfPageResizeMode.Fit, Margin = ResizeMargin ?? OfficeLength.Zero
            };
        }
        WriteObject(options);
    }
}
