using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Rotates pages, sets boundary boxes or resizes pages through the native PDF engine.</summary>
[Cmdlet(VerbsCommon.Set, "OfficePdfPage", DefaultParameterSetName = "Document", SupportsShouldProcess = true)]
[OutputType(typeof(PdfDocument), typeof(FileInfo))]
public sealed class SetOfficePdfPageCommand : PdfMutationCmdlet {
    private bool _resize;
    private bool _orientationBound;
    private PdfPageResizeOptions? _resizeOptions;
    /// <summary>One-based pages; omission selects every page.</summary>
    [Parameter, OfficePageRangeTransformation]
    public OfficePageRange? Page { get; set; }
    /// <summary>Absolute rotation in degrees.</summary>
    [Parameter, ValidateSet("0", "90", "180", "270")]
    public int Rotation { get; set; }
    /// <summary>Page boundary box to set.</summary>
    [Parameter]
    public PdfPageBoundaryBox? BoxName { get; set; }
    /// <summary>Left boundary coordinate.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? BoxLeft { get; set; }
    /// <summary>Bottom boundary coordinate.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? BoxBottom { get; set; }
    /// <summary>Right boundary coordinate.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? BoxRight { get; set; }
    /// <summary>Top boundary coordinate.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? BoxTop { get; set; }
    /// <summary>Paper preset or physical dimensions for resizing.</summary>
    [Parameter, OfficePageSizeTransformation(typeof(OfficePageSize))]
    [ArgumentCompleter(typeof(OfficePdfPageSizeArgumentCompleter))]
    public OfficePageSize? PageSize { get; set; }
    /// <summary>Custom page width.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? Width { get; set; }
    /// <summary>Custom page height.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? Height { get; set; }
    /// <summary>Orientation of the target page.</summary>
    [Parameter]
    public OfficePageOrientation Orientation { get; set; }
    /// <summary>How original page content fits the resized page.</summary>
    [Parameter]
    public PdfPageResizeMode ResizeMode { get; set; } = PdfPageResizeMode.Fit;
    /// <summary>Margin reserved around fitted content.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? ResizeMargin { get; set; }
    /// <inheritdoc />
    protected override string Operation => "Set PDF page properties";
    /// <inheritdoc />
    protected override void Prepare() {
        _orientationBound = MyInvocation.BoundParameters.ContainsKey(nameof(Orientation));
        _resize = PageSize.HasValue || Width.HasValue || Height.HasValue || _orientationBound ||
            ResizeMargin.HasValue || MyInvocation.BoundParameters.ContainsKey(nameof(ResizeMode));
        bool rotate = MyInvocation.BoundParameters.ContainsKey(nameof(Rotation));
        bool boxCoordinates = BoxLeft.HasValue || BoxBottom.HasValue || BoxRight.HasValue || BoxTop.HasValue;
        if ((_resize ? 1 : 0) + (BoxName.HasValue ? 1 : 0) + (rotate ? 1 : 0) != 1)
            throw new PSArgumentException("Choose one operation: resize, rotation, or a boundary box with all four coordinates.");
        if (boxCoordinates && !BoxName.HasValue)
            throw new PSArgumentException("Boundary coordinates require -BoxName.");
        if (BoxName.HasValue && (!BoxLeft.HasValue || !BoxBottom.HasValue || !BoxRight.HasValue || !BoxTop.HasValue))
            throw new PSArgumentException("-BoxName requires all four boundary coordinates.");
        if (_resize) {
            if (Width.HasValue != Height.HasValue && !PageSize.HasValue)
                throw new PSArgumentException("Supply both Width and Height, or a PageSize to provide the other dimension.");
            var size = new PdfPageSize(Width ?? PageSize?.Width ?? PdfPageSizes.Letter.Width,
                Height ?? PageSize?.Height ?? PdfPageSizes.Letter.Height);
            if (_orientationBound) size = size.WithOrientation(Orientation);
            _resizeOptions = new PdfPageResizeOptions(size) { Mode = ResizeMode, Margin = ResizeMargin ?? OfficeLength.Zero };
        }
    }
    /// <inheritdoc />
    protected override PdfDocument Apply(PdfDocument document, CancellationToken cancellationToken) {
        if (_resize) return document.Pages.Resize(_resizeOptions!, Page?.Selection, cancellationToken);
        if (BoxName.HasValue) return document.Pages.SetPageBox(BoxName.Value, BoxLeft!.Value.ToPoints(),
            BoxBottom!.Value.ToPoints(), BoxRight!.Value.ToPoints(), BoxTop!.Value.ToPoints(), Page?.Selection, cancellationToken);
        return Page != null ? document.Pages.Rotate(Rotation, Page.Selection, cancellationToken)
            : document.Pages.Rotate(Rotation, PdfPageSelector.Parse("all"), cancellationToken);
    }
}
