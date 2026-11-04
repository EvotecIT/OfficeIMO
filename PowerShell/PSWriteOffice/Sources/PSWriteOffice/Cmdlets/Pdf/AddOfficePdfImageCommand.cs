using System;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.ChartForgeX;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;
using PSWriteOffice.Services.Visuals;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Adds a raster image or reusable native visual to PDF content.</summary>
[Cmdlet(VerbsCommon.Add, "OfficePdfImage", DefaultParameterSetName = "Path")]
[Alias("PdfImage")]
[OutputType(typeof(PdfDocument))]
public sealed class AddOfficePdfImageCommand : OfficeVisualCommandBase {
    /// <summary>Path for this content.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "DocumentPath", Mandatory = true, Position = 0)]
    public string? Path { get; set; }
    /// <summary>Alignment for this content.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "InputObject")]
    [Parameter(ParameterSetName = "DocumentPath")]
    [Parameter(ParameterSetName = "DocumentInputObject")]
    public OfficeAlignment? Alignment { get; set; }
    /// <summary>AlternativeText for this content.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "InputObject")]
    [Parameter(ParameterSetName = "DocumentPath")]
    [Parameter(ParameterSetName = "DocumentInputObject")]
    public string? AlternativeText { get; set; }
    /// <summary>Link for this content.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "InputObject")]
    [Parameter(ParameterSetName = "DocumentPath")]
    [Parameter(ParameterSetName = "DocumentInputObject")]
    public string? Link { get; set; }
    /// <summary>LinkContents for this content.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "InputObject")]
    [Parameter(ParameterSetName = "DocumentPath")]
    [Parameter(ParameterSetName = "DocumentInputObject")]
    public string? LinkContents { get; set; }
    /// <summary>InputObject for this content.</summary>
    [Parameter(ParameterSetName = "InputObject", Mandatory = true, ValueFromPipeline = true)]
    [Parameter(ParameterSetName = "DocumentInputObject", Mandatory = true, ValueFromPipeline = true)]
    public object? InputObject { get; set; }
    /// <summary>SpacingBefore for this content.</summary>
    [Parameter(ParameterSetName = "InputObject")]
    [Parameter(ParameterSetName = "DocumentInputObject")]
    [OfficeLengthTransformation]
    public OfficeLength? SpacingBefore { get; set; }
    /// <summary>SpacingAfter for this content.</summary>
    [Parameter(ParameterSetName = "InputObject")]
    [Parameter(ParameterSetName = "DocumentInputObject")]
    [OfficeLengthTransformation]
    public OfficeLength? SpacingAfter { get; set; }
    /// <summary>Document for this content.</summary>
    [Parameter(ParameterSetName = "DocumentPath", Mandatory = true)]
    [Parameter(ParameterSetName = "DocumentInputObject", Mandatory = true)]
    public PdfDocument? Document { get; set; }
    /// <summary>PassThru for this content.</summary>
    [Parameter(ParameterSetName = "DocumentPath")]
    [Parameter(ParameterSetName = "DocumentInputObject")]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override string? VisualAlternativeText => AlternativeText;
    /// <summary>SvgPolicy for a reusable visual input.</summary>
    [Parameter(ParameterSetName = "InputObject")]
    [Parameter(ParameterSetName = "DocumentInputObject")]
    public new OfficeVisualSvgPolicy SvgPolicy { get => base.SvgPolicy; set => base.SvgPolicy = value; }
    /// <summary>PointsPerPixel for a reusable visual input.</summary>
    [Parameter(ParameterSetName = "InputObject")]
    [Parameter(ParameterSetName = "DocumentInputObject")]
    public new double PointsPerPixel { get => base.PointsPerPixel; set => base.PointsPerPixel = value; }
    /// <summary>MaximumSvgElements for a reusable visual input.</summary>
    [Parameter(ParameterSetName = "InputObject")]
    [Parameter(ParameterSetName = "DocumentInputObject")]
    public new int? MaximumSvgElements { get => base.MaximumSvgElements; set => base.MaximumSvgElements = value; }
    /// <summary>MaximumSvgViewportDimension for a reusable visual input.</summary>
    [Parameter(ParameterSetName = "InputObject")]
    [Parameter(ParameterSetName = "DocumentInputObject")]
    public new double? MaximumSvgViewportDimension { get => base.MaximumSvgViewportDimension; set => base.MaximumSvgViewportDimension = value; }
    /// <summary>MaximumSvgViewportPixelCount for a reusable visual input.</summary>
    [Parameter(ParameterSetName = "InputObject")]
    [Parameter(ParameterSetName = "DocumentInputObject")]
    public new long? MaximumSvgViewportPixelCount { get => base.MaximumSvgViewportPixelCount; set => base.MaximumSvgViewportPixelCount = value; }
    /// <summary>Id for a reusable visual input.</summary>
    [Parameter(ParameterSetName = "InputObject")]
    [Parameter(ParameterSetName = "DocumentInputObject")]
    public new string? Id { get => base.Id; set => base.Id = value; }
    /// <summary>Title for a reusable visual input.</summary>
    [Parameter(ParameterSetName = "InputObject")]
    [Parameter(ParameterSetName = "DocumentInputObject")]
    public new string? Title { get => base.Title; set => base.Title = value; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        var cell = Document == null ? OfficeDocumentContext.Find<PdfCellContent>() : null;
        if (InputObject != null) {
            var visual = ResolveVisual(InputObject, nativePlacement: true);
            if (Width.HasValue || Height.HasValue) {
                double width = Width?.ToPoints() ?? visual.WidthPoints * Height!.Value.ToPoints() / visual.HeightPoints;
                visual = visual.WithSize(width, Height?.ToPoints(), OfficeImageFit.Stretch);
            }
            if (cell != null) {
                if (visual.PlacementFormat == OfficeVisualMediaFormat.Svg)
                    throw new PSNotSupportedException("Native PDF table-cell images require a raster placement; this vector visual needs a native drawing cell.");
                cell.Images.Add(new PdfTableCellImage(visual.GetPlacementBytes(), OfficeLength.Points(visual.WidthPoints), OfficeLength.Points(visual.HeightPoints),
                    new PdfImageStyle { AltText = AlternativeText ?? visual.AlternativeText }, Link, LinkContents));
                return;
            }
            var target = ParameterSetName == "DocumentInputObject" ? "DocumentInputObject" : "DocumentPath";
            var result = PdfCommandUtilities.ComposeContent(this, Document, ParameterSetName, target,
                content => content.AddVisualArtifact(visual, Alignment, SpacingBefore?.ToPoints(), SpacingAfter?.ToPoints(),
                    AlternativeText == null ? null : new PdfDrawingStyle { AlternativeText = AlternativeText }, Link, LinkContents));
            if (PassThru && result != null) WriteObject(result);
            return;
        }
        string path = OfficePath.ResolveInput(SessionState, Path!, literal: true)[0];
        var bytes = PdfImageDocumentSource.FromFile(path).GetBytes();
        var info = OfficeImageReader.Identify(bytes, path);
        if (info.Width <= 0 || info.Height <= 0) throw new PSArgumentException("The image has no usable intrinsic dimensions.");
        var widthLength = Width ?? (Height.HasValue ? OfficeLength.Points(Height.Value.ToPoints() * info.Width / info.Height) : OfficeLength.Inches(info.Width / info.DpiX));
        var heightLength = Height ?? OfficeLength.Points(widthLength.ToPoints() * info.Height / info.Width);
        if (cell != null) {
            cell.Images.Add(new PdfTableCellImage(bytes, widthLength, heightLength, new PdfImageStyle { AltText = AlternativeText }, Link, LinkContents));
            return;
        }
        var document = PdfCommandUtilities.ComposeContent(this, Document, ParameterSetName, "DocumentPath",
            content => content.AddImage(bytes, widthLength, heightLength, Alignment, linkUri: Link, linkContents: LinkContents, altText: AlternativeText));
        if (PassThru && document != null) WriteObject(document);
    }
}
