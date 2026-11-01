using System;
using System.Management.Automation;
using System.Threading.Tasks;
using OfficeIMO;
using OfficeIMO.ChartForgeX;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Excel;
using PSWriteOffice.Services.Visuals;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Embeds a file, HTTP image or reusable visual at a worksheet cell or range.</summary>
/// <para>A cell uses a one-cell anchor; a range uses a two-cell anchor. Numbers mean points and unit suffixes are accepted. With no dimensions or Scale, file images retain their detected size. Range anchors determine their own size. URL downloads use the native network policy and cancellation token.</para>
/// <example><summary>Place a logo at its natural size.</summary>
/// <code>ExcelSheet 'Cover' { ExcelImage -Path ./logo.png -Address A1 -AltText 'Contoso' }</code></example>
/// <example><summary>Pin a reusable visual to the report header.</summary>
/// <code>$visual | ExcelImage -Document $book -Sheet Summary -Address B2:G12 -Name Health</code></example>
/// <example><summary>Download and resize a status badge.</summary>
/// <code>ExcelImage -Url https://status.example.com/badge.png -Address H1 -Width 4cm</code></example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelImage", DefaultParameterSetName = "Path")]
[Alias("ExcelImage")]
[OutputType(typeof(ExcelImage))]
public sealed class AddOfficeExcelImageCommand : OfficeVisualCommandBase {
    /// <summary>Literal local image path.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Path"), ValidateNotNullOrEmpty]
    public string? Path { get; set; }

    /// <summary>HTTP or HTTPS image URI; local files use Path.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Url"), ValidateNotNull]
    public Uri? Url { get; set; }

    /// <summary>ChartForgeX artifact, OfficeVisualSource, prepared visual or SVG path.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "Visual"), ValidateNotNull]
    public object? InputObject { get; set; }

    /// <summary>Workbook to modify outside the Excel DSL.</summary>
    [Parameter(ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Sheet name, one-based position or native worksheet belonging to Document.</summary>
    [Parameter, OfficeSheetTransformation]
    public OfficeSheetSelector? Sheet { get; set; }

    /// <summary>A1 cell or range; defaults to A1. Use B2:B2 for a two-cell anchor over one cell.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string Address { get; set; } = "A1";

    /// <summary>Positive percentage of the source size; exclusive with dimensions and range anchors.</summary>
    [Parameter]
    public double? Scale { get; set; }

    /// <summary>Horizontal start-marker offset; numbers mean points.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength OffsetX { get; set; }

    /// <summary>Vertical start-marker offset; numbers mean points.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength OffsetY { get; set; }

    /// <summary>Horizontal end-marker offset for a range anchor; numbers mean points.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength EndOffsetX { get; set; }

    /// <summary>Vertical end-marker offset for a range anchor; numbers mean points.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength EndOffsetY { get; set; }

    /// <summary>Drawing name in Excel's selection pane.</summary>
    [Parameter]
    public string? Name { get; set; }

    /// <summary>Accessible image description; omitted visual descriptions are retained.</summary>
    [Parameter]
    public string? AltText { get; set; }

    /// <inheritdoc />
    protected override string? VisualAlternativeText => AltText;

    /// <inheritdoc />
    protected override string AlternativeTextParameterName => nameof(AltText);

    /// <summary>Clear accessible title and description and mark the image decorative.</summary>
    [Parameter]
    public SwitchParameter Decorative { get; set; }

    /// <summary>Allow interactive aspect-ratio changes in Excel; the default locks the ratio.</summary>
    [Parameter]
    public SwitchParameter NoLockAspectRatio { get; set; }

    /// <summary>Movement and resizing behavior of a two-cell anchor.</summary>
    [Parameter]
    public ExcelImagePlacement Placement { get; set; } = ExcelImagePlacement.MoveAndSize;

    /// <summary>Clockwise rotation in degrees.</summary>
    [Parameter]
    public double Rotation { get; set; }

    /// <summary>Return the native editable ExcelImage.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override async Task ProcessRecordAsync() {
        bool range = Address.Contains(":");
        ValidatePlacement(range);
        // All PowerShell binding/context work finishes before native asynchronous I/O.
        ExcelSheet sheet = ExcelSheetResolver.ResolveTarget(this, Document, Sheet);
        string? path = Path == null ? null : OfficePath.ResolveInput(SessionState, Path, literal: true)[0];
        Uri? url = Url;
        if (url != null) ExcelImageUrlPolicy.RequireWebUrl(url.ToString());
        OfficeVisualConversionResult? visual = ParameterSetName == "Visual" ? ResolveVisual(InputObject!, nativePlacement: true) : null;
        var options = new ExcelImageOptions {
            Width = Width, Height = Height, ScalePercent = Scale,
            OffsetX = OffsetX, OffsetY = OffsetY, EndOffsetX = EndOffsetX, EndOffsetY = EndOffsetY,
            Name = Name, Title = Decorative ? null : Title ?? visual?.Title,
            Description = Decorative ? null : AltText ?? visual?.AlternativeText,
            LockAspectRatio = !NoLockAspectRatio, Placement = Placement, RotationDegrees = Rotation
        };
        ExcelImage image;
        if (visual != null) {
            options.ContentType = visual.PlacementMediaType;
            if (!range) {
                options.ScalePercent = null;
                options.Width ??= OfficeLength.Points(visual.WidthPoints * (Scale ?? 100) / 100);
                options.Height ??= OfficeLength.Points(visual.HeightPoints * (Scale ?? 100) / 100);
            }
            image = sheet.AddImage(Address, visual.GetPlacementBytes(), options);
        } else if (url != null) {
            image = await sheet.AddImageAsync(Address, url, options, CancelToken).ConfigureAwait(false);
        } else {
            image = sheet.AddImage(Address, path!, options);
        }
        if (Decorative || visual?.IsDecorative == true && AltText == null && Title == null) image.Decorative();
        if (PassThru) WriteObject(image);
    }

    private void ValidatePlacement(bool range) {
        if (Width.HasValue && Width.Value <= OfficeLength.Zero || Height.HasValue && Height.Value <= OfficeLength.Zero)
            throw new PSArgumentException("Image dimensions must be positive.");
        if (Scale.HasValue && (Scale.Value <= 0 || double.IsNaN(Scale.Value) || double.IsInfinity(Scale.Value)))
            throw new PSArgumentException("Scale must be a positive finite percentage.");
        if (Scale.HasValue && (Width.HasValue || Height.HasValue))
            throw new PSArgumentException("Scale cannot be combined with Width or Height.");
        if (range && (Scale.HasValue || Width.HasValue || Height.HasValue))
            throw new PSArgumentException("A range Address determines the image size; omit Scale, Width and Height.");
        if (Decorative && (!string.IsNullOrWhiteSpace(AltText) || !string.IsNullOrWhiteSpace(Title)))
            throw new PSArgumentException("Decorative images cannot also define AltText or Title.");
    }
}
