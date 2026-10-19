using System;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;
using PSWriteOffice.Services.Text;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Adds PowerShell-friendly text or rich text runs to the active fixed-position PDF canvas.</summary>
/// <remarks>
/// Use this command inside <c>Add-OfficePdfCanvas -Content</c>. Coordinates are PDF points measured
/// from the visual top-left of the page. Rich runs accept <c>TextRun</c> output, hashtables, and objects;
/// callers do not need to construct native <see cref="PdfTextRun"/> arrays. Fixed-position canvas
/// runs are visual text and do not support link targets. Width and height default to the remaining
/// page area from the supplied coordinates.
/// </remarks>
/// <example>
///   <summary>Place mixed-format text on each selected page.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficePdfCanvas -Path .\Report.pdf -OutputPath .\Stamped.pdf -Content {
///     PdfCanvasText -Run @(
///       TextRun 'Owner: ' -Bold
///       TextRun 'Platform' -Color '#0F766E'
///     ) -X 36 -Y 24 -FontSize 10
/// }</code>
///   <para>The enclosing callback supplies the active page, while the run collection remains an ordinary PowerShell array.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePdfCanvasText", DefaultParameterSetName = ParameterSetText)]
[Alias("PdfCanvasText")]
public sealed class AddOfficePdfCanvasTextCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    private const string ParameterSetText = "Text";
    private const string ParameterSetRun = "Run";

    /// <summary>Plain text values to concatenate into one positioned text block.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetText)]
    public string[] Text { get; set; } = Array.Empty<string>();

    /// <summary>Rich run specifications created with TextRun or supplied as hashtables or objects. Link targets are not supported.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetRun)]
    [Alias("Runs")]
    public object[] Run { get; set; } = Array.Empty<object>();

    /// <summary>Horizontal position in PDF points from the visual left edge.</summary>
    [Parameter(Mandatory = true)]
    [Alias("Left")]
    public double X { get; set; }

    /// <summary>Vertical position in PDF points from the visual top edge.</summary>
    [Parameter(Mandatory = true)]
    [Alias("Top")]
    public double Y { get; set; }

    /// <summary>Available text width in PDF points. Defaults to the remaining page width.</summary>
    [Parameter]
    public double? Width { get; set; }

    /// <summary>Available text height in PDF points. Defaults to the remaining page height.</summary>
    [Parameter]
    public double? Height { get; set; }

    /// <summary>Default text color. Named and hexadecimal colors are accepted.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? Color { get; set; }

    /// <summary>Text alignment within the positioned rectangle.</summary>
    [Parameter]
    public PdfAlign Align { get; set; } = PdfAlign.Left;

    /// <summary>Default font size in PDF points.</summary>
    [Parameter]
    public double? FontSize { get; set; }

    /// <summary>Optional line height in PDF points.</summary>
    [Parameter]
    public double? LineHeight { get; set; }

    /// <summary>Make plain -Text input bold.</summary>
    [Parameter(ParameterSetName = ParameterSetText)]
    public SwitchParameter Bold { get; set; }

    /// <summary>Make plain -Text input italic.</summary>
    [Parameter(ParameterSetName = ParameterSetText)]
    public SwitchParameter Italic { get; set; }

    /// <summary>Underline plain -Text input.</summary>
    [Parameter(ParameterSetName = ParameterSetText)]
    public SwitchParameter Underline { get; set; }

    /// <summary>Strike through plain -Text input.</summary>
    [Parameter(ParameterSetName = ParameterSetText)]
    public SwitchParameter Strike { get; set; }

    /// <summary>Background color for plain -Text input.</summary>
    [Parameter(ParameterSetName = ParameterSetText)]
    [OfficeColorArgumentTransformation]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? BackgroundColor { get; set; }

    /// <summary>Standard PDF font for plain -Text input.</summary>
    [Parameter(ParameterSetName = ParameterSetText)]
    public PdfStandardFont? Font { get; set; }

    /// <summary>Baseline for plain -Text input.</summary>
    [Parameter(ParameterSetName = ParameterSetText)]
    public PdfTextBaseline Baseline { get; set; } = PdfTextBaseline.Normal;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var context = PdfCanvasDslContext.Require(this);
        var width = Width ?? context.Page.Width - X;
        var height = Height ?? context.Page.Height - Y;
        if (width <= 0 || height <= 0) {
            throw new PSArgumentException(
                "The text area must remain inside the page. Adjust -X/-Y or provide positive -Width/-Height values.");
        }

        var runs = ParameterSetName == ParameterSetRun
            ? PdfRichTextRunBuilder.ToCanvasTextRuns(Run)
            : CreatePlainTextRuns();

        context.Canvas.Text(
            runs,
            X,
            Y,
            width,
            height,
            PdfCommandUtilities.ParseColor(Color),
            Align,
            FontSize,
            LineHeight);
        WritePassThru(context.Page);
    }

    private PdfTextRun[] CreatePlainTextRuns() {
        if (Text.Length == 0) {
            throw new PSArgumentException("Provide at least one text value.");
        }

        return PdfRichTextRunBuilder.ToCanvasTextRuns(new[]
        {
            new OfficeTextRunSpec
            {
                Text = string.Concat(Text),
                Bold = Bold.IsPresent,
                Italic = Italic.IsPresent,
                Underline = Underline.IsPresent,
                Strike = Strike.IsPresent,
                BackgroundColor = BackgroundColor,
                FontName = Font?.ToString(),
                Baseline = Baseline.ToString()
            }
        });
    }
}