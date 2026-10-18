using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Adds a rich inline-text paragraph to a generated PDF document.</summary>
/// <remarks>
/// Use <c>Add-OfficePdfText</c> when a paragraph needs mixed emphasis, highlight color, font settings, baseline changes, or links.
/// Plain paragraphs can continue to use <c>Add-OfficePdfParagraph</c>. Rich text runs are translated directly to the OfficeIMO.Pdf paragraph builder.
/// URI links and bookmark links are supported; a single run cannot target both.
/// </remarks>
/// <example>
///   <summary>Add styled text with the DSL alias.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePdf -Path .\Report.pdf { PdfText 'Approved for review' -Bold -Color '#0F766E' -BackgroundColor '#ECFDF5' }</code>
///   <para>Creates a PDF with one styled paragraph.</para>
/// </example>
/// <example>
///   <summary>Add mixed rich runs with URI and bookmark links.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePdf -Path .\Report.pdf {
///     PdfBookmark 'summary'
///     PdfText -Run @(
///       @{ Text = 'Read the ' }
///       @{ Text = 'website'; LinkUri = 'https://evotec.xyz'; Color = '#2563EB' }
///       @{ Text = ' or jump to ' }
///       @{ Text = 'summary'; LinkDestinationName = 'summary'; Color = '#7C3AED' }
///       @{ Text = '.' }
///     )
///   }</code>
///   <para>Creates one paragraph with an external link and an internal named-destination link.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePdfText", DefaultParameterSetName = ParameterSetContext)]
[Alias("PdfText")]
[OutputType(typeof(PdfDocument))]
public sealed class AddOfficePdfTextCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>PDF document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Plain text values to add as one styled paragraph.</summary>
    [Parameter(Position = 0)]
    public string[]? Text { get; set; }

    /// <summary>Rich run specifications. Each run may define Text, Bold, Italic, Underline, Strike, Color, BackgroundColor, FontSize, Font, Baseline, LinkUri, LinkDestinationName, LinkContents, Type, or Kind.</summary>
    [Parameter]
    [Alias("Runs")]
    public object[]? Run { get; set; }

    /// <summary>Paragraph alignment.</summary>
    [Parameter]
    public PdfAlign Align { get; set; } = PdfAlign.Left;

    /// <summary>Default paragraph color. Named colors and hexadecimal colors are accepted.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? Color { get; set; }

    /// <summary>Run background color for -Text input. Named colors and hexadecimal colors are accepted.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? BackgroundColor { get; set; }

    /// <summary>Font size for -Text input in PDF points.</summary>
    [Parameter]
    public double? FontSize { get; set; }

    /// <summary>Standard PDF font for -Text input.</summary>
    [Parameter]
    public PdfStandardFont? Font { get; set; }

    /// <summary>Make -Text input bold.</summary>
    [Parameter]
    public SwitchParameter Bold { get; set; }

    /// <summary>Make -Text input italic.</summary>
    [Parameter]
    public SwitchParameter Italic { get; set; }

    /// <summary>Underline -Text input.</summary>
    [Parameter]
    public SwitchParameter Underline { get; set; }

    /// <summary>Strike through -Text input.</summary>
    [Parameter]
    public SwitchParameter Strike { get; set; }

    /// <summary>Baseline for -Text input.</summary>
    [Parameter]
    public PdfTextBaseline Baseline { get; set; } = PdfTextBaseline.Normal;

    /// <summary>Absolute URI link target for -Text input.</summary>
    [Parameter]
    public string? LinkUri { get; set; }

    /// <summary>Named destination link target for -Text input.</summary>
    [Parameter]
    public string? LinkDestinationName { get; set; }

    /// <summary>Optional link annotation contents for -Text input.</summary>
    [Parameter]
    public string? LinkContents { get; set; }

    /// <summary>Emit the updated document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if ((Text == null || Text.Length == 0) && (Run == null || Run.Length == 0))
        {
            throw new PSArgumentException("Provide -Text or -Run content.");
        }

        if (Text is { Length: > 0 } && Run is { Length: > 0 })
        {
            throw new PSArgumentException("Use either -Text or -Run, not both.");
        }

        var defaultColor = PdfCommandUtilities.ParseColor(Color);
        var document = PdfCommandUtilities.ComposeContent(this, Document, ParameterSetName, ParameterSetDocument, content => content.Paragraph(
            builder =>
            {
                if (Run is { Length: > 0 })
                {
                    PdfRichTextRunBuilder.ApplyRuns(builder, Run);
                }
                else
                {
                    PdfRichTextRunBuilder.ApplyText(
                        builder,
                        Text!,
                        Bold.IsPresent,
                        Italic.IsPresent,
                        Underline.IsPresent,
                        Strike.IsPresent,
                        Baseline,
                        defaultColor,
                        PdfCommandUtilities.ParseColor(BackgroundColor),
                        FontSize,
                        Font,
                        LinkUri,
                        LinkDestinationName,
                        LinkContents);
                }
            },
            Align,
            defaultColor));

        if (PassThru.IsPresent && document != null)
        {
            WriteObject(document);
        }
    }
}
