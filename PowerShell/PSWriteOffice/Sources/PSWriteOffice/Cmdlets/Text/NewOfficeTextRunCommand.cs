using System.Management.Automation;
using PSWriteOffice.Services.Text;

namespace PSWriteOffice.Cmdlets.Text;

/// <summary>Creates a reusable rich text run specification for Word, Excel, PowerPoint, and PDF commands.</summary>
[Cmdlet(VerbsCommon.New, "OfficeTextRun")]
[Alias("TextRun", "PdfTextRun", "WordTextRun", "ExcelTextRun", "PowerPointTextRun", "PptTextRun")]
[OutputType(typeof(OfficeTextRunSpec))]
public sealed class NewOfficeTextRunCommand : PSCmdlet
{
    /// <summary>Run text.</summary>
    [Parameter(Position = 0)]
    public string? Text { get; set; }

    /// <summary>Run kind such as Text, LineBreak, Tab, Superscript, or Subscript.</summary>
    [Parameter]
    [ArgumentCompleter(typeof(OfficeTextRunKindArgumentCompleter))]
    public string? Kind { get; set; }

    /// <summary>Render the run in bold.</summary>
    [Parameter]
    public SwitchParameter Bold { get; set; }

    /// <summary>Render the run in italics.</summary>
    [Parameter]
    public SwitchParameter Italic { get; set; }

    /// <summary>Render the run with underline.</summary>
    [Parameter]
    public SwitchParameter Underline { get; set; }

    /// <summary>Optional underline style name when the target format supports it.</summary>
    [Parameter]
    [ArgumentCompleter(typeof(OfficeUnderlineStyleArgumentCompleter))]
    public string? UnderlineStyle { get; set; }

    /// <summary>Render the run with strikethrough.</summary>
    [Parameter]
    public SwitchParameter Strike { get; set; }

    /// <summary>Text color. Named colors and hexadecimal colors are accepted.</summary>
    [Parameter]
    [Alias("TextColor", "FontColor")]
    [OfficeColorArgumentTransformation]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? Color { get; set; }

    /// <summary>Run background or highlight color. Named colors and hexadecimal colors are accepted; None clears the background or highlight where supported.</summary>
    [Parameter]
    [Alias("HighlightColor", "FillColor")]
    [OfficeColorArgumentTransformation(AllowNone = true)]
    [ArgumentCompleter(typeof(OfficeHighlightColorArgumentCompleter))]
    public string? BackgroundColor { get; set; }

    /// <summary>Font size in points.</summary>
    [Parameter]
    public double? FontSize { get; set; }

    /// <summary>Font name, family, or target-specific font identifier.</summary>
    [Parameter]
    [Alias("Font", "Typeface", "FontFamily")]
    public string? FontName { get; set; }

    /// <summary>Target-specific baseline name, such as Superscript or Subscript.</summary>
    [Parameter]
    [ArgumentCompleter(typeof(OfficeTextBaselineArgumentCompleter))]
    public string? Baseline { get; set; }

    /// <summary>URI link target when supported by the target format.</summary>
    [Parameter]
    [Alias("Uri", "Url", "Href")]
    public string? LinkUri { get; set; }

    /// <summary>Named destination or bookmark target when supported by the target format.</summary>
    [Parameter]
    [Alias("DestinationName", "Bookmark", "BookmarkName")]
    public string? LinkDestinationName { get; set; }

    /// <summary>Optional link tooltip or annotation contents.</summary>
    [Parameter]
    [Alias("Contents", "Tooltip")]
    public string? LinkContents { get; set; }

    /// <summary>PDF tab leader style name.</summary>
    [Parameter]
    [Alias("Leader")]
    [ArgumentCompleter(typeof(OfficeTabLeaderArgumentCompleter))]
    public string? TabLeader { get; set; }

    /// <summary>Tab alignment name.</summary>
    [Parameter]
    [Alias("Alignment")]
    [ArgumentCompleter(typeof(OfficeTabAlignmentArgumentCompleter))]
    public string? TabAlignment { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        WriteObject(new OfficeTextRunSpec
        {
            Text = Text ?? string.Empty,
            Kind = Kind,
            Bold = Bold.IsPresent,
            Italic = Italic.IsPresent,
            Underline = Underline.IsPresent || !string.IsNullOrWhiteSpace(UnderlineStyle),
            UnderlineStyle = UnderlineStyle,
            Strike = Strike.IsPresent,
            Color = Color,
            BackgroundColor = BackgroundColor,
            FontSize = FontSize,
            FontName = FontName,
            Baseline = Baseline,
            LinkUri = LinkUri,
            LinkDestinationName = LinkDestinationName,
            LinkContents = LinkContents,
            TabLeader = TabLeader,
            TabAlignment = TabAlignment
        });
    }
}
