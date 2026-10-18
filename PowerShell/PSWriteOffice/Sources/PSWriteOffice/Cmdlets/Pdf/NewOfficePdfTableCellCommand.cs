using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;
using PSWriteOffice.Services.Table;
using PSWriteOffice.Services.Text;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Creates a reusable PDF table cell definition for explicit table rows.</summary>
/// <example>
///   <summary>Create a full-width PDF table section row.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$row = @(New-OfficePdfTableCell -Text 'Identity systems' -ColumnSpan 3 -FillColor '#DBEAFE' -TextColor '#1E3A8A' -Bold)</code>
///   <para>The returned cell can be passed to PdfTable inside explicit row arrays.</para>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficePdfTableCell")]
[Alias("PdfTableCell")]
[OutputType(typeof(OfficeTableCellSpec))]
public sealed class NewOfficePdfTableCellCommand : PSCmdlet
{
    /// <summary>Cell text.</summary>
    [Parameter(Position = 0)]
    public string? Text { get; set; }

    /// <summary>Rich text runs for the cell. Each run can be created with TextRun/PdfTextRun or provided as a hashtable/object.</summary>
    [Parameter]
    [Alias("Runs")]
    public object[]? Run { get; set; }

    /// <summary>Number of logical columns covered by the cell.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int ColumnSpan { get; set; } = 1;

    /// <summary>Number of logical rows covered by the cell.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int RowSpan { get; set; } = 1;

    /// <summary>Cell text color. Named colors and hexadecimal colors are accepted.</summary>
    [Parameter]
    [Alias("Color", "FontColor")]
    [OfficeColorArgumentTransformation]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? TextColor { get; set; }

    /// <summary>Cell fill color. Named colors and hexadecimal colors are accepted.</summary>
    [Parameter]
    [Alias("BackgroundColor", "CellFill")]
    [OfficeColorArgumentTransformation]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? FillColor { get; set; }

    /// <summary>Cell font size in PDF points.</summary>
    [Parameter]
    [ValidateRange(0.1D, double.MaxValue)]
    public double? FontSize { get; set; }

    /// <summary>Render the cell text in bold.</summary>
    [Parameter]
    public SwitchParameter Bold { get; set; }

    /// <summary>Render the cell text in italics.</summary>
    [Parameter]
    public SwitchParameter Italic { get; set; }

    /// <summary>Render the cell text with underline.</summary>
    [Parameter]
    public SwitchParameter Underline { get; set; }

    /// <summary>Optional underline style name. PDF table rendering treats any supported value as underline.</summary>
    [Parameter]
    [ArgumentCompleter(typeof(OfficeUnderlineStyleArgumentCompleter))]
    public string? UnderlineStyle { get; set; }

    /// <summary>Render the cell text with strikethrough.</summary>
    [Parameter]
    public SwitchParameter Strike { get; set; }

    /// <summary>Horizontal cell alignment.</summary>
    [Parameter]
    public PdfColumnAlign? Align { get; set; }

    /// <summary>Vertical cell alignment.</summary>
    [Parameter]
    public PdfCellVerticalAlign? VerticalAlign { get; set; }

    /// <summary>Typed check boxes rendered inside the cell.</summary>
    [Parameter]
    [Alias("CheckBoxes")]
    public PdfTableCellCheckBox[]? CheckBox { get; set; }

    /// <summary>Typed images rendered inside the cell.</summary>
    [Parameter]
    [Alias("Images")]
    public PdfTableCellImage[]? Image { get; set; }

    /// <summary>Typed text or choice form fields rendered inside the cell.</summary>
    [Parameter]
    [Alias("FormFields")]
    public PdfTableCellFormField[]? FormField { get; set; }

    /// <summary>Absolute or catalog-base-relative URI linked from the cell.</summary>
    [Parameter]
    public string? LinkUri { get; set; }

    /// <summary>Named PDF destination linked from the cell.</summary>
    [Parameter]
    public string? LinkDestinationName { get; set; }

    /// <summary>Accessible annotation text for the cell link.</summary>
    [Parameter]
    public string? LinkContents { get; set; }

    /// <summary>Named PDF destination defined at this cell.</summary>
    [Parameter]
    public string? NamedDestinationName { get; set; }

    /// <summary>Keep the cell content on one visual line.</summary>
    [Parameter]
    public SwitchParameter NoWrap { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (!string.IsNullOrEmpty(Text) && Run is { Length: > 0 })
        {
            throw new PSArgumentException("Use either -Text or -Run, not both.");
        }

        var runs = Run is { Length: > 0 } ? OfficeTextRunParser.ParseMany(Run) : null;
        var style = CreateStyle();
        if (!HasTypedPdfContent())
        {
            WriteObject(new OfficeTableCellSpec(Text, ColumnSpan, RowSpan, style, runs));
            return;
        }

        var nativeCell = CreateTypedPdfCell(style, runs);
        WriteObject(new OfficeTableCellSpec(
            nativeCell.Text,
            nativeCell.ColumnSpan,
            nativeCell.RowSpan,
            style,
            runs: null,
            nativeCell: nativeCell));
    }

    private bool HasTypedPdfContent()
        => CheckBox is { Length: > 0 } ||
           Image is { Length: > 0 } ||
           FormField is { Length: > 0 } ||
           LinkUri != null ||
           LinkDestinationName != null ||
           LinkContents != null ||
           NamedDestinationName != null ||
           NoWrap.IsPresent;

    private PdfTableCell CreateTypedPdfCell(OfficeTableCellStyle? style, OfficeTextRunSpec[]? runs)
    {
        PdfTableCell cell;
        if (runs is { Length: > 0 })
        {
            var styledRuns = OfficeTableCellTextRunStyle.Apply(runs, style);
            cell = new PdfTableCell(
                PdfRichTextRunBuilder.ToTextRuns(styledRuns),
                ColumnSpan,
                LinkUri,
                LinkContents,
                RowSpan,
                CheckBox,
                FormField,
                Image,
                LinkDestinationName,
                NamedDestinationName);
        }
        else if (style?.HasTextStyle == true)
        {
            var run = new PdfTextRun(
                Text ?? string.Empty,
                bold: style.Bold,
                underline: style.Underline,
                color: PdfCommandUtilities.ParseColor(style.TextColor),
                italic: style.Italic,
                strike: style.Strike,
                fontSize: style.FontSize);
            cell = new PdfTableCell(
                new[] { run },
                ColumnSpan,
                LinkUri,
                LinkContents,
                RowSpan,
                CheckBox,
                FormField,
                Image,
                LinkDestinationName,
                NamedDestinationName);
        }
        else
        {
            cell = new PdfTableCell(
                Text,
                ColumnSpan,
                LinkUri,
                LinkContents,
                RowSpan,
                CheckBox,
                FormField,
                Image,
                LinkDestinationName,
                NamedDestinationName);
        }

        return NoWrap.IsPresent ? cell.WithNoWrap() : cell;
    }

    private OfficeTableCellStyle? CreateStyle()
    {
        var style = new OfficeTableCellStyle
        {
            TextColor = TextColor,
            FillColor = FillColor,
            FontSize = FontSize,
            Bold = Bold.IsPresent,
            Italic = Italic.IsPresent,
            Underline = Underline.IsPresent || !string.IsNullOrWhiteSpace(UnderlineStyle),
            UnderlineStyle = UnderlineStyle,
            Strike = Strike.IsPresent,
            Align = Align?.ToString(),
            VerticalAlign = VerticalAlign?.ToString()
        };

        return style.HasAnyValue ? style : null;
    }
}
