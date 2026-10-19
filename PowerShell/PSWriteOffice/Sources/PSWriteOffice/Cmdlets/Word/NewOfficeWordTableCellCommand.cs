using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Table;
using PSWriteOffice.Services.Text;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Creates a reusable Word table cell definition for explicit table rows.</summary>
/// <example>
///   <summary>Create a full-width Word table section row.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$row = @(New-OfficeWordTableCell -Text 'Identity systems' -ColumnSpan 3)</code>
///   <para>The returned cell can be passed to WordTable inside explicit row arrays.</para>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeWordTableCell")]
[Alias("WordTableCellSpec")]
[OutputType(typeof(OfficeTableCellSpec))]
public sealed class NewOfficeWordTableCellCommand : PSCmdlet
{
    /// <summary>Cell text.</summary>
    [Parameter(Position = 0)]
    public string? Text { get; set; }

    /// <summary>Rich text runs for the cell. Each run can be created with TextRun/WordTextRun or provided as a hashtable/object.</summary>
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

    /// <summary>Cell font size in points.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
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

    /// <summary>Optional Word underline style.</summary>
    [Parameter]
    public WordUnderlineStyle? UnderlineStyle { get; set; }

    /// <summary>Render the cell text with strikethrough.</summary>
    [Parameter]
    public SwitchParameter Strike { get; set; }

    /// <summary>Horizontal cell alignment.</summary>
    [Parameter]
    public WordParagraphAlignment? Align { get; set; }

    /// <summary>Vertical cell alignment.</summary>
    [Parameter]
    public WordTableVerticalAlignment? VerticalAlign { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (!string.IsNullOrEmpty(Text) && Run is { Length: > 0 })
        {
            throw new PSArgumentException("Use either -Text or -Run, not both.");
        }

        var runs = Run is { Length: > 0 } ? OfficeTextRunParser.ParseMany(Run) : null;
        WriteObject(new OfficeTableCellSpec(Text, ColumnSpan, RowSpan, CreateStyle(), runs));
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
            Underline = Underline.IsPresent || UnderlineStyle.HasValue,
            UnderlineStyle = UnderlineStyle?.ToString(),
            Strike = Strike.IsPresent,
            Align = Align?.ToString(),
            VerticalAlign = VerticalAlign?.ToString()
        };

        return style.HasAnyValue ? style : null;
    }
}
