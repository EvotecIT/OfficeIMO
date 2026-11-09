using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;
using PSWriteOffice.Services.Table;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Creates a native PDF cell containing text, images or form fields.</summary>
[Cmdlet(VerbsCommon.New, "OfficePdfTableCell", DefaultParameterSetName = "Text")]
[Alias("PdfTableCell")]
[OutputType(typeof(OfficeTableCellSpec))]
public sealed class NewOfficePdfTableCellCommand : OfficeCmdlet {
    /// <summary>Text for this content.</summary>
    [Parameter(ParameterSetName = "Text", Position = 0)]
    public string? Text { get; set; }
    /// <summary>ColumnSpan for this content.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Content")]
    public int ColumnSpan { get; set; } = 1;
    /// <summary>RowSpan for this content.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Content")]
    public int RowSpan { get; set; } = 1;
    /// <summary>Color for this content.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Content")]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? Color { get; set; }
    /// <summary>BackgroundColor for this content.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Content")]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? BackgroundColor { get; set; }
    /// <summary>Size for this content.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Content")]
    [OfficeLengthTransformation]
    public OfficeLength? Size { get; set; }
    /// <summary>Bold for this content.</summary>
    [Parameter(ParameterSetName = "Text")]
    public SwitchParameter Bold { get; set; }
    /// <summary>Italic for this content.</summary>
    [Parameter(ParameterSetName = "Text")]
    public SwitchParameter Italic { get; set; }
    /// <summary>Underline for this content.</summary>
    [Parameter(ParameterSetName = "Text")]
    public OfficeTextDecorationStyle? Underline { get; set; }
    /// <summary>Strikethrough for this content.</summary>
    [Parameter(ParameterSetName = "Text")]
    public SwitchParameter Strikethrough { get; set; }
    /// <summary>Alignment for this content.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Content")]
    public OfficeAlignment? Alignment { get; set; }
    /// <summary>VerticalAlignment for this content.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Content")]
    public OfficeTextVerticalAlignment? VerticalAlignment { get; set; }
    /// <summary>Link for this content.</summary>
    [Parameter(ParameterSetName = "Text")]
    public string? Link { get; set; }
    /// <summary>LinkBookmark for this content.</summary>
    [Parameter(ParameterSetName = "Text")]
    public string? LinkBookmark { get; set; }
    /// <summary>LinkContents for this content.</summary>
    [Parameter(ParameterSetName = "Text")]
    public string? LinkContents { get; set; }
    /// <summary>Bookmark for this content.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Content")]
    public string? Bookmark { get; set; }
    /// <summary>NoWrap for this content.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Content")]
    public SwitchParameter NoWrap { get; set; }
    /// <summary>Content for this content.</summary>
    [Parameter(ParameterSetName = "Content", Mandatory = true, Position = 0)]
    public ScriptBlock? Content { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var content = new PdfCellContent();
        if (Content != null) {
            using (OfficeDocumentContext.Push(content))
            using (OfficeDocumentContext.Push(content.Runs)) Content.InvokeReturnAsIs();
        } else content.Runs.Add(new PdfTextRun(Text ?? string.Empty, bold: Bold, italic: Italic,
            strike: Strikethrough, color: Color, fontSize: Size, underlineStyle: Underline ?? OfficeTextDecorationStyle.None));
        foreach (var run in content.Runs) {
            if (run.Color == null && Color.HasValue) run.SetColor(Color.Value);
            if (run.FontSize == null && Size.HasValue) run.SetFontSize(Size.Value);
        }
        var cell = new PdfTableCell(content.Runs, ColumnSpan, Link, LinkContents, RowSpan,
            content.CheckBoxes, content.Fields, content.Images, LinkBookmark, Bookmark).WithNoWrap(NoWrap);
        string? vertical = VerticalAlignment == OfficeTextVerticalAlignment.Center ? "Middle" : VerticalAlignment?.ToString();
        var style = new OfficeTableCellStyle { FillColor = BackgroundColor?.ToString(), Align = Alignment?.ToString(), VerticalAlign = vertical };
        WriteObject(new OfficeTableCellSpec(cell.Text, ColumnSpan, RowSpan, style.HasAnyValue ? style : null, runs: null, nativeCell: cell));
    }
}
