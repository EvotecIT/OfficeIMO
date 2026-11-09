using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Word;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Table;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Creates a plain or recorded rich Word cell for explicit table rows.</summary>
[Cmdlet(VerbsCommon.New, "OfficeWordTableCell", DefaultParameterSetName = "Text")]
[OutputType(typeof(OfficeTableCellSpec))]
public sealed class NewOfficeWordTableCellCommand : PSCmdlet {
    /// <summary>Text for the cell.</summary>
    [Parameter(ParameterSetName = "Text", Position = 0)]
    public string? Text { get; set; }
    /// <summary>ColumnSpan for the cell.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Content")]
    public int ColumnSpan { get; set; } = 1;
    /// <summary>RowSpan for the cell.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Content")]
    public int RowSpan { get; set; } = 1;
    /// <summary>BackgroundColor for the cell.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Content")]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? BackgroundColor { get; set; }
    /// <summary>Alignment for the cell.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Content")]
    public OfficeAlignment? Alignment { get; set; }
    /// <summary>VerticalAlignment for the cell.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Content")]
    public WordTableVerticalAlignment? VerticalAlignment { get; set; }
    /// <summary>Bold for the cell.</summary>
    [Parameter(ParameterSetName = "Text")]
    public SwitchParameter Bold { get; set; }
    /// <summary>Italic for the cell.</summary>
    [Parameter(ParameterSetName = "Text")]
    public SwitchParameter Italic { get; set; }
    /// <summary>Strikethrough for the cell.</summary>
    [Parameter(ParameterSetName = "Text")]
    public SwitchParameter Strikethrough { get; set; }
    /// <summary>Underline for the cell.</summary>
    [Parameter(ParameterSetName = "Text")]
    public WordUnderlineStyle? Underline { get; set; }
    /// <summary>Color for the cell.</summary>
    [Parameter(ParameterSetName = "Text")]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? Color { get; set; }
    /// <summary>Font for the cell.</summary>
    [Parameter(ParameterSetName = "Text")]
    public string? Font { get; set; }
    /// <summary>Size for the cell.</summary>
    [Parameter(ParameterSetName = "Text")]
    [OfficeLengthTransformation]
    public OfficeLength? Size { get; set; }
    /// <summary>Content for the cell.</summary>
    [Parameter(ParameterSetName = "Content", Mandatory = true, Position = 0)]
    public ScriptBlock? Content { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        var content = new WordCellContent();
        if (Content != null) { using (OfficeDocumentContext.Push(content)) Content.InvokeReturnAsIs(); }
        else {
            string text = Text ?? string.Empty;
            bool bold = Bold, italic = Italic, strike = Strikethrough;
            var underline = Underline; var color = Color; var font = Font; var size = Size;
            content.Text.Add(paragraph => {
                var run = paragraph.AddFormattedText(text, bold, italic, underline);
                run.Strikethrough = strike;
                if (color.HasValue) run.Color = color.Value;
                if (font != null) run.FontFamily = font;
                if (size.HasValue) run.FontSize = size.Value;
            });
        }
        var style = new OfficeTableCellStyle {
            FillColor = BackgroundColor?.ToString(), Align = Alignment?.ToString(), VerticalAlign = VerticalAlignment?.ToString()
        };
        WriteObject(new OfficeTableCellSpec(Text, ColumnSpan, RowSpan, style.HasAnyValue ? style : null, runs: null, nativeCell: content));
    }
}
