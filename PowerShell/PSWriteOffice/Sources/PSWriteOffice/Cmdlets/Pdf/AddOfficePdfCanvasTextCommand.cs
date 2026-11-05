using System;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;
using PSWriteOffice.Services.Table;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Adds plain or rich native text to the active PDF canvas.</summary>
[Cmdlet(VerbsCommon.Add, "OfficePdfCanvasText", DefaultParameterSetName = "Text")]
[Alias("PdfCanvasText")]
public sealed class AddOfficePdfCanvasTextCommand : OfficeCmdlet {
    /// <summary>Text for this content.</summary>
    [Parameter(ParameterSetName = "Text", Mandatory = true, Position = 0)]
    public string? Text { get; set; }
    /// <summary>X for this content.</summary>
    [Parameter(ParameterSetName = "Text", Mandatory = true)]
    [Parameter(ParameterSetName = "Content", Mandatory = true)]
    [OfficeLengthTransformation]
    public OfficeLength X { get; set; } = OfficeLength.Zero;
    /// <summary>Y for this content.</summary>
    [Parameter(ParameterSetName = "Text", Mandatory = true)]
    [Parameter(ParameterSetName = "Content", Mandatory = true)]
    [OfficeLengthTransformation]
    public OfficeLength Y { get; set; } = OfficeLength.Zero;
    /// <summary>Width for this content.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Content")]
    [OfficeLengthTransformation]
    public OfficeLength? Width { get; set; }
    /// <summary>Height for this content.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Content")]
    [OfficeLengthTransformation]
    public OfficeLength? Height { get; set; }
    /// <summary>Alignment for this content.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Content")]
    public OfficeAlignment? Alignment { get; set; }
    /// <summary>LineHeight for this content.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Content")]
    [OfficeLengthTransformation]
    public OfficeLength? LineHeight { get; set; }
    /// <summary>Bold for this content.</summary>
    [Parameter(ParameterSetName = "Text")]
    public SwitchParameter Bold { get; set; }
    /// <summary>Italic for this content.</summary>
    [Parameter(ParameterSetName = "Text")]
    public SwitchParameter Italic { get; set; }
    /// <summary>Strikethrough for this content.</summary>
    [Parameter(ParameterSetName = "Text")]
    public SwitchParameter Strikethrough { get; set; }
    /// <summary>Underline for this content.</summary>
    [Parameter(ParameterSetName = "Text")]
    public OfficeTextDecorationStyle? Underline { get; set; }
    /// <summary>Color for this content.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Content")]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? Color { get; set; }
    /// <summary>BackgroundColor for this content.</summary>
    [Parameter(ParameterSetName = "Text")]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? BackgroundColor { get; set; }
    /// <summary>Font for this content.</summary>
    [Parameter(ParameterSetName = "Text")]
    public string? Font { get; set; }
    /// <summary>Size for this content.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "Content")]
    [OfficeLengthTransformation]
    public OfficeLength? Size { get; set; }
    /// <summary>Baseline for this content.</summary>
    [Parameter(ParameterSetName = "Text")]
    public PdfTextBaseline? Baseline { get; set; }
    /// <summary>Content for this content.</summary>
    [Parameter(ParameterSetName = "Content", Mandatory = true, Position = 0)]
    public ScriptBlock? Content { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var context = PdfCanvasDslContext.Require(this);
        var width = Width ?? OfficeLength.Points(context.Page.Width.ToPoints() - X.ToPoints());
        var height = Height ?? OfficeLength.Points(context.Page.Height.ToPoints() - Y.ToPoints());
        if (width <= OfficeLength.Zero || height <= OfficeLength.Zero) throw new PSArgumentException("The text frame must have positive dimensions inside the page.");
        var runs = new List<PdfTextRun>();
        if (Content != null) { using (OfficeDocumentContext.Push(runs)) Content.InvokeReturnAsIs(); }
        else {
            PdfStandardFont? standard = null; string? family = Font;
            if (Font != null && Enum.TryParse(Font, true, out PdfStandardFont parsed)) { standard = parsed; family = null; }
            runs.Add(new PdfTextRun(Text ?? string.Empty, bold: Bold, italic: Italic, strike: Strikethrough,
                underlineStyle: Underline ?? OfficeTextDecorationStyle.None, backgroundColor: BackgroundColor,
                standardFont: standard, fontFamily: family, baseline: Baseline ?? PdfTextBaseline.Normal));
        }
        foreach (var run in runs) if (run.LinkUri != null || run.LinkDestinationName != null)
            throw new PSArgumentException("Canvas text does not support links; use flowing PdfText.");
        context.Canvas.Text(runs, X, Y, width, height, Color, Alignment ?? OfficeAlignment.Left, Size, LineHeight);
    }
}
