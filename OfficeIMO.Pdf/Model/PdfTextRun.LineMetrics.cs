namespace OfficeIMO.Pdf;

public sealed partial class PdfTextRun {
    /// <summary>Whether the run supplies the descent of its paragraph line box.</summary>
    /// <remarks>Numbering labels supply ascent only; glyph painting retains their complete font metrics.</remarks>
    internal bool ContributesLineBoxDescent { get; set; } = true;

    /// <summary>Whether an empty run retains font metrics without adding ink or text advance.</summary>
    /// <remarks>Used for the paragraph mark of an empty numbered paragraph.</remarks>
    internal bool IsLineBoxStrut { get; set; }

    /// <summary>Vertical offset of the glyphs from the line's baseline, in points; positive raises them.</summary>
    /// <remarks>Imported superscript and subscript text carries its own size and offset; the line box ignores the offset.</remarks>
    internal double BaselineShift { get; set; }

    /// <summary>Copies this run as baseline text at an explicit size, raised or lowered by <paramref name="baselineShift"/>.</summary>
    internal PdfTextRun WithScriptGeometry(double fontSize, double baselineShift) {
        PdfTextRun copy = CopyRunFormattingTo(new PdfTextRun(Text, Bold, Underline, Color, Italic, Strike,
            global::OfficeIMO.OfficeLength.Points(fontSize), StandardFont, LinkUri, LinkContents, PdfTextBaseline.Normal,
            LinkDestinationName, TabLeader, TabAlignment, BackgroundColor, FontFamily, UnderlineStyle, StrikeStyle, DecorationColor));
        copy.BaselineShift = baselineShift;
        return copy;
    }
}
