namespace OfficeIMO.Pdf;

public sealed partial class PdfTextRun {
    /// <summary>Whether the run supplies the descent of its paragraph line box.</summary>
    /// <remarks>Numbering labels supply ascent only; glyph painting retains their complete font metrics.</remarks>
    internal bool ContributesLineBoxDescent { get; set; } = true;

    /// <summary>Whether an empty run retains font metrics without adding ink or text advance.</summary>
    /// <remarks>Used for the paragraph mark of an empty numbered paragraph.</remarks>
    internal bool IsLineBoxStrut { get; set; }
}
