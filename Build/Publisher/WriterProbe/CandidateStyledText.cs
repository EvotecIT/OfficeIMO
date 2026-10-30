using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Defines the semantic three-run sample independently of its Quill encoding.</summary>
internal static class CandidateStyledText {
    internal static OfficeRichTextParagraph Create() => new(new[] {
        new OfficeRichTextRun("Plain 12 | ", 12, OfficeColor.Black, fontFamily: "Arial"),
        new OfficeRichTextRun("Bold 18", 18, OfficeColor.Black, bold: true, fontFamily: "Arial"),
        new OfficeRichTextRun(" | Plain 12", 12, OfficeColor.Black, fontFamily: "Arial")
    });
}
