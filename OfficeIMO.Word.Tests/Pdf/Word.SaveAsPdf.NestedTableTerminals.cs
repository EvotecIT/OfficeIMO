using DocumentFormat.OpenXml;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(0, 12D)]
    [InlineData(0, 24D)]
    [InlineData(1, 12D)]
    [InlineData(2, 12D)]
    public void SaveAsPdf_NestedTableCellEndMarkDoesNotAddAParagraphLine(int noteKind, double lineHeight) {
        double structural = MeasureNestedTableTerminalGap(noteKind, lineHeight, NestedTableTerminalKind.Structural);
        double formatted = MeasureNestedTableTerminalGap(noteKind, lineHeight, NestedTableTerminalKind.FormattedStructural);
        double visible = MeasureNestedTableTerminalGap(noteKind, lineHeight, NestedTableTerminalKind.Visible);
        double multiple = MeasureNestedTableTerminalGap(noteKind, lineHeight, NestedTableTerminalKind.TwoEmptyParagraphs);

        // Independent native controls show that formatting a lone cell-end mark
        // does not turn it into another line, while real text and authored blank
        // paragraphs retain their complete source line heights.
        Assert.Equal(structural, formatted, precision: 2);
        Assert.Equal(lineHeight, visible - structural, precision: 2);
        Assert.Equal(2D * lineHeight, multiple - structural, precision: 2);
    }

    private enum NestedTableTerminalKind { Structural, FormattedStructural, Visible, TwoEmptyParagraphs }

    private static double MeasureNestedTableTerminalGap(int noteKind, double lineHeight, NestedTableTerminalKind kind) {
        using WordDocument document = noteKind == 0 ? WordDocument.Create() : CreateFootnoteDiagnosticDocument();
        WordTable outer = document.AddTable(1, 1, WordTableStyle.TableGrid);
        ConfigureNestedLayoutTable(outer, 6000, 6000);
        WordTableCell parent = outer.Rows[0].Cells[0];
        WordTable inner = parent.AddTable(1, 2, WordTableStyle.TableGrid, removePrecedingParagraph: true);
        ConfigureNestedLayoutTable(inner, 6000, 3000, 3000);
        foreach ((WordTableCell Cell, string Text) item in new[] {
            (inner.Rows[0].Cells[0], "TERMINALLEFT"), (inner.Rows[0].Cells[1], "TERMINALRIGHT") }) {
            ConfigureNestedLayoutParagraph(item.Cell.Paragraphs[0], item.Text);
            item.Cell.Paragraphs[0].LineSpacing = checked((int)(lineHeight * 20D));
        }

        W.Paragraph terminal = Assert.Single(parent._tableCell.Elements<W.Paragraph>());
        var terminalView = new WordParagraph(document, terminal);
        ConfigureNestedLayoutParagraph(terminalView, kind == NestedTableTerminalKind.Visible ? "VISIBLETAIL" : string.Empty);
        terminalView.LineSpacing = checked((int)(lineHeight * 20D));
        if (kind == NestedTableTerminalKind.FormattedStructural) {
            terminalView.LineSpacing = checked((int)(2D * lineHeight * 20D));
            terminalView.LineSpacingAfter = global::OfficeIMO.OfficeLength.Points(6D);
        }
        // The structural paragraph has paragraph properties but no empty run
        // carriers. Content children represent an independently preserved boundary.
        if (kind != NestedTableTerminalKind.Visible)
            foreach (OpenXmlElement child in terminal.ChildElements.Where(child => child is not W.ParagraphProperties).ToArray())
                child.Remove();
        if (kind == NestedTableTerminalKind.TwoEmptyParagraphs)
            parent._tableCell.Append(terminal.CloneNode(true));

        WordParagraph after = document.AddParagraph();
        ConfigureNestedLayoutParagraph(after, "TERMINALAFTER");
        after.LineSpacing = checked((int)(lineHeight * 20D));
        if (noteKind != 0) {
            OpenXmlCompositeElement note = AddRichNoteFixture(document, endnote: noteKind == 2);
            note.RemoveAllChildren();
            outer._table.Remove();
            after._paragraph.Remove();
            note.Append(new W.Paragraph(NoteBodyMarker(noteKind == 2), RichNoteRun("TERMINALNOTEBEFORE")),
                outer._table, after._paragraph);
        }
        Assert.Empty(document.ValidateDocument());
        var main = document._wordprocessingDocument.MainDocumentPart!;
        string OriginalXml() => main.Document.OuterXml + main.StyleDefinitionsPart?.Styles?.OuterXml +
            main.DocumentSettingsPart?.Settings?.OuterXml + main.FootnotesPart?.Footnotes?.OuterXml +
            main.EndnotesPart?.Endnotes?.OuterXml;
        string original = OriginalXml();
        PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(noteKind == 0
            ? new WordPdfExportOptions {
                IncludePageNumbers = false, PageSize = new PdfPageSize(OfficeIMO.OfficeLength.Points(400D),OfficeIMO.OfficeLength.Points(500D)), Margins = PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(30D)),
                ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
            } : FootnoteDiagnosticOptions());
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        var words = pdf.GetPages().SelectMany(page => page.GetWords()).ToArray();
        var left = Assert.Single(words, word => word.Text == "TERMINALLEFT");
        var right = Assert.Single(words, word => word.Text == "TERMINALRIGHT");
        var following = Assert.Single(words, word => word.Text == "TERMINALAFTER");
        Assert.Equal(left.BoundingBox.Bottom, right.BoundingBox.Bottom, precision: 2);
        Assert.True(right.BoundingBox.Left > left.BoundingBox.Right);
        if (kind == NestedTableTerminalKind.Visible)
            Assert.Single(words, word => word.Text == "VISIBLETAIL");
        Assert.DoesNotContain(result.Warnings, warning => warning.Code is "native-nested-table-layout-approximated" or
            "native-footnote-placement-approximated" or "native-endnote-content-approximated");
        Assert.Equal(original, OriginalXml());
        return left.BoundingBox.Bottom - following.BoundingBox.Bottom;
    }
}
