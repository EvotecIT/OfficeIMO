using OfficeIMO.IWork;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Word;
using PdfCore = OfficeIMO.Pdf;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData("explicit")]
    [InlineData("plain-gap")]
    [InlineData("marker-font")]
    [InlineData("third-level")]
    public void Saved_DOCX_child_starts_apply_once_before_parent_reset(string boundary) {
        using var package = ChildResetMappingPackage(IWorkDocumentKind.Pages, boundary);
        using var result = WordIWorkConverter.ConvertPagesToWordResult(package,
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly,
                AllowPartialEditableReconstruction = true });
        var source = Assert.Single(result.Projection.Tables).Cells[0].RichText!;
        Assert.True(source.IsFormattingComplete);
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = WordDocument.Load(saved);
        var indices = WordDocumentTraversal.BuildListIndices(reopened);
        Assert.Equal(source.Paragraphs.Where(paragraph => paragraph.ListNumber.HasValue)
                .Select(paragraph => paragraph.ListNumber!.Value),
            indices.Values.Select(value => value.Index));
        var markers = WordDocumentTraversal.BuildListMarkers(reopened);
        Assert.Equal(source.Paragraphs.Where(paragraph => paragraph.ListNumber.HasValue)
                .Select(paragraph => paragraph.ListNumber + "."),
            markers.Values.Select(value => value.Marker));
        using var xmlStream = new MemoryStream(saved.ToArray());
        using var xml = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(xmlStream, false);
        var numbered = xml.MainDocumentPart!.Document!.Descendants<W.Paragraph>()
            .Select(paragraph => paragraph.ParagraphProperties?.NumberingProperties)
            .Where(properties => properties?.NumberingId != null).ToArray();
        for (int index = 1; index < numbered.Length; index++) {
            if (numbered[index]!.NumberingLevelReference!.Val!.Value
                < numbered[index - 1]!.NumberingLevelReference!.Val!.Value) {
                // The native Pages reference needs a fresh instance here; its
                // importer otherwise advances the parent and reuses child starts.
                Assert.NotEqual(numbered[index - 1]!.NumberingId!.Val!.Value,
                    numbered[index]!.NumberingId!.Val!.Value);
            }
        }
    }

    [Theory]
    [InlineData("explicit")]
    [InlineData("plain-gap")]
    [InlineData("marker-font")]
    [InlineData("third-level")]
    public void Saved_Keynote_PPTX_table_counters_reach_drawing_and_PDF(string boundary) {
        using var package = ChildResetMappingPackage(IWorkDocumentKind.Keynote, boundary);
        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package,
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly,
                AllowPartialEditableReconstruction = true });
        var source = Assert.Single(Assert.Single(result.Projection.Slides).Tables).Cells[0].RichText!;
        Assert.True(source.IsFormattingComplete);
        // Give this counter-only parser fixture room to render every paragraph.
        result.Value.SlideSize.SetSizePoints(600, 500);
        var table = Assert.Single(Assert.Single(result.Value.Slides).Tables);
        table.WidthPoints = 500;
        table.HeightPoints = 400;
        table.SetColumnWidthPoints(0, 500);
        table.SetRowHeightPoints(0, 400);
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = PowerPointPresentation.Load(saved);
        Assert.Empty(reopened.ValidateDocument());
        var actualParagraphs = Assert.Single(reopened.Slides[0].Tables).GetCell(0, 0).Paragraphs;
        for (int index = 1; index < actualParagraphs.Count; index++) {
            if (source.Paragraphs[index].ListNumber.HasValue
                && source.Paragraphs[index - 1].ListNumber.HasValue
                && source.Paragraphs[index].ListLevel < source.Paragraphs[index - 1].ListLevel)
                Assert.Equal(source.Paragraphs[index].ListNumber, actualParagraphs[index].NumberingStartAt);
        }
        string[] expected = source.Paragraphs.Select(paragraph => paragraph.ListNumber is int number
            ? number + ". " + paragraph.Text : paragraph.Text).ToArray();
        var drawing = reopened.Slides[0].CreateVisualSnapshot().Drawing;
        Assert.Equal(string.Concat(expected), DrawingText(drawing));
        using var pdf = reopened.ToPdf(null, out OfficeConversionReport report);
        string text = PdfCore.PdfDocument.Load(new System.IO.MemoryStream(pdf.ToBytes())).ExtractText();
        int previous = -1;
        foreach (string item in expected) {
            int position = text.IndexOf(item, StringComparison.Ordinal);
            Assert.True(position > previous, "Missing or out-of-order iWork list item: " + item + "\n" + text);
            previous = position;
        }
        Assert.Equal(OfficeConversionLossKind.Approximation,
            Assert.Single(report.Diagnostics, diagnostic => diagnostic.Code == OfficeDiagnosticCodes.TableListLayoutApproximation).LossKind);
    }

    // Painted glyph positioning can split a logical paragraph into adjacent runs.
    private static string DrawingText(OfficeDrawing drawing) => string.Concat(drawing.Elements.Select(element => element switch {
        OfficeDrawingText text => text.Text, OfficeDrawingRichText rich => rich.PlainText, _ => string.Empty
    }));

    private static MemoryStream ChildResetMappingPackage(IWorkDocumentKind kind, string boundary) {
        string[] text = boundary == "third-level"
            ? new[] { "Value", "Child one", "Grandchild five", "Grandchild six", "Child two", "Grandchild reset",
                "Parent eight", "Child reset", "Grandchild root reset" }
            : boundary == "explicit"
            ? new[] { "Value", "Child five", "Parent eight", "Child reset", "Child second" }
            : boundary == "plain-gap"
                ? new[] { "Value", "Child one", "Gap", "Child continued", "Parent eight", "Child reset" }
                : new[] { "Value", "Child one", "Child continued", "Parent eight", "Child reset" };
        int[] levels = boundary == "third-level" ? new[] { 0, 1, 2, 2, 1, 2, 0, 1, 2 }
            : boundary == "explicit" ? new[] { 0, 1, 0, 1, 1 }
            : boundary == "plain-gap" ? new[] { 0, 1, 0, 1, 0, 1 } : new[] { 0, 1, 1, 0, 1 };
        ulong[] offsets = ParagraphOffsets(text);
        byte[] styles = boundary == "plain-gap"
            ? AttributeTable(7, AttributeEntry(0, ReferenceField(2, 19)),
                AttributeEntry(offsets[2], ReferenceField(2, 22)),
                AttributeEntry(offsets[3], ReferenceField(2, 19)))
            : AttributeTable(7, AttributeEntry(0, ReferenceField(2, 19)));
        byte[] paragraphStyles = boundary == "marker-font"
            ? AttributeTable(5, AttributeEntry(0, ReferenceField(2, 31)),
                AttributeEntry(offsets[2], ReferenceField(2, 32)))
            : AttributeTable(5, AttributeEntry(0, ReferenceField(2, 31)));
        return SelectedRichTextPackage(kind,
            Message(StringField(3, "\n" + string.Join("\n", text.Skip(1))),
                AttributeTable(6, offsets.Select((offset, index) => ListLevelEntry(offset,
                    (ulong)levels[index], index == 0 || boundary == "explicit" && index == 1
                        || boundary == "third-level" && index == 2 ? 1ul : 0ul)).ToArray()),
                AttributeTable(14, offsets.Select((offset, index) => ListLevelEntry(offset,
                    index == 0 ? 7ul : boundary == "explicit" && index == 1
                        || boundary == "third-level" && index == 2 ? 5ul : 0ul)).ToArray()),
                styles, paragraphStyles, AttributeTable(8, AttributeEntry(0, ReferenceField(2, 33)))),
            additionalRecords: Message(
                ArchiveRecord(19, 2023, Message(VarintField(11, 3), VarintField(11, 3),
                    VarintField(15, 0), VarintField(15, 0), VarintField(25, 0), VarintField(25, 0),
                    FloatField(12, 1.5f), FloatField(12, 1.5f), FloatField(13, 0), FloatField(13, 18),
                    BytesField(14, Message(FloatField(1, 0.9166667f), FloatField(2, 0), VarintField(3, 1))),
                    BytesField(14, Message(FloatField(1, 0.9166667f), FloatField(2, 0), VarintField(3, 1))),
                    boundary == "third-level" ? Message(VarintField(11, 3), VarintField(15, 0), VarintField(25, 0),
                        FloatField(12, 1.5f), FloatField(13, 36),
                        BytesField(14, Message(FloatField(1, 0.9166667f), FloatField(2, 0), VarintField(3, 1)))) : Array.Empty<byte>())),
                ArchiveRecord(22, 2023, VarintField(11, 0)),
                ArchiveRecord(31, 2022, BytesField(11, Message(FloatField(3, 11), StringField(5, "Calibri")))),
                ArchiveRecord(32, 2022, BytesField(11, Message(FloatField(3, 11), StringField(5, "Times New Roman")))),
                ArchiveRecord(33, 2021, BytesField(11, Message(FloatField(3, 12), StringField(5, "Arial"))))));
    }
}
