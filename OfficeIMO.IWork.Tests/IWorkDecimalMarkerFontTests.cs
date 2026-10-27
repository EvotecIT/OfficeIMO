using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void DOCX_marker_font_changes_cannot_restyle_an_earlier_numbering_scope() {
        using var package = DecimalMarkerFontPackage(IWorkDocumentKind.Pages);
        using var result = WordIWorkConverter.ConvertPagesToWordResult(package,
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly,
                AllowPartialEditableReconstruction = true });
        Assert.True(Assert.Single(result.Projection.Tables).Cells[0].RichText!.IsFormattingComplete);
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var word = WordprocessingDocument.Open(saved, false);
        var paragraphs = word.MainDocumentPart!.Document!.Body!.Descendants<W.Paragraph>()
            .Where(paragraph => paragraph.InnerText is "Value" or "Next").ToArray();
        int[] ids = paragraphs.Select(paragraph => paragraph.ParagraphProperties!.NumberingProperties!.NumberingId!.Val!.Value).ToArray();
        Assert.NotEqual(ids[0], ids[1]);
        var numbering = word.MainDocumentPart.NumberingDefinitionsPart!.Numbering!;
        W.Level Level(int id) {
            int abstractId = numbering.Elements<W.NumberingInstance>().Single(number => number.NumberID!.Value == id)
                .AbstractNumId!.Val!.Value;
            return numbering.Elements<W.AbstractNum>().Single(number => number.AbstractNumberId!.Value == abstractId)
                .Elements<W.Level>().Single();
        }
        Assert.Equal(new[] { "Calibri", "Times New Roman" }, ids.Select(id => Level(id)
            .GetFirstChild<W.NumberingSymbolRunProperties>()!.GetFirstChild<W.RunFonts>()!.Ascii!.Value));
        Assert.Equal(new[] { 7, 8 }, ids.Select(id => Level(id).StartNumberingValue!.Val!.Value));
        Assert.All(ids, id => Assert.Equal("22", Level(id).GetFirstChild<W.NumberingSymbolRunProperties>()!
            .GetFirstChild<W.FontSize>()!.Val!.Value));
        Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(word));
    }

    [Fact]
    public void PPTX_decimal_markers_use_exact_point_size_independently_of_body_runs() {
        using var package = DecimalMarkerFontPackage(IWorkDocumentKind.Keynote);
        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package,
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly,
                AllowPartialEditableReconstruction = true });
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = PowerPointDocument.Load(saved);
        var paragraphs = Assert.Single(Assert.Single(reopened.Slides).Tables).GetCell(0, 0).Paragraphs;
        Assert.All(paragraphs, paragraph => {
            Assert.Equal(11, paragraph.BulletSizePoints);
            Assert.Null(paragraph.BulletSizePercent);
            Assert.Equal(16.5d, paragraph.LeftMarginPoints);
            Assert.Equal(-16.5d, paragraph.IndentPoints);
            Assert.All(paragraph.Runs.Where(run => run.Text.Length > 0), run => Assert.Equal(12, run.FontSize));
        });
        Assert.Equal(new[] { "Calibri", "Times New Roman" }, paragraphs.Select(paragraph => paragraph.BulletFontName));
        Assert.Empty(reopened.ValidateDocument());
    }

    private static MemoryStream DecimalMarkerFontPackage(IWorkDocumentKind kind) => SelectedRichTextPackage(kind,
        Message(StringField(3, "\nNext"), AttributeTable(6, ListLevelEntry(0, 0)),
            AttributeTable(14, ListLevelEntry(0, 7), ListLevelEntry(6, 0)),
            AttributeTable(7, AttributeEntry(0, ReferenceField(2, 19))),
            AttributeTable(5, AttributeEntry(0, ReferenceField(2, 31)), AttributeEntry(6, ReferenceField(2, 32))),
            AttributeTable(8, AttributeEntry(0, ReferenceField(2, 33)))),
        additionalRecords: Message(
            ArchiveRecord(19, 2023, Message(VarintField(11, 3), VarintField(15, 0), VarintField(25, 0),
                FloatField(12, 1.5f), FloatField(13, 0),
                BytesField(14, Message(FloatField(1, 0.9166667f), FloatField(2, 0), VarintField(3, 1))))),
            ArchiveRecord(31, 2022, BytesField(11, Message(FloatField(3, 11), StringField(5, "Calibri")))),
            ArchiveRecord(32, 2022, BytesField(11, Message(FloatField(3, 11), StringField(5, "Times New Roman")))),
            ArchiveRecord(33, 2021, BytesField(11, Message(FloatField(3, 12), StringField(5, "Arial"))))));
}
