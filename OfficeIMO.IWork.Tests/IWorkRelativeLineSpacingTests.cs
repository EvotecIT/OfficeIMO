using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Relative_line_spacing_is_recovered_and_destination_layout_is_preserved_or_reported(IWorkDocumentKind kind) {
        using var package = ParagraphLayoutPackage(kind, BytesField(13, FloatField(2, 1.15f)));
        var text = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package, kind), kind).Item1.Cells).RichText!;
        Assert.True(text.IsFormattingComplete);
        Assert.Equal((double)1.15f, Assert.Single(text.Paragraphs).Style.LineSpacingMultiplier);
        package.Position = 0;
        using var saved = new MemoryStream();
        var options = new IWorkConversionOptions { AllowPartialEditableReconstruction = true };
        if (kind == IWorkDocumentKind.Pages) {
            using var result = WordIWorkConverter.ConvertPagesToWordResult(package, conversionOptions: options);
            Assert.Empty(result.Report.SourceDeclarationIssues);
            result.Value.Save(saved);
        } else if (kind == IWorkDocumentKind.Keynote) {
            using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package, conversionOptions: options);
            Assert.Empty(result.Report.SourceDeclarationIssues);
            result.Value.Save(saved);
        } else {
            using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package, conversionOptions: options);
            Assert.Empty(result.Report.SourceDeclarationIssues);
            Assert.Contains(result.Report.Diagnostics, d => d.Code == "IWORK_NUMBERS_EXCEL_RICH_TEXT_PARTIAL");
            Assert.Equal("Value", result.Value.Sheets[0].CellAt(1, 1).GetValue<string>());
            return;
        }
        saved.Position = 0;
        using var archive = new ZipArchive(saved, ZipArchiveMode.Read);
        var part = kind == IWorkDocumentKind.Pages ? archive.GetEntry("word/document.xml")!
            : Assert.Single(archive.Entries, entry => entry.FullName.StartsWith("ppt/slides/slide", StringComparison.Ordinal)
                && entry.FullName.EndsWith(".xml", StringComparison.Ordinal));
        using var stream = part.Open();
        var xml = XDocument.Load(stream);
        XNamespace ns = kind == IWorkDocumentKind.Pages
            ? "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            : "http://schemas.openxmlformats.org/drawingml/2006/main";
        var paragraph = Assert.Single(xml.Descendants(ns + "p"), p => string.Concat(p.Descendants(ns + "t").Select(t => t.Value)) == "Value");
        if (kind == IWorkDocumentKind.Pages) {
            var spacing = paragraph.Element(ns + "pPr")!.Element(ns + "spacing")!;
            Assert.Equal("auto", (string?)spacing.Attribute(ns + "lineRule"));
            Assert.Equal("276", (string?)spacing.Attribute(ns + "line"));
        } else {
            Assert.Equal("115000", (string?)paragraph.Element(ns + "pPr")!.Element(ns + "lnSpc")!.Element(ns + "spcPct")!.Attribute("val"));
        }
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    [InlineData(5)]
    public void Unqualified_line_spacing_does_not_reuse_an_ancestor_multiplier(int variant) {
        byte[] spacing = variant switch {
            1 => Message(VarintField(1, 2), FloatField(2, 12)),
            2 => Message(FloatField(2, 1.5f), FloatField(3, 0)),
            3 => Message(FloatField(2, 1.5f), FloatField(2, 2)),
            4 => FloatField(2, float.NaN),
            _ => Message(FloatField(2, 1.5f), VarintField(4, 1))
        };
        using var package = SelectedRichTextPackage(IWorkDocumentKind.Numbers,
            AttributeTable(5, AttributeEntry(0, ReferenceField(2, 19))),
            additionalRecords: Message(
                ArchiveRecord(19, 2022, Message(BytesField(1, ReferenceField(3, 20)), BytesField(12, BytesField(13, spacing)))),
                ArchiveRecord(20, 2022, BytesField(12, BytesField(13, FloatField(2, 2))))));
        var text = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package), IWorkDocumentKind.Numbers).Item1.Cells).RichText!;
        Assert.False(text.IsFormattingComplete);
        Assert.Null(Assert.Single(text.Paragraphs).Style.LineSpacingMultiplier);
        package.Position = 0;
        Assert.Contains(ConvertUnitReport(package, IWorkDocumentKind.Numbers).SourceDeclarationIssues,
            issue => issue.Owner.RecordIdentifier == 19 && issue.FieldPath == "12/13");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Relative_line_spacing_inherits_and_explicit_clear_removes_the_ancestor_value(bool clear) {
        using var package = SelectedRichTextPackage(IWorkDocumentKind.Numbers,
            AttributeTable(5, AttributeEntry(0, ReferenceField(2, 19))),
            additionalRecords: Message(
                ArchiveRecord(19, 2022, Message(BytesField(1, ReferenceField(3, 20)), BytesField(12, clear ? VarintField(12, 1) : Message()))),
                ArchiveRecord(20, 2022, BytesField(12, BytesField(13, Message(VarintField(1, 0), FloatField(2, 2)))))));
        var text = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package), IWorkDocumentKind.Numbers).Item1.Cells).RichText!;
        Assert.True(text.IsFormattingComplete);
        Assert.Equal(clear ? null : (double?)2, Assert.Single(text.Paragraphs).Style.LineSpacingMultiplier);
    }
    [Theory]
    [InlineData(IWorkDocumentKind.Pages, 10000000f)]
    [InlineData(IWorkDocumentKind.Pages, 0.000001f)]
    [InlineData(IWorkDocumentKind.Keynote, 133f)]
    [InlineData(IWorkDocumentKind.Keynote, 0.000001f)]
    [InlineData(IWorkDocumentKind.Keynote, 1.333333f)]
    public void Relative_line_spacing_outside_destination_range_or_precision_uses_fallback(IWorkDocumentKind kind, float multiplier) {
        using var package = ParagraphLayoutPackage(kind, BytesField(13, FloatField(2, multiplier)));
        var options = new IWorkConversionOptions { AllowPartialEditableReconstruction = true };
        if (kind == IWorkDocumentKind.Pages) {
            using var result = WordIWorkConverter.ConvertPagesToWordResult(package, conversionOptions: IWorkTestPolicy.ForIncompletePreview(options));
            Assert.True(result.IsVisualFallback);
        } else {
            using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package, conversionOptions: IWorkTestPolicy.ForIncompletePreview(options));
            Assert.True(result.IsVisualFallback);
        }
    }

    [Fact]
    public void Reader_reports_relative_spacing_not_represented_by_Markdown() {
        using var package = CreatePackage(("Index/Document.iwa", FrameIwa(Message(
            ArchiveRecord(1, 10000, ReferenceField(4, 2)),
            ArchiveRecord(2, 2001, Message(StringField(3, "Body"), AttributeTable(5, AttributeEntry(0, ReferenceField(2, 19))))),
            ArchiveRecord(19, 2022, BytesField(12, BytesField(13, FloatField(2, 1.5f))))))));
        var result = OfficeIMO.Reader.IWork.IWorkReaderAdapter.ReadDocument(package, "spacing.pages",
            new OfficeIMO.Reader.ReaderOptions(), new OfficeIMO.Reader.IWork.ReaderIWorkOptions(), System.Threading.CancellationToken.None);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "IWORK_READER_TEXT_STYLE_PARTIAL");
        Assert.Contains(result.Blocks, block => block.Text == "Body");
    }

}
