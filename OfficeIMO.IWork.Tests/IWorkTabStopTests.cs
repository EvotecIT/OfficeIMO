using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Qualified_tab_stops_survive_document_conversion_or_report_destination_loss(IWorkDocumentKind kind) {
        byte[] tabs = Message(Enumerable.Range(0, 4).Select(i => BytesField(1,
            Message(FloatField(1, (i + 1) * 18), VarintField(2, (ulong)i), StringField(3, "")))).ToArray());
        using var package = ParagraphLayoutPackage(kind, BytesField(25, tabs));
        var text = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package, kind), kind).Item1.Cells).RichText!;
        Assert.True(text.IsFormattingComplete);
        Assert.Equal(new double[] { 18, 36, 54, 72 }, Assert.Single(text.Paragraphs).Style.TabStops!.Select(t => t.PositionPoints));
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
            Assert.Contains(result.Report.Diagnostics, d => d.Code == "IWORK_NUMBERS_EXCEL_RICH_TEXT_PARTIAL");
            Assert.Equal("Value", result.Value.Sheets[0].CellAt(1, 1).GetValue<string>());
            return;
        }
        saved.Position = 0;
        using var archive = new ZipArchive(saved, ZipArchiveMode.Read);
        var part = kind == IWorkDocumentKind.Pages ? archive.GetEntry("word/document.xml")!
            : Assert.Single(archive.Entries, e => e.FullName.StartsWith("ppt/slides/slide", StringComparison.Ordinal)
                && e.FullName.EndsWith(".xml", StringComparison.Ordinal));
        using var stream = part.Open();
        var xml = XDocument.Load(stream);
        XNamespace ns = kind == IWorkDocumentKind.Pages ? "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            : "http://schemas.openxmlformats.org/drawingml/2006/main";
        var paragraph = Assert.Single(xml.Descendants(ns + "p"), p => string.Concat(p.Descendants(ns + "t").Select(t => t.Value)) == "Value");
        var stops = paragraph.Descendants(ns + "tab").ToArray();
        Assert.Equal(4, stops.Length);
        string[] alignments = kind == IWorkDocumentKind.Pages ? ["left", "center", "right", "decimal"] : ["l", "ctr", "r", "dec"];
        for (int i = 0; i < stops.Length; i++) {
            Assert.Equal(alignments[i], (string?)stops[i].Attribute(kind == IWorkDocumentKind.Pages ? ns + "val" : "algn"));
            Assert.Equal(((i + 1) * 18 * (kind == IWorkDocumentKind.Pages ? 20 : 12700)).ToString(),
                (string?)stops[i].Attribute(kind == IWorkDocumentKind.Pages ? ns + "pos" : "pos"));
        }
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    [InlineData(5)]
    public void Unsupported_tab_lists_remain_unassessed_and_do_not_reuse_parent_tabs(int defect) {
        byte[] tab = defect switch {
            0 => Message(FloatField(1, 18), StringField(3, ".")),
            1 => FloatField(1, float.NaN),
            2 => Message(FloatField(1, 18), VarintField(2, 4)),
            3 => Message(FloatField(1, 18), VarintField(4, 0)),
            4 => new byte[] { 0x80 },
            _ => FloatField(1, -18)
        };
        using var package = TabInheritancePackage(BytesField(25, BytesField(1, tab)));
        var text = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package), IWorkDocumentKind.Numbers).Item1.Cells).RichText!;
        Assert.False(text.IsFormattingComplete);
        Assert.Empty(Assert.Single(text.Paragraphs).Style.TabStops!);
        package.Position = 0;
        Assert.Contains(ConvertUnitReport(package, IWorkDocumentKind.Numbers).SourceDeclarationIssues,
            issue => issue.Owner.RecordIdentifier == 19 && issue.FieldPath == "12/25");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Tab_lists_inherit_and_explicit_clear_removes_parent_stops(bool clear) {
        using var package = TabInheritancePackage(clear ? VarintField(24, 1) : Message());
        var text = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package), IWorkDocumentKind.Numbers).Item1.Cells).RichText!;
        Assert.True(text.IsFormattingComplete);
        Assert.Equal(clear ? 0 : 1, Assert.Single(text.Paragraphs).Style.TabStops!.Count);
    }

    [Fact]
    public void Tab_stop_materialization_uses_the_source_wide_text_budget() {
        using var baseline = ParagraphLayoutPackage(IWorkDocumentKind.Numbers, Message());
        Assert.NotEmpty(IWorkSourceDocument.Open(baseline, new IWorkReadOptions { MaximumProjectedTextItems = 20 }).ReadNumbers().Sheets);
        using var package = ParagraphLayoutPackage(IWorkDocumentKind.Numbers, BytesField(25,
            Message(Enumerable.Range(1, 32).Select(i => BytesField(1, FloatField(1, i))).ToArray())));
        var source = IWorkSourceDocument.Open(package, new IWorkReadOptions { MaximumProjectedTextItems = 20 });
        Assert.Contains("Text item", Assert.Throws<InvalidDataException>(() => source.ReadNumbers()).Message);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Table_default_tabs_are_budgeted_for_every_rich_paragraph(bool selectedCellStyle) {
        byte[] style = ArchiveRecord(19, 2022, BytesField(12, BytesField(25,
            Message(Enumerable.Range(1, 20).Select(i => BytesField(1, FloatField(1, i))).ToArray()))));
        byte[] catalog = ArchiveRecord(20, 6005, Message(VarintField(1, 4),
            BytesField(3, Message(VarintField(1, 1), ReferenceField(4, 19)))));
        using var package = SelectedRichTextPackage(IWorkDocumentKind.Numbers,
            selectedText: string.Join("\n", Enumerable.Repeat("Value", 30)),
            modelFields: selectedCellStyle ? Message() : ReferenceField(24, 19),
            storeFields: selectedCellStyle ? ReferenceField(5, 20) : Message(),
            selectCellTextStyle: selectedCellStyle, additionalRecords: Message(style, catalog));
        var bounded = IWorkSourceDocument.Open(package, new IWorkReadOptions { MaximumProjectedTextItems = 200 });
        Assert.Contains("Text item", Assert.Throws<InvalidDataException>(() => bounded.ReadNumbers()).Message);
        package.Position = 0;
        var source = IWorkSourceDocument.Open(package, new IWorkReadOptions { MaximumProjectedTextItems = 1000 });
        var table = Assert.Single(Assert.Single(source.ReadNumbers().Sheets).Tables);
        Assert.Equal(30, Assert.Single(table.Cells).RichText!.Paragraphs.Count);
        Assert.Equal(20, table.GetParagraphStyle(1, 1)!.TabStops!.Count);
    }

    private static MemoryStream TabInheritancePackage(byte[] child) => SelectedRichTextPackage(IWorkDocumentKind.Numbers,
        AttributeTable(5, AttributeEntry(0, ReferenceField(2, 19))), additionalRecords: Message(
            ArchiveRecord(19, 2022, Message(BytesField(1, ReferenceField(3, 20)), BytesField(12, child))),
            ArchiveRecord(20, 2022, BytesField(12, BytesField(25, BytesField(1, FloatField(1, 18)))))));
}
