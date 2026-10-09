using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;
using OfficeIMO.Reader.IWork;
using OfficeIMO.Reader;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Decimal_starts_restarts_and_contiguous_counters_are_recovered(IWorkDocumentKind kind) {
        using var package = DecimalCounterPackage(kind);
        var content = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package), kind).Item1.Cells).RichText!;
        Assert.True(content.IsFormattingComplete);
        Assert.Equal(new int?[] { 7, 8, 3, 4 }, content.Paragraphs.Select(p => p.ListNumber));
        Assert.Equal(new int?[] { 7, null, 3, null }, content.Paragraphs.Select(p => p.ListStartNumber));
        Assert.All(content.Paragraphs, paragraph => Assert.Equal("1.", paragraph.ListLabel));
        Assert.Equal(new[] { "7. Value", "8. Next", "3. Restart", "4. Continue" },
            content.Paragraphs.Select(p => IWorkReadProjection.RichTextMarkdown(p)));
        package.Position = 0;
        var policy = new IWorkConversionOptions { Mode = IWorkConversionMode.EditableOnly,
            AllowPartialEditableReconstruction = true };
        if (kind == IWorkDocumentKind.Pages) {
            using var result = IWorkSourceDocument.Open(package).ToWordDocumentResult(policy);
            using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
            using var word = WordprocessingDocument.Open(saved, false);
            var paragraphs = word.MainDocumentPart!.Document!.Body!.Descendants<W.Paragraph>()
                .Where(p => p.InnerText is "Value" or "Next" or "Restart" or "Continue").ToArray();
            Assert.Equal(4, paragraphs.Length);
            int[] ids = paragraphs.Select(p => p.ParagraphProperties!.NumberingProperties!.NumberingId!.Val!.Value).ToArray();
            Assert.Equal(ids[0], ids[1]); Assert.Equal(ids[2], ids[3]); Assert.NotEqual(ids[0], ids[2]);
            var numbering = word.MainDocumentPart.NumberingDefinitionsPart!.Numbering!;
            Assert.Equal(new[] { 7, 3 }, new[] { ids[0], ids[2] }.Select(id => {
                int abstractId = numbering.Elements<W.NumberingInstance>().Single(n => n.NumberID!.Value == id)
                    .AbstractNumId!.Val!.Value;
                return numbering.Elements<W.AbstractNum>().Single(n => n.AbstractNumberId!.Value == abstractId)
                    .Elements<W.Level>().Single().StartNumberingValue!.Val!.Value;
            }));
            Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(word));
        } else if (kind == IWorkDocumentKind.Keynote) {
            using var result = IWorkSourceDocument.Open(package).ToPowerPointPresentationResult(policy);
            using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
            using var pptx = PowerPointPresentation.Load(saved);
            var paragraphs = Assert.Single(Assert.Single(pptx.Slides).Tables).GetCell(0, 0).Paragraphs;
            Assert.Equal(new int?[] { 7, null, 3, null }, paragraphs.Select(p => p.NumberingStartAt));
            Assert.All(paragraphs, paragraph => Assert.Equal(PowerPointNumberingScheme.ArabicPeriod, paragraph.NumberingScheme));
            Assert.Empty(pptx.ValidateDocument());
        }
    }

    [Theory]
    [InlineData("missing-start")]
    [InlineData("invalid-value")]
    [InlineData("unknown-second")]
    [InlineData("inside-paragraph")]
    [InlineData("tiered")]
    [InlineData("nested")]
    public void Unqualified_counter_metadata_never_exposes_a_qualified_number(string fault) {
        byte[] starts = fault switch {
            "missing-start" => AttributeTable(14, ListLevelEntry(0, 0)),
            "invalid-value" => AttributeTable(14, Message(VarintField(1, 0), FloatField(2, 7), VarintField(3, 0))),
            "unknown-second" => AttributeTable(14, ListLevelEntry(0, 7, 1)),
            "inside-paragraph" => AttributeTable(14, ListLevelEntry(1, 7)),
            _ => AttributeTable(14, ListLevelEntry(0, 7), ListLevelEntry(6, 0))
        };
        using var package = DecimalCounterPackage(IWorkDocumentKind.Numbers, starts,
            tiered: fault == "tiered", level: fault == "nested" ? 1ul : 0ul);
        var content = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package),
            IWorkDocumentKind.Numbers).Item1.Cells).RichText!;
        Assert.False(content.IsFormattingComplete);
        Assert.All(content.Paragraphs, paragraph => {
            Assert.Null(paragraph.ListNumber); Assert.Null(paragraph.ListStartNumber);
        });
    }

    [Fact]
    public void Counter_continuation_after_non_list_text_remains_unqualified() {
        using var package = DecimalCounterPackage(IWorkDocumentKind.Numbers,
            listStyles: AttributeTable(7, AttributeEntry(0, ReferenceField(2, 19)),
                AttributeEntry(6, ReferenceField(2, 22)), AttributeEntry(11, ReferenceField(2, 19))));
        var content = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package),
            IWorkDocumentKind.Numbers).Item1.Cells).RichText!;
        // The explicit restart after the gap is qualified; a zero continuation after
        // the gap would require independent evidence before it can reuse a counter.
        Assert.True(content.IsFormattingComplete);
        Assert.Equal(new int?[] { 7, null, 3, 4 }, content.Paragraphs.Select(p => p.ListNumber));
        using var continuation = DecimalCounterPackage(IWorkDocumentKind.Numbers,
            AttributeTable(14, ListLevelEntry(0, 7), ListLevelEntry(6, 0)),
            listStyles: AttributeTable(7, AttributeEntry(0, ReferenceField(2, 19)),
                AttributeEntry(6, ReferenceField(2, 22)), AttributeEntry(11, ReferenceField(2, 19))));
        var unqualified = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(continuation),
            IWorkDocumentKind.Numbers).Item1.Cells).RichText!;
        Assert.False(unqualified.IsFormattingComplete);
        Assert.Equal(new int?[] { 7, null, null, null }, unqualified.Paragraphs.Select(p => p.ListNumber));
    }

    [Fact]
    public void PPTX_counter_range_is_checked_before_destination_mutation() {
        using var package = DecimalCounterPackage(IWorkDocumentKind.Keynote,
            AttributeTable(14, ListLevelEntry(0, 32_767), ListLevelEntry(6, 0)));
        var source = IWorkSourceDocument.Open(package);
        Assert.Equal(32_768, source.ReadKeynote().Slides[0].Tables[0].Cells[0].RichText!.Paragraphs[1].ListNumber);
        Assert.Contains("counter outside", Assert.Throws<InvalidDataException>(() => source.ToPowerPointPresentationResult(new() {
            Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true })).Message);
    }

    [Fact]
    public void Reader_retains_large_counter_metadata_and_reports_Markdown_range_loss() {
        var source = IWorkSourceDocument.Open(CorpusFixture("native-exports/pages-paragraph-list-v15.4.pages"));
        var original = source.ReadPages();
        var selected = original.Body.Paragraphs[3];
        var paragraph = new IWorkTextParagraph(selected.Runs, selected.Style, selected.ListIdentifier,
            0, "1.", IWorkParagraphBreakKind.None, listMarkerKind: IWorkListMarkerKind.Number,
            listNumber: int.MaxValue, listStartNumber: int.MaxValue);
        var body = new IWorkTextContent(new[] { paragraph }, true, true);
        var pages = new IWorkPagesProjection(source, body, Array.Empty<IWorkPagesSection>(),
            Array.Empty<IWorkTextBox>(), Array.Empty<IWorkImageAsset>(), Array.Empty<IWorkTable>(),
            Array.Empty<IWorkPagesDrawable>(), original.PageLayout, Array.Empty<IWorkDiagnostic>(), true);
        var output = new OfficeDocumentReadResult { Kind = ReaderInputKind.IWork };
        var reader = new IWorkReadProjection(output, "large-counter.pages", new ReaderOptions(),
            new ReaderIWorkOptions(), System.Threading.CancellationToken.None);
        reader.AddPages(pages); reader.Complete(source);
        Assert.Equal("2147483647.", Assert.Single(output.Blocks, block => block.Kind == "list-item").Marker);
        Assert.Contains("- List first", output.Markdown);
        Assert.Contains(output.Diagnostics, diagnostic => diagnostic.Code == "IWORK_READER_LIST_COUNTER_LIMIT");
    }

    [Fact]
    public void Decimal_counter_overflow_is_unqualified_instead_of_wrapping() {
        using var package = DecimalCounterPackage(IWorkDocumentKind.Numbers,
            AttributeTable(14, ListLevelEntry(0, int.MaxValue), ListLevelEntry(6, 0),
                ListLevelEntry(11, 3), ListLevelEntry(19, 0)));
        var content = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package),
            IWorkDocumentKind.Numbers).Item1.Cells).RichText!;
        Assert.False(content.IsFormattingComplete);
        Assert.Equal(new int?[] { int.MaxValue, null, 3, 4 }, content.Paragraphs.Select(p => p.ListNumber));
    }

    [Fact]
    public void Counter_boundaries_charge_the_shared_text_budget() {
        using var package = DecimalCounterPackage(IWorkDocumentKind.Numbers,
            AttributeTable(14, Enumerable.Range(0, 20).Select(_ => ListLevelEntry(0, 0)).ToArray()));
        var source = IWorkSourceDocument.Open(package, new IWorkReadOptions { MaximumProjectedTextItems = 10 });
        Assert.Contains("attribute", Assert.Throws<InvalidDataException>(() => source.ReadNumbers()).Message,
            StringComparison.OrdinalIgnoreCase);
    }

    private static MemoryStream DecimalCounterPackage(IWorkDocumentKind kind, byte[]? starts = null,
        bool tiered = false, ulong level = 0, byte[]? listStyles = null) => SelectedRichTextPackage(kind,
        Message(StringField(3, "\nNext\nRestart\nContinue"),
            AttributeTable(6, ListLevelEntry(0, level)),
            starts ?? AttributeTable(14, ListLevelEntry(0, 7), ListLevelEntry(6, 0),
                ListLevelEntry(11, 3), ListLevelEntry(19, 0)),
            listStyles ?? AttributeTable(7, AttributeEntry(0, ReferenceField(2, 19)))),
        additionalRecords: Message(
            ArchiveRecord(19, 2023, Message(VarintField(11, 3), VarintField(11, 3),
                VarintField(15, 0), VarintField(15, 0), VarintField(25, tiered ? 1ul : 0ul), VarintField(25, 0))),
            ArchiveRecord(22, 2023, Message(VarintField(11, 0)))));
}
