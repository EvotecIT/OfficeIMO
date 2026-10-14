using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(".")]
    [InlineData(",")]
    public void PPTX_reports_an_explicit_decimal_tab_symbol_omission_and_requires_partial_acceptance(string symbol) {
        using var package = ParagraphLayoutPackage(IWorkDocumentKind.Keynote, Message(StringField(3, symbol),
            BytesField(25, BytesField(1, Message(FloatField(1, 90), VarintField(2, 3))))));
        using var fallback = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package,
            conversionOptions: new() { RequireCompleteVisualCoverage = false });
        Assert.True(fallback.IsVisualFallback);
        package.Position = 0;
        using var partial = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package,
            conversionOptions: new() { AllowPartialEditableReconstruction = true });
        Assert.True(partial.Report.IsPartialEditableReconstruction);
        Assert.Contains(partial.Report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_KEYNOTE_DECIMAL_TAB_SYMBOL_OMITTED"
            && diagnostic.LossKind == global::OfficeIMO.OfficeConversionLossKind.Omission);
    }
    [Theory]
    [InlineData(".")]
    [InlineData(",")]
    public void Uniform_explicit_decimal_tab_symbol_survives_saved_DOCX_settings(string symbol) {
        using var package = DecimalTabPackage(StringField(3, symbol), StringField(3, symbol));
        using var result = WordIWorkConverter.ConvertPagesToWordResult(package,
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly });
        result.RequireCompleteEditableReconstruction();
        Assert.All(result.Projection.Body.Paragraphs, paragraph => Assert.Equal(symbol, paragraph.Style.DecimalTabSymbol));
        using var output = new MemoryStream(); result.Value.Save(output); output.Position = 0;
        using var reopened = WordprocessingDocument.Open(output, false);
        Assert.Equal(symbol, reopened.MainDocumentPart!.DocumentSettingsPart!.Settings!.GetFirstChild<W.DecimalSymbol>()!.Val!.Value);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Different_or_unspecified_decimal_tab_symbols_require_destination_fallback(bool unspecified) {
        using var package = DecimalTabPackage(StringField(3, "."), unspecified ? Array.Empty<byte>() : StringField(3, ","));
        using var result = WordIWorkConverter.ConvertPagesToWordResult(package,
            conversionOptions: new() { AllowPartialEditableReconstruction = true, RequireCompleteVisualCoverage = false });
        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_PAGES_WORD_DESTINATION_UNSUPPORTED"
            && diagnostic.Message.Contains("separators", StringComparison.Ordinal));
        Assert.Empty(result.Report.SourceDeclarationIssues);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void Rejected_decimal_tab_override_clears_its_parent_and_retains_source_evidence(int variant) {
        byte[] child = variant switch {
            0 => StringField(3, "other"),
            1 => Message(StringField(3, "."), StringField(3, ",")),
            2 => VarintField(3, 1),
            _ => Message(VarintField(2, 1), StringField(3, ","))
        };
        using var package = SelectedRichTextPackage(IWorkDocumentKind.Pages,
            AttributeTable(5, AttributeEntry(0, ReferenceField(2, 19))),
            additionalRecords: Message(
                ArchiveRecord(20, 2022, BytesField(12, StringField(3, "."))),
                ArchiveRecord(19, 2022, Message(BytesField(1, ReferenceField(3, 20)), BytesField(12, child)))));
        var projection = IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages).ReadPages();
        var rich = Assert.Single(projection.Tables).Cells.Single().RichText!;
        Assert.Null(Assert.Single(rich.Paragraphs).Style.DecimalTabSymbol);
        Assert.False(rich.IsFormattingComplete);
        Assert.Contains(projection.SourceDeclarationIssues, issue => issue.Owner.RecordIdentifier == 19 && issue.FieldPath == "12/3");
    }

    [Theory]
    [InlineData(0, ".", true)]
    [InlineData(1, null, true)]
    [InlineData(2, null, false)]
    public void Decimal_tab_symbol_distinguishes_inheritance_clear_and_invalid_clear(int variant, string? expected, bool complete) {
        using var package = SelectedRichTextPackage(IWorkDocumentKind.Pages,
            AttributeTable(5, AttributeEntry(0, ReferenceField(2, 19))),
            additionalRecords: Message(
                ArchiveRecord(20, 2022, BytesField(12, StringField(3, "."))),
                ArchiveRecord(19, 2022, Message(BytesField(1, ReferenceField(3, 20)),
                    BytesField(12, variant == 0 ? Message() : VarintField(2, (ulong)variant))))));
        var projection = IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages).ReadPages();
        var rich = Assert.Single(projection.Tables).Cells.Single().RichText!;
        Assert.Equal(expected, Assert.Single(rich.Paragraphs).Style.DecimalTabSymbol);
        Assert.Equal(complete, rich.IsFormattingComplete);
    }

    private static MemoryStream DecimalTabPackage(byte[] first, byte[] second) {
        const string text = "First\t1.2\nSecond\t3,4";
        byte[] tabs = BytesField(25, BytesField(1, Message(FloatField(1, 90), VarintField(2, 3))));
        byte[] records = Message(
            ArchiveRecord(1, 10000, ReferenceField(4, 2), new ulong[] { 2 }),
            ArchiveRecord(2, 2001, Message(StringField(3, text), AttributeTable(5,
                AttributeEntry(0, ReferenceField(2, 3)), AttributeEntry(10, ReferenceField(2, 4)))), new ulong[] { 3, 4 }),
            ArchiveRecord(3, 2022, BytesField(12, Message(first, tabs))),
            ArchiveRecord(4, 2022, BytesField(12, Message(second, tabs))));
        return CreatePackage(("Index/Document.iwa", FrameIwa(records)), ("preview.png", ValidPreviewPng()));
    }

    [Theory]
    [InlineData(0, ",", true)]
    [InlineData(1, ",", true)]
    [InlineData(2, ",", false)]
    [InlineData(3, ",", false)]
    [InlineData(4, ".", true)]
    [InlineData(5, ",", true)]
    [InlineData(6, ".", true)]
    public void Effective_table_decimal_tabs_respect_overrides_clears_and_unused_defaults(int variant, string wordSymbol, bool separatorOmitted) {
        foreach (IWorkDocumentKind kind in new[] { IWorkDocumentKind.Pages, IWorkDocumentKind.Keynote }) {
            byte[] tabs = BytesField(25, BytesField(1, Message(FloatField(1, 90), VarintField(2, 3))));
            byte[] defaults = Message(tabs, variant >= 5 ? Message() : StringField(3, variant == 4 ? "," : "."));
            byte[] selected = variant switch {
                0 => StringField(3, ","), // Inherited stops, explicit separator.
                1 => Message(tabs, StringField(3, ",")), // Fully replaced default.
                2 => VarintField(2, 1), // Explicit separator clear, inherited stops.
                3 => Message(BytesField(25, Message()), StringField(3, ",")), // Explicit tab clear.
                4 => Message(tabs, StringField(3, ".")),
                5 => StringField(3, ","),
                _ => StringField(3, ".")
            };
            using var package = SelectedRichTextPackage(kind,
                AttributeTable(5, AttributeEntry(0, ReferenceField(2, 19))), selectedText: "Value\t1,2",
                modelFields: Message(ReferenceField(24, 30), ReferenceField(25, 31)),
                additionalRecords: Message(ArchiveRecord(30, 2022, BytesField(12, defaults)),
                    ArchiveRecord(31, 2022, BytesField(12, Message(tabs, StringField(3, ",")))),
                    ArchiveRecord(19, 2022, BytesField(12, selected))));
            var source = IWorkSourceDocument.Open(package, kind);
            // The header role is inactive and must not affect active reconstruction.
            var policy = new IWorkConversionOptions { Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true };
            if (kind == IWorkDocumentKind.Pages) {
                using var result = source.ToWordDocumentResult(policy);
                Assert.False(result.IsVisualFallback);
                Assert.Equal(wordSymbol, result.Value.Settings.DecimalSymbol);
                using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
                using var reopened = WordprocessingDocument.Open(saved, false);
                Assert.Equal(wordSymbol, reopened.MainDocumentPart!.DocumentSettingsPart!.Settings!.GetFirstChild<W.DecimalSymbol>()!.Val!.Value);
            } else {
                using var result = source.ToPowerPointPresentationResult(policy);
                Assert.False(result.IsVisualFallback);
                Assert.Equal(separatorOmitted, result.Report.Diagnostics.Any(d => d.Code == "IWORK_KEYNOTE_DECIMAL_TAB_SYMBOL_OMITTED"));
                using var strict = source.ToPowerPointPresentationResult(IWorkTestPolicy.ForIncompletePreview(new()));
                Assert.Equal(separatorOmitted, strict.IsVisualFallback);
            }
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Covered_empty_cell_decimal_styles_do_not_change_active_destination_separators(bool anchorDecimal) {
        foreach (IWorkDocumentKind kind in new[] { IWorkDocumentKind.Pages, IWorkDocumentKind.Keynote }) {
            byte[] first = new byte[16]; first[0] = 5; WriteUInt32(first, 8, 1u << 6); WriteUInt32(first, 12, 1);
            byte[] second = (byte[])first.Clone(); WriteUInt32(second, 12, 2);
            byte[] tabs = BytesField(25, BytesField(1, Message(FloatField(1, 90), VarintField(2, 3))));
            byte[] store = Message(ReferenceField(5, 13), BytesField(3, BytesField(1, Message(VarintField(1, 0), ReferenceField(2, 12)))));
            using var package = TableDependencyPackage(kind, Message(), columns: 2,
                modelPayload: Message(BytesField(4, store), VarintField(6, 1), VarintField(7, 2), MergeTestOwner(MergeTestPair(0, 0, 0, 1))),
                tilePayload: BytesField(5, Message(VarintField(1, 0), BytesField(6, Message(first, second)), BytesField(7, new byte[] { 0, 0, 16, 0 }))),
                additionalRecords: Message(ArchiveRecord(13, 6005, Message(VarintField(1, 4),
                    BytesField(3, Message(VarintField(1, 1), ReferenceField(4, 30))), BytesField(3, Message(VarintField(1, 2), ReferenceField(4, 31))))),
                    ArchiveRecord(30, 2022, BytesField(12, anchorDecimal ? Message(tabs, StringField(3, ".")) : Message())),
                    ArchiveRecord(31, 2022, BytesField(12, Message(tabs, StringField(3, ","))))));
            var source = IWorkSourceDocument.Open(package, kind);
            var policy = new IWorkConversionOptions { Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true };
            if (kind == IWorkDocumentKind.Pages) {
                using var result = source.ToWordDocumentResult(policy);
                Assert.False(result.IsVisualFallback);
                Assert.Single(result.Projection.Tables[0].MergedRanges);
                Assert.All(result.Projection.Tables[0].Cells, cell => Assert.Equal(IWorkCellKind.Empty, cell.Kind));
                Assert.Equal(anchorDecimal ? "." : ",", result.Value.Settings.DecimalSymbol);
            } else {
                using var result = source.ToPowerPointPresentationResult(policy);
                Assert.False(result.IsVisualFallback);
                Assert.Equal(anchorDecimal, result.Report.Diagnostics.Any(d => d.Code == "IWORK_KEYNOTE_DECIMAL_TAB_SYMBOL_OMITTED"));
            }
        }
    }
}
