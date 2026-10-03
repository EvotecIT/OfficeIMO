using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(5, IWorkDocumentKind.Pages, false)]
    [InlineData(5, IWorkDocumentKind.Pages, true)]
    [InlineData(7, IWorkDocumentKind.Pages, false)]
    [InlineData(7, IWorkDocumentKind.Pages, true)]
    [InlineData(8, IWorkDocumentKind.Pages, false)]
    [InlineData(8, IWorkDocumentKind.Pages, true)]
    [InlineData(11, IWorkDocumentKind.Pages, false)]
    [InlineData(11, IWorkDocumentKind.Pages, true)]
    [InlineData(5, IWorkDocumentKind.Keynote, false)]
    [InlineData(5, IWorkDocumentKind.Keynote, true)]
    [InlineData(7, IWorkDocumentKind.Keynote, false)]
    [InlineData(7, IWorkDocumentKind.Keynote, true)]
    [InlineData(8, IWorkDocumentKind.Keynote, false)]
    [InlineData(8, IWorkDocumentKind.Keynote, true)]
    [InlineData(11, IWorkDocumentKind.Keynote, false)]
    [InlineData(11, IWorkDocumentKind.Keynote, true)]
    public void Decoded_text_attribute_tables_retain_missing_reference_paths(int field, IWorkDocumentKind kind, bool visual) {
        using MemoryStream package = AttributeReferencePackage(kind,
            AttributeTable(field, AttributeEntry(0, ReferenceField(2, 999))));
        IWorkConversionReport report = ConvertUnitReport(package, kind, visual);
        AssertMissingReference(Assert.Single(report.SourceReferenceIssues), kind == IWorkDocumentKind.Pages ? 2ul : 11ul,
            field + "/1[1]/2", 999);
        Assert.DoesNotContain(report.SourceUnits, unit => unit.Identity.RecordIdentifier == 999);
        Assert.Equal(OfficeConversionLossKind.Unassessed, Assert.Single(report.FidelityDiagnostics,
            diagnostic => diagnostic.Code == "IWORK_SOURCE_REFERENCES_UNRESOLVED").LossKind);
    }

    [Theory]
    [InlineData(5, 2022u)]
    [InlineData(7, 2023u)]
    [InlineData(8, 2021u)]
    public void Traversed_paragraph_list_and_character_styles_retain_parent_failures(int field, uint type) {
        using MemoryStream package = AttributeReferencePackage(IWorkDocumentKind.Pages,
            AttributeTable(field, AttributeEntry(0, ReferenceField(2, 10))),
            ArchiveRecord(10, type, Message(BytesField(1, Message(ReferenceField(3, 999))))));
        AssertMissingReference(Assert.Single(ConvertUnitReport(package, IWorkDocumentKind.Pages).SourceReferenceIssues),
            10, "1/3", 999);
    }

    [Fact]
    public void Shared_style_parents_are_reported_once_across_different_inherited_text_styles() {
        using MemoryStream package = AttributeReferencePackage(IWorkDocumentKind.Pages,
            Message(AttributeTable(5, AttributeEntry(0, ReferenceField(2, 12)), AttributeEntry(2, ReferenceField(2, 13))),
                AttributeTable(8, AttributeEntry(0, ReferenceField(2, 10)))),
            Message(ArchiveRecord(10, 2021, Message(BytesField(1, Message(ReferenceField(3, 999))))),
                ArchiveRecord(12, 2022, Message(BytesField(11, Message(FloatField(3, 10))))),
                ArchiveRecord(13, 2022, Message(BytesField(11, Message(FloatField(3, 20)))))), text: "A\nB");
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages,
            new IWorkReadOptions { MaximumSourceReferenceIssues = 1 });
        IWorkPagesProjection pages = source.ReadPages();
        AssertMissingReference(Assert.Single(pages.SourceReferenceIssues), 10, "1/3", 999);
        Assert.Equal(2, pages.Body.Paragraphs.Count);
        Assert.Equal(new double?[] { 10d, 20d }, pages.Body.Paragraphs.Select(paragraph => paragraph.Runs[0].Style.FontSizePoints));
    }

    [Fact]
    public void Attribute_entry_paths_preserve_physical_positions_after_undecodable_siblings() {
        byte[] table = BytesField(8, Message(VarintField(1, 0), BytesField(1, new byte[] { 0x80 }),
            BytesField(1, AttributeEntry(0, ReferenceField(2, 999)))));
        using MemoryStream package = AttributeReferencePackage(IWorkDocumentKind.Pages, table);
        AssertMissingReference(Assert.Single(ConvertUnitReport(package, IWorkDocumentKind.Pages).SourceReferenceIssues),
            2, "8/1[3]/2", 999);
    }

    [Fact]
    public void Rejected_attribute_reference_sets_retain_readable_and_malformed_occurrences() {
        using MemoryStream package = AttributeReferencePackage(IWorkDocumentKind.Pages,
            AttributeTable(8, AttributeEntry(0, Message(ReferenceField(2, 10), BytesField(2, new byte[] { 0x80 })))),
            ArchiveRecord(10, 2021, Message()));
        IWorkConversionReport report = ConvertUnitReport(package, IWorkDocumentKind.Pages);
        Assert.Equal(2, report.SourceReferenceIssues.Count);
        Assert.Equal(IWorkSourceReferenceIssueKind.RejectedReferenceSet, report.SourceReferenceIssues[0].Kind);
        Assert.Equal(10ul, report.SourceReferenceIssues[0].TargetIdentifier);
        Assert.Equal(IWorkSourceReferenceIssueKind.MalformedReference, report.SourceReferenceIssues[1].Kind);
        Assert.Null(report.SourceReferenceIssues[1].TargetIdentifier);
        Assert.All(report.SourceReferenceIssues, issue => Assert.Equal("8/1[1]/2", issue.FieldPath));
    }

    [Fact]
    public void Decoded_end_boundaries_are_assessed_without_traversing_their_unused_style_records() {
        using MemoryStream package = AttributeReferencePackage(IWorkDocumentKind.Pages,
            Message(AttributeTable(8, AttributeEntry(4, ReferenceField(2, 999))),
                AttributeTable(5, AttributeEntry(4, ReferenceField(2, 10)))),
            ArchiveRecord(10, 2022, Message(BytesField(1, Message(ReferenceField(3, 998))))));
        AssertMissingReference(Assert.Single(ConvertUnitReport(package, IWorkDocumentKind.Pages).SourceReferenceIssues),
            2, "8/1[1]/2", 999);
    }

    [Fact]
    public void Invalid_offset_attribute_entries_remain_outside_reference_assessment() {
        using MemoryStream package = AttributeReferencePackage(IWorkDocumentKind.Pages,
            AttributeTable(8, AttributeEntry(5, ReferenceField(2, 999))));
        Assert.Empty(ConvertUnitReport(package, IWorkDocumentKind.Pages).SourceReferenceIssues);
    }

    [Fact]
    public void Reused_body_and_header_storage_reports_attribute_fields_once() {
        using MemoryStream package = AttributeReferencePackage(IWorkDocumentKind.Pages,
            Message(AttributeTable(8, AttributeEntry(0, ReferenceField(2, 999))),
                BytesField(17, Message(BytesField(1, Message(ReferenceField(2, 3)))))),
            Message(ArchiveRecord(3, 10011, Message(ReferenceField(25, 4))),
                ArchiveRecord(4, 10143, Message(ReferenceField(1, 2)))));
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages,
            new IWorkReadOptions { MaximumSourceReferenceIssues = 1 });
        IWorkPagesProjection pages = source.ReadPages();
        AssertMissingReference(Assert.Single(pages.SourceReferenceIssues), 2, "8/1[1]/2", 999);
        Assert.Equal("Text", Assert.Single(pages.HeaderContents).PlainText);
    }

    [Fact]
    public void Ambiguous_style_parent_identifiers_remain_malformed_evidence() {
        using MemoryStream package = AttributeReferencePackage(IWorkDocumentKind.Pages,
            AttributeTable(5, AttributeEntry(0, ReferenceField(2, 10))),
            ArchiveRecord(10, 2022, Message(BytesField(1, Message(BytesField(3,
                Message(VarintField(1, 999), VarintField(1, 998))))))));
        IWorkSourceReferenceIssue issue = Assert.Single(ConvertUnitReport(package, IWorkDocumentKind.Pages).SourceReferenceIssues);
        Assert.Equal(10ul, issue.Owner.RecordIdentifier);
        Assert.Equal("1/3", issue.FieldPath);
        Assert.Equal(IWorkSourceReferenceIssueKind.MalformedReference, issue.Kind);
        Assert.Null(issue.TargetIdentifier);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Attribute_and_parent_reference_limit_failures_are_fatal(IWorkDocumentKind kind) {
        using MemoryStream package = AttributeReferencePackage(kind,
            Message(AttributeTable(5, AttributeEntry(0, ReferenceField(2, 10))),
                AttributeTable(11, AttributeEntry(0, ReferenceField(2, 999)))),
            ArchiveRecord(10, 2022, Message(BytesField(1, Message(ReferenceField(3, 998))))));
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, kind,
            new IWorkReadOptions { MaximumSourceReferenceIssues = 1 });
        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => {
            if (kind == IWorkDocumentKind.Pages) source.ReadPages(); else source.ReadKeynote();
        });
        Assert.Contains("reference issues", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Header_and_presenter_note_attributes_reach_their_conversion_reports(bool note) {
        byte[] storage = Message(StringField(3, "Text"), AttributeTable(11, AttributeEntry(0, ReferenceField(2, 999))));
        byte[] roots = note ? Message(
            ArchiveRecord(1, 1, Message(ReferenceField(2, 2))),
            ArchiveRecord(2, 2, KeynoteShow(Message(ReferenceField(2, 3)))),
            ArchiveRecord(3, 4, Message(ReferenceField(2, 4))),
            ArchiveRecord(4, 5, Message(ReferenceField(27, 5))),
            ArchiveRecord(5, 15, Message(ReferenceField(1, 11))))
            : Message(ArchiveRecord(1, 10000, Message(ReferenceField(4, 2))),
                ArchiveRecord(2, 2001, Message(StringField(3, "Body"), BytesField(17,
                    Message(BytesField(1, Message(ReferenceField(2, 3))))))),
                ArchiveRecord(3, 10011, Message(ReferenceField(25, 4))),
                ArchiveRecord(4, 10143, Message(ReferenceField(1, 11))));
        using MemoryStream package = CreatePackage(("Index/Document.iwa", FrameIwa(Message(roots, ArchiveRecord(11, 2001, storage)))),
            ("preview.png", ValidPreviewPng()));
        AssertMissingReference(Assert.Single(ConvertUnitReport(package,
            note ? IWorkDocumentKind.Keynote : IWorkDocumentKind.Pages).SourceReferenceIssues), 11, "11/1[1]/2", 999);
    }

    [Fact]
    public void Numbers_plain_shape_text_does_not_claim_formatting_reference_assessment() {
        using MemoryStream package = TextReferencePackage(IWorkDocumentKind.Numbers, ReferenceField(2, 11),
            ArchiveRecord(11, 2001, Message(StringField(3, "Text"),
                AttributeTable(8, AttributeEntry(0, ReferenceField(2, 999))))));
        Assert.Empty(ConvertUnitReport(package, IWorkDocumentKind.Numbers).SourceReferenceIssues);
    }

    private static byte[] AttributeEntry(ulong offset, byte[] references) => Message(VarintField(1, offset), references);

    private static byte[] AttributeTable(int field, params byte[][] entries) =>
        BytesField(field, Message(entries.Select(entry => BytesField(1, entry)).ToArray()));

    private static MemoryStream AttributeReferencePackage(IWorkDocumentKind kind, byte[] attributes,
        byte[]? additionalRecords = null, string text = "Text") {
        byte[] payload = Message(StringField(3, text), attributes);
        if (kind == IWorkDocumentKind.Keynote)
            return CreatePackage(("Index/Document.iwa", FrameIwa(Message(
                ArchiveRecord(1, 1, Message(ReferenceField(2, 2))),
                ArchiveRecord(2, 2, KeynoteShow(Message(ReferenceField(2, 3)))),
                ArchiveRecord(3, 4, Message(ReferenceField(2, 4))),
                ArchiveRecord(4, 5, Message(ReferenceField(7, 20))),
                ArchiveRecord(20, 2011, Message(ReferenceField(2, 11))),
                ArchiveRecord(11, 2001, payload), additionalRecords ?? Array.Empty<byte>()))),
                ("preview.png", ValidPreviewPng()));
        return CreatePackage(("Index/Document.iwa", FrameIwa(Message(
            ArchiveRecord(1, 10000, Message(ReferenceField(4, 2))), ArchiveRecord(2, 2001, payload),
            additionalRecords ?? Array.Empty<byte>()))), ("preview.png", ValidPreviewPng()));
    }
}
