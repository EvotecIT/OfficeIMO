using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Unused_rich_text_storage_does_not_consume_text_budget_or_reference_evidence(IWorkDocumentKind kind) {
        using MemoryStream package = SelectedRichTextPackage(kind);
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, kind,
            new IWorkReadOptions { MaximumProjectedTextCharacters = 40 });
        (IWorkTable table, IReadOnlyList<IWorkSourceReferenceIssue> issues) = ReadSelectedRichTable(source, kind);
        Assert.Equal("Value", Assert.Single(table.Cells).RichText!.PlainText);
        Assert.Empty(issues);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages, 0)]
    [InlineData(IWorkDocumentKind.Pages, 1)]
    [InlineData(IWorkDocumentKind.Pages, 2)]
    [InlineData(IWorkDocumentKind.Pages, 3)]
    [InlineData(IWorkDocumentKind.Pages, 4)]
    [InlineData(IWorkDocumentKind.Numbers, 0)]
    [InlineData(IWorkDocumentKind.Numbers, 1)]
    [InlineData(IWorkDocumentKind.Numbers, 2)]
    [InlineData(IWorkDocumentKind.Numbers, 3)]
    [InlineData(IWorkDocumentKind.Numbers, 4)]
    [InlineData(IWorkDocumentKind.Keynote, 0)]
    [InlineData(IWorkDocumentKind.Keynote, 1)]
    [InlineData(IWorkDocumentKind.Keynote, 2)]
    [InlineData(IWorkDocumentKind.Keynote, 3)]
    [InlineData(IWorkDocumentKind.Keynote, 4)]
    public void Selected_table_dependencies_retain_reference_owner_paths(IWorkDocumentKind kind, int failure) {
        byte[] attributes = failure == 0 ? AttributeTable(8, AttributeEntry(0, ReferenceField(2, 999))) : Message();
        using MemoryStream package = SelectedRichTextPackage(kind, attributes, failure);
        IWorkConversionReport report = ConvertUnitReport(package, kind);
        (ulong owner, string path) = failure switch {
            0 => (15ul, "8/1[1]/2"),
            1 => (13ul, "3[1]/9"),
            2 => (14ul, "1"),
            3 => (11ul, "4/17"),
            _ => (10ul, "2")
        };
        AssertMissingReference(Assert.Single(report.SourceReferenceIssues), owner, path, 999);
        Assert.DoesNotContain(report.SourceUnits, unit => unit.Identity.RecordIdentifier is 18 or 999);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Selected_table_text_style_parents_are_assessed(IWorkDocumentKind kind) {
        using MemoryStream package = SelectedRichTextPackage(kind,
            AttributeTable(5, AttributeEntry(0, ReferenceField(2, 19))),
            additionalRecords: ArchiveRecord(19, 2022, Message(BytesField(1, Message(ReferenceField(3, 999))))));
        AssertMissingReference(Assert.Single(ConvertUnitReport(package, kind).SourceReferenceIssues), 19, "1/3", 999);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Rich_text_aliases_share_storage_and_physical_reference_accounting(IWorkDocumentKind kind) {
        using MemoryStream package = SelectedRichTextPackage(kind,
            AttributeTable(8, AttributeEntry(0, ReferenceField(2, 999))), aliases: true);
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, kind,
            new IWorkReadOptions { MaximumSourceReferenceIssues = 1 });
        (IWorkTable table, IReadOnlyList<IWorkSourceReferenceIssue> issues) = ReadSelectedRichTable(source, kind);
        AssertMissingReference(Assert.Single(issues), 15, "8/1[1]/2", 999);
        Assert.Equal(2, table.Cells.Count);
        Assert.Same(table.Cells[0].RichText, table.Cells[1].RichText);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Selected_rich_text_reference_limit_failures_propagate(IWorkDocumentKind kind) {
        using MemoryStream package = SelectedRichTextPackage(kind,
            Message(AttributeTable(8, AttributeEntry(0, ReferenceField(2, 999))),
                AttributeTable(11, AttributeEntry(0, ReferenceField(2, 998)))));
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, kind,
            new IWorkReadOptions { MaximumSourceReferenceIssues = 1 });
        Assert.Throws<InvalidDataException>(() => ReadSelectedRichTable(source, kind));
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Repeated_rich_catalog_keys_remain_rejected_after_a_third_declaration(IWorkDocumentKind kind) {
        using MemoryStream package = SelectedRichTextPackage(kind, duplicateKeys: true);
        (IWorkTable table, IReadOnlyList<IWorkSourceReferenceIssue> issues) = ReadSelectedRichTable(
            IWorkSourceDocument.Open(package, kind), kind);
        Assert.Equal(IWorkCellKind.Error, Assert.Single(table.Cells).Kind);
        Assert.Empty(issues);
    }

    private static (IWorkTable, IReadOnlyList<IWorkSourceReferenceIssue>) ReadSelectedRichTable(
        IWorkSourceDocument source, IWorkDocumentKind kind) {
        if (kind == IWorkDocumentKind.Pages) {
            IWorkPagesProjection pages = source.ReadPages();
            return (Assert.Single(pages.Tables), pages.SourceReferenceIssues);
        }
        if (kind == IWorkDocumentKind.Numbers) {
            IWorkNumbersProjection numbers = source.ReadNumbers();
            return (Assert.Single(Assert.Single(numbers.Sheets).Tables), numbers.SourceReferenceIssues);
        }
        IWorkKeynoteProjection keynote = source.ReadKeynote();
        return (Assert.Single(Assert.Single(keynote.Slides).Tables), keynote.SourceReferenceIssues);
    }

    private static MemoryStream SelectedRichTextPackage(IWorkDocumentKind kind, byte[]? attributes = null,
        int failure = -1, byte[]? additionalRecords = null, bool aliases = false, bool duplicateKeys = false,
        byte[]? catalogPayload = null, int wrongTypeRecord = -1, uint catalogType = 6005, bool includeCatalogKind = true) {
        byte[] cell = new byte[16]; cell[0] = 5; cell[1] = 9;
        WriteUInt32(cell, 8, 1u << 4); WriteUInt32(cell, 12, 1);
        byte[] secondCell = (byte[])cell.Clone(); WriteUInt32(secondCell, 12, 2);
        byte[] store = Message(BytesField(3, Message(BytesField(1,
            Message(VarintField(1, 0), ReferenceField(2, 12))))), ReferenceField(17, failure == 3 ? 999ul : 13ul));
        byte[] roots = kind switch {
            IWorkDocumentKind.Pages => Message(
                ArchiveRecord(1, 10000, Message(ReferenceField(4, 2)), new ulong[] { 2, 10 }),
                ArchiveRecord(2, 2001, Message(StringField(3, "Body")))),
            IWorkDocumentKind.Numbers => Message(ArchiveRecord(1, 1, Message(ReferenceField(1, 2))),
                ArchiveRecord(2, 2, Message(StringField(1, "Sheet"), ReferenceField(2, 10)))),
            _ => Message(ArchiveRecord(1, 1, Message(ReferenceField(2, 2))),
                ArchiveRecord(2, 2, KeynoteShow(Message(ReferenceField(2, 3)))),
                ArchiveRecord(3, 4, Message(ReferenceField(2, 4))),
                ArchiveRecord(4, 5, Message(ReferenceField(7, 10))))
        };
        byte[] selectedEntry = BytesField(3, Message(VarintField(1, 1), ReferenceField(9, failure == 1 ? 999ul : 14ul)));
        byte[] catalog = Message(selectedEntry,
            duplicateKeys ? Message(selectedEntry, selectedEntry)
                : BytesField(3, Message(VarintField(1, 2), ReferenceField(9, aliases ? 14ul : 17ul))));
        catalog = Message(includeCatalogKind ? VarintField(1, 8) : Message(), catalogPayload ?? catalog);
        return CreatePackage(("Index/Document.iwa", FrameIwa(Message(roots,
            ArchiveRecord(10, 6000, Message(BytesField(1, GeometryDrawable(72, 72, 120, 40)),
                ReferenceField(2, failure == 4 ? 999ul : 11ul))),
            ArchiveRecord(11, wrongTypeRecord == 11 ? 2021u : 6001u, Message(BytesField(4, store), VarintField(6, 1), VarintField(7, aliases ? 2ul : 1ul))),
            ArchiveRecord(12, 6002, Message(BytesField(5, Message(VarintField(1, 0),
                BytesField(6, aliases ? Message(cell, secondCell) : cell),
                BytesField(7, aliases ? new byte[] { 0, 0, 16, 0 } : new byte[] { 0, 0 }))))),
            ArchiveRecord(13, wrongTypeRecord == 13 ? 2021u : catalogType, catalog),
            ArchiveRecord(14, wrongTypeRecord == 14 ? 2021u : 6218u, Message(ReferenceField(1, failure == 2 ? 999ul : 15ul))),
            ArchiveRecord(15, wrongTypeRecord == 15 ? 2021u : 2001u, Message(StringField(3, "Value"), attributes ?? Message())),
            ArchiveRecord(17, wrongTypeRecord == 17 ? 2021u : 6218u, Message(ReferenceField(1, 18))),
            ArchiveRecord(18, 2001, Message(StringField(3, new string('X', 1000)),
                AttributeTable(8, AttributeEntry(0, ReferenceField(2, 997))))),
            additionalRecords ?? Message()))), ("preview.png", ValidPreviewPng()));
    }
}
