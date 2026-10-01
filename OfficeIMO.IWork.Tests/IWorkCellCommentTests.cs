using OfficeIMO.IWork;
using OfficeIMO.Excel;
using OfficeIMO.Reader.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Qualified_root_cell_comments_retain_text_author_time_and_comment_only_empty_cells(IWorkDocumentKind kind) {
        using var package = CommentPackage(kind, empty: true);
        IWorkTableCell cell = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package, kind), kind).Item1.Cells);
        Assert.Equal(IWorkCellKind.Empty, cell.Kind);
        IWorkCellComment comment = Assert.IsType<IWorkCellComment>(cell.Comment);
        Assert.Equal("Review this\nplease ", comment.Text);
        Assert.Equal("Reviewer", comment.Author);
        Assert.Equal(new DateTime(2001, 1, 1, 0, 0, 42, DateTimeKind.Utc), comment.CreationDateUtc);
        Assert.Equal(14ul, comment.SourceIdentity.RecordIdentifier);
        Assert.Equal(15ul, comment.SourceAuthorIdentity.RecordIdentifier);
        Assert.Equal(IWorkCellUnsupportedFeatures.None, cell.UnsupportedFeatures);
        package.Position = 0;
        if (kind == IWorkDocumentKind.Numbers) {
            using var result = IWorkSourceDocument.Open(package).ToExcelDocumentResult();
            Assert.False(result.IsVisualFallback);
            using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
            using var reopened = ExcelDocument.Load(saved);
            var destination = Assert.Single(reopened.Sheets[0].GetThreadedComments());
            Assert.Equal(comment.Text, destination.Text); Assert.Equal(comment.Author, destination.Author);
            Assert.Equal("A1", destination.CellReference);
            Assert.Equal(comment.CreationDateUtc, destination.Date);
        } else {
            using var source = CommentPackage(kind);
            Assert.True(ConvertUnitReport(source, kind, visual: false).IsPartialEditableReconstruction);
            using var strict = CommentPackage(kind);
            IWorkSourceDocument strictSource = IWorkSourceDocument.Open(strict, kind);
            if (kind == IWorkDocumentKind.Pages) {
                using var result = strictSource.ToWordDocumentResult(); Assert.True(result.IsVisualFallback);
                Assert.Contains(result.Report.Diagnostics, d => d.Code == "IWORK_PAGES_WORD_DESTINATION_UNSUPPORTED");
            } else {
                using var result = strictSource.ToPowerPointPresentationResult(); Assert.True(result.IsVisualFallback);
            }
        }
    }

    [Theory]
    [InlineData(0)] // Missing catalog target.
    [InlineData(1)] // Duplicate entry keys cannot choose an arbitrary root.
    [InlineData(2)] // Missing comment record.
    [InlineData(3)] // Missing author record.
    [InlineData(4)] // Replies are not flattened into a root-only comment.
    [InlineData(5)] // Invalid timestamp is not replaced with the conversion time.
    [InlineData(6)] // Duplicate text does not choose the last value.
    public void Unqualified_selected_comments_keep_values_and_precise_reference_or_declaration_evidence(int fault) {
        byte[] entry = Message(VarintField(1, 1), ReferenceField(10, fault == 2 ? 999ul : 14ul));
        byte[] comment = Message(StringField(1, "Review this\nplease "),
            BytesField(2, Message(DoubleField(1, fault == 5 ? double.NaN : 42d))),
            ReferenceField(3, fault == 3 ? 999ul : 15ul), fault == 4 ? ReferenceField(4, 999) : Message(),
            fault == 6 ? StringField(1, "different") : Message());
        using var package = TableDependencyPackage(IWorkDocumentKind.Numbers, ReferenceField(19, fault == 0 ? 999ul : 13ul),
            cellPayload: CommentCell(false), additionalRecords: Message(
                ArchiveRecord(13, 6005, Message(VarintField(1, 10), BytesField(3, entry), fault == 1 ? BytesField(3, entry) : Message())),
                ArchiveRecord(14, 3056, comment), ArchiveRecord(15, 212, StringField(1, "Reviewer"))));
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(package).ReadNumbers();
        IWorkTableCell cell = Assert.Single(Assert.Single(Assert.Single(projection.Sheets).Tables).Cells);
        Assert.Null(cell.Comment); Assert.Equal(42d, cell.Value);
        Assert.Equal(IWorkCellUnsupportedFeatures.Comment, cell.UnsupportedFeatures);
        Assert.False(projection.HasEditableContent);
        if (fault is 0 or 2 or 3) {
            var issue = Assert.Single(projection.SourceReferenceIssues);
            Assert.Equal(999ul, issue.TargetIdentifier);
            Assert.Equal(fault == 0 ? "4/19" : fault == 2 ? "3[1]/10" : "3", issue.FieldPath);
        } else {
            Assert.Empty(projection.SourceReferenceIssues);
            Assert.Contains(projection.SourceDeclarationIssues, issue => issue.FieldPath == (fault == 1 ? "3[2]/1" : fault == 4 ? "4" : fault == 5 ? "2" : "1"));
        }
    }

    [Fact]
    public void Comments_charge_source_wide_text_and_catalog_limits_for_repeated_uses() {
        using var characters = CommentPackage(IWorkDocumentKind.Numbers, repeatModel: true);
        Assert.Throws<InvalidDataException>(() => IWorkSourceDocument.Open(characters,
            new IWorkReadOptions { MaximumProjectedTextCharacters = 40 }).ReadNumbers());
        using var entries = CommentPackage(IWorkDocumentKind.Numbers, repeatModel: true);
        Assert.Throws<InvalidDataException>(() => IWorkSourceDocument.Open(entries,
            new IWorkReadOptions { MaximumTableCatalogEntries = 1 }).ReadNumbers());
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages, "pages")]
    [InlineData(IWorkDocumentKind.Numbers, "numbers")]
    [InlineData(IWorkDocumentKind.Keynote, "key")]
    public void Reader_retains_qualified_cell_comment_content_and_metadata_without_replacing_values(IWorkDocumentKind kind, string extension) {
        using var package = CommentPackage(kind);
        var reader = new OfficeIMO.Reader.OfficeDocumentReaderBuilder().AddIWorkHandler().Build();
        var result = reader.ReadDocument(package, "comments." + extension);
        Assert.Equal("42", Assert.Single(Assert.Single(result.Tables).Rows[0]));
        var block = Assert.Single(result.Blocks, b => b.Kind == "comment");
        Assert.Equal("Review this\nplease ", block.Text);
        Assert.Equal("A1", block.Location.A1Range);
        Assert.Equal(0, block.Location.TableIndex);
        Assert.Equal("table-cell-comment", block.Location.SourceBlockKind);
        var metadata = Assert.Single(result.Metadata, m => m.Category == "table.comment");
        Assert.Equal(block.Text, metadata.Value);
        Assert.Equal(block.Id, metadata.Location!.BlockAnchor);
        Assert.Equal("Reviewer", metadata.Attributes["author"]);
        Assert.Equal("2001-01-01T00:00:42.0000000Z", metadata.Attributes["creationDateUtc"]);
        Assert.Equal("14", metadata.SourceObjectId);
        Assert.Equal("15", metadata.Attributes["authorRecordIdentifier"]);
        Assert.Contains("Comment on", result.Markdown);
        Assert.Contains("Review this\nplease ", result.Markdown);
        Assert.DoesNotContain(result.Diagnostics, d => d.Code == "IWORK_READER_TABLE_COMMENTS_OMITTED");
    }

    [Theory]
    [InlineData(" ", "Reviewer")]
    [InlineData("Text", " Reviewer ")]
    [InlineData("Text", "")]
    public void Xlsx_comments_reject_silent_text_or_author_rewriting(string text, string author) {
        using var package = TableDependencyPackage(IWorkDocumentKind.Numbers, ReferenceField(19, 13),
            cellPayload: CommentCell(false), additionalRecords: Message(
                ArchiveRecord(13, 6005, Message(VarintField(1, 10), BytesField(3, Message(VarintField(1, 1), ReferenceField(10, 14))))),
                ArchiveRecord(14, 3056, Message(StringField(1, text), BytesField(2, Message(DoubleField(1, 42d))), ReferenceField(3, 15))),
                ArchiveRecord(15, 212, StringField(1, author))));
        using var result = IWorkSourceDocument.Open(package).ToExcelDocumentResult();
        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "IWORK_NUMBERS_EXCEL_DESTINATION_UNSUPPORTED");
    }

    [Fact]
    public void Xlsx_comments_do_not_merge_case_distinct_author_names() {
        byte[] first = CommentCell(false), second = CommentCell(false); WriteUInt32(second, second.Length - 4, 2);
        using var package = TableDependencyPackage(IWorkDocumentKind.Numbers, ReferenceField(19, 13), columns: 2,
            tilePayload: BytesField(5, Message(VarintField(1, 0), BytesField(6, Message(first, second)),
                BytesField(7, new byte[] { 0, 0, (byte)first.Length, 0 }))), additionalRecords: Message(
                ArchiveRecord(13, 6005, Message(VarintField(1, 10),
                    BytesField(3, Message(VarintField(1, 1), ReferenceField(10, 14))),
                    BytesField(3, Message(VarintField(1, 2), ReferenceField(10, 16))))),
                ArchiveRecord(14, 3056, Message(StringField(1, "First"), BytesField(2, Message(DoubleField(1, 42d))), ReferenceField(3, 15))),
                ArchiveRecord(15, 212, StringField(1, "Reviewer")),
                ArchiveRecord(16, 3056, Message(StringField(1, "Second"), BytesField(2, Message(DoubleField(1, 43d))), ReferenceField(3, 17))),
                ArchiveRecord(17, 212, StringField(1, "reviewer"))));
        using var result = IWorkSourceDocument.Open(package).ToExcelDocumentResult();
        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Report.Diagnostics, d => d.Message.Contains("differ only by case", StringComparison.Ordinal));
    }

    private static MemoryStream CommentPackage(IWorkDocumentKind kind, bool empty = false, bool repeatModel = false) =>
        TableDependencyPackage(kind, ReferenceField(19, 13), repeatModel: repeatModel, cellPayload: CommentCell(empty),
            additionalRecords: Message(
                ArchiveRecord(13, 6005, Message(VarintField(1, 10), BytesField(3, Message(VarintField(1, 1), ReferenceField(10, 14))))),
                ArchiveRecord(14, 3056, Message(StringField(1, "Review this\nplease "), BytesField(2, Message(DoubleField(1, 42d))), ReferenceField(3, 15))),
                ArchiveRecord(15, 212, StringField(1, "Reviewer"))));

    private static byte[] CommentCell(bool empty) {
        byte[] cell = FeatureCell(empty, 1u << 19);
        WriteUInt32(cell, cell.Length - 4, 1);
        return cell;
    }
}
