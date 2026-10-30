using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Publisher.Pdf;
using OfficeIMO.TestSupport;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherTableEditingTests {
    [Theory]
    [InlineData("Sample.pub")]
    [InlineData("Sample_2010.pub")]
    [InlineData("NativeComparison/TableLabel/native-tablelabel.pub")]
    public void Different_length_cell_edits_keep_later_cells_complete_story_styles_and_detached_snapshots(string fixture) {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture(fixture));
        PublisherTable original = Assert.Single(document.Pages.SelectMany(page => page.Tables));
        Assert.True(original.CanEditCellText, original.CellTextEditLimitation);
        Assert.True(original.CanEditBounds, original.BoundsEditLimitation);
        PublisherTextStory oldStory = document.TextStories.Single(story => story.Id == original.StoryId);
        PublisherPage oldPage = document.Pages.Single(page => page.Id == original.PageId);
        string oldSvg = OfficeDrawingSvgExporter.ToSvg(oldPage.CreateDrawing());
        var sourceReport = document.SourceReadReport;
        string[] neighbors = original.Cells.Skip(1).Select(cell => cell.Text).ToArray();
        var label = OfficeTextParagraphLabel.AtPosition(new OfficeRichTextRun("•", 9, OfficeColor.Black, fontFamily: "Arial"), 0);
        var tabs = new OfficeTextTabStops(new[] { new OfficeTextTabStop(30) });
        var first = new OfficeRichTextParagraph(new[] {
            new OfficeRichTextRun("A longer replacement", 9, OfficeColor.FromRgb(120, 20, 30), bold: true, fontFamily: "Arial")
        }, label).WithTabStops(tabs);
        var second = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("Next\tvalue", 9, OfficeColor.Black, italic: true, fontFamily: "Arial") }).WithTabStops(tabs);
        var input = new List<OfficeRichTextParagraph> { first, second };
        document.ReplaceTableCellParagraphs(original.Id, 0, 0, input);
        input.Clear();
        PublisherTable current = Table(document, original.Id);
        Assert.Equal(neighbors, current.Cells.Skip(1).Select(cell => cell.Text));
        Assert.Equal("• A longer replacement\nNext\tvalue", current.Cells[0].Text);
        Assert.True(current.Cells[0].Paragraphs[0].Runs[0].Bold);
        Assert.True(current.Cells[0].Paragraphs[1].Runs[0].Italic);
        Assert.Equal(30, current.Cells[0].Paragraphs[1].TabStops!.Stops[0].Position);
        PublisherTextStory story = document.TextStories.Single(item => item.Id == original.StoryId);
        Assert.StartsWith("A longer replacement\nNext\tvalue\n", story.Text);
        Assert.Equal(oldStory.Text.Substring(oldStory.Text.IndexOf('\n') + 1), story.Text.Substring(story.Text.IndexOf("value", StringComparison.Ordinal) + 6));
        Assert.Equal(oldStory.Text.EndsWith("\n", StringComparison.Ordinal), story.Text.EndsWith("\n", StringComparison.Ordinal));
        string editedSvg = document.ToSvg(document.Pages.ToList().FindIndex(page => page.Id == original.PageId));
        Assert.Contains("A longer replacement", VisibleSvgText(editedSvg));
        var pdfOptions = new PdfOptions();
        if (fixture == "NativeComparison/TableLabel/native-tablelabel.pub") pdfOptions.UseTextFallbacks();
        PdfReadDocument pdf = PdfReadDocument.Open(document.ToPdfBytes(pdfOptions));
        PdfReadPage pdfPage = pdf.Pages[document.Pages.ToList().FindIndex(page => page.Id == current.PageId)];
        PublisherTableCell editedCell = current.Cells[0];
        double left = current.X + editedCell.X, right = left + editedCell.Width;
        double top = pdfPage.GetPageSize().Height - current.Y - editedCell.Y, bottom = top - editedCell.Height;
        // Whole-page reading order interleaves neighboring cells with wrapped
        // lines. Verify the complete visible phrase inside its actual cell.
        string cellPdfText = NormalizeWhitespace(string.Join(" ", pdfPage.GetTextSpans()
            .Where(span => span.IsVisible && span.X >= left - .1 && span.X < right - .1 && span.Y >= bottom - .1 && span.Y <= top + .1)
            .OrderByDescending(span => span.Y).ThenBy(span => span.X).Select(span => span.Text)));
        Assert.Contains("A longer replacement", cellPdfText);
        Assert.DoesNotContain(original.Cells[1].Text, cellPdfText);
        Assert.Equal(oldSvg, OfficeDrawingSvgExporter.ToSvg(oldPage.CreateDrawing()));
        Assert.Equal(neighbors, original.Cells.Skip(1).Select(cell => cell.Text));
        Assert.DoesNotContain("A longer replacement", oldStory.Text);
        Assert.Same(sourceReport, document.SourceReadReport);
        Assert.Equal(original.StoryId, current.StoryId); Assert.Equal(original.PageId, current.PageId);
    }

    [Fact]
    public void Clearing_and_refilling_first_and_final_cells_preserves_neighbor_boundaries_and_terminators() {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("Sample.pub"));
        PublisherTable table = Table(document, 299);
        string[] neighbors = table.Cells.Skip(1).Take(table.Cells.Count - 2).Select(cell => cell.Text).ToArray();
        PublisherTextStory originalStory = document.TextStories.Single(story => story.Id == table.StoryId);
        foreach (PublisherTableCell cell in new[] { table.Cells[0], table.Cells.Last() })
            document.ReplaceTableCellParagraphs(table.Id, cell.RowIndex, cell.ColumnIndex, Array.Empty<OfficeRichTextParagraph>());
        Assert.Empty(Table(document, 299).Cells[0].Text); Assert.Empty(Table(document, 299).Cells.Last().Text);
        Assert.Equal(neighbors, Table(document, 299).Cells.Skip(1).Take(neighbors.Length).Select(cell => cell.Text));
        Assert.StartsWith("\n", document.TextStories.Single(story => story.Id == table.StoryId).Text);
        document.ReplaceTableCellParagraphs(299, 0, 0, Paragraphs("Refilled first"));
        document.ReplaceTableCellParagraphs(299, table.Cells.Last().RowIndex, table.Cells.Last().ColumnIndex, Paragraphs("Refilled last"));
        PublisherTextStory current = document.TextStories.Single(story => story.Id == table.StoryId);
        Assert.StartsWith("Refilled first\n", current.Text); Assert.Contains("Refilled last", current.Text);
        Assert.Equal(originalStory.Text.EndsWith("\n", StringComparison.Ordinal), current.Text.EndsWith("\n", StringComparison.Ordinal));
        Assert.Equal(neighbors, Table(document, 299).Cells.Skip(1).Take(neighbors.Length).Select(cell => cell.Text));
    }

    [Fact]
    public void Merged_anchor_edit_preserves_span_and_other_cells_and_rejects_covered_coordinates() {
        byte[] bytes = PublisherTableFixture.MergeFirstRow(File.ReadAllBytes(PublisherNativeTests.Fixture("Sample.pub")));
        PublisherDocument document = PublisherDocument.Load(bytes);
        PublisherTable original = Table(document, 299);
        document.ReplaceTableCellParagraphs(299, 0, 0, Paragraphs("One merged owner", "Second paragraph"));
        PublisherTable edited = Table(document, 299);
        Assert.Equal(2, edited.Cells[0].ColumnSpan); Assert.Equal(original.Cells[0].Width, edited.Cells[0].Width);
        Assert.Equal("One merged owner\nSecond paragraph", edited.Cells[0].Text);
        Assert.Equal(original.Cells.Skip(1).Select(cell => cell.Text), edited.Cells.Skip(1).Select(cell => cell.Text));
        var pages = document.Pages;
        Assert.Throws<ArgumentException>(() => document.ReplaceTableCellParagraphs(299, 0, 1, Paragraphs("covered")));
        Assert.Same(pages, document.Pages);
    }

    [Fact]
    public void Compatible_shared_tables_update_one_complete_story_without_detaching_occurrences() {
        PublisherDocument document = PublisherTableEditingFixture.Load("shared");
        PublisherTable original = Table(document, 299);
        int storyCount = document.TextStories.Count;
        document.ReplaceTableCellParagraphs(900, 0, 0, Paragraphs("Shared replacement"));
        Assert.Equal("Shared replacement", Table(document, 299).Cells[0].Text);
        Assert.Equal("Shared replacement", Table(document, 900).Cells[0].Text);
        Assert.Equal(original.StoryId, Table(document, 900).StoryId);
        Assert.Equal(storyCount, document.TextStories.Count);
        Assert.Single(document.TextStories, story => story.Text.Contains("Shared replacement"));
        Assert.NotEqual("Shared replacement", original.Cells[0].Text);
    }

    [Theory]
    [InlineData("conflict")]
    [InlineData("ordinary-frame")]
    [InlineData("partition-conflict")]
    [InlineData("unresolved")]
    public void Unqualified_cell_ownership_is_visible_and_rejects_edits_atomically(string variant) {
        PublisherDocument document = PublisherTableEditingFixture.Load(variant);
        PublisherTable table = Table(document, 299);
        Assert.False(table.CanEditCellText); Assert.NotNull(table.CellTextEditLimitation);
        var pages = document.Pages; var stories = document.TextStories; var report = document.ReadReport;
        Assert.Throws<NotSupportedException>(() => document.ReplaceTableCellParagraphs(299, 0, 0, Paragraphs("unsafe")));
        Assert.Same(pages, document.Pages); Assert.Same(stories, document.TextStories); Assert.Same(report, document.ReadReport);
        Assert.Contains(document.ReadReport.FidelityDiagnostics, diagnostic => diagnostic.Code == "PUB_TABLE_CELL_TEXT_EDIT_UNSUPPORTED"
            && diagnostic.Location == "Contents/object/299" && diagnostic.LossKind == OfficeConversionLossKind.None);
    }

    [Theory]
    [InlineData("shared-missing-grid")]
    [InlineData("shared-unresolved-owner")]
    public void Unresolved_native_table_occurrences_block_shared_story_edits_without_losing_recovered_content(string variant) {
        PublisherDocument document = PublisherTableEditingFixture.Load(variant);
        PublisherTable table = Assert.Single(document.Pages.Concat(document.MasterPages).SelectMany(page => page.Tables));
        Assert.Equal(299U, table.Id); Assert.True(table.HasTextMapping);
        Assert.Equal(6, table.Cells.Count); Assert.Equal("Table on page 2", table.Cells[0].Text);
        PublisherTextStory story = document.TextStories.Single(item => item.Id == table.StoryId);
        Assert.Contains("Bottom Right", story.Text);
        var pages = document.Pages; var stories = document.TextStories; var report = document.ReadReport;
        Assert.False(table.CanEditCellText); Assert.NotNull(table.CellTextEditLimitation);
        Assert.Throws<NotSupportedException>(() => document.ReplaceTableCellParagraphs(table.Id, 0, 0, Paragraphs("unsafe shared change")));
        Assert.Same(pages, document.Pages); Assert.Same(stories, document.TextStories); Assert.Same(report, document.ReadReport);
        Assert.Same(story, document.TextStories.Single(item => item.Id == table.StoryId));
        Assert.Equal("Table on page 2", Table(document, table.Id).Cells[0].Text);
        Assert.Contains(report.FidelityDiagnostics, diagnostic => diagnostic.Code == "PUB_TABLE_CELL_TEXT_EDIT_UNSUPPORTED"
            && diagnostic.Location == "Contents/object/299" && diagnostic.LossKind == OfficeConversionLossKind.None);

        PublisherShape added = document.AddShape(table.PageId, OfficeShape.Rectangle(10, 10), 0, 0);
        document.SetShapePaint(added.Id, OfficeColor.Black, null);
        document.SetTableBounds(table.Id, table.X + 1, table.Y, table.Width, table.Height);
        Assert.False(Table(document, table.Id).CanEditCellText);
        PublisherTable authored = document.AddTable(table.PageId, 0, 0, new[] { 40D }, new[] { 30D });
        Assert.True(authored.CanEditCellText, authored.CellTextEditLimitation);
        document.ReplaceTableCellParagraphs(authored.Id, 0, 0, Paragraphs("Independent table story"));
        Assert.Equal("Independent table story", Table(document, authored.Id).Cells[0].Text);
        Assert.False(Table(document, table.Id).CanEditCellText);
        pages = document.Pages; stories = document.TextStories; report = document.ReadReport;
        Assert.Throws<NotSupportedException>(() => document.ReplaceTableCellParagraphs(table.Id, 0, 0, Paragraphs("unsafe after model copies")));
        Assert.Same(pages, document.Pages); Assert.Same(stories, document.TextStories); Assert.Same(report, document.ReadReport);
        Assert.Same(story, document.TextStories.Single(item => item.Id == table.StoryId));
        Assert.Equal("Table on page 2", Table(document, table.Id).Cells[0].Text);
    }

    [Fact]
    public void Bounds_edits_scale_tracks_preserve_text_and_update_current_text_wrap_exclusion() {
        PublisherDocument document = PublisherTableEditingFixture.Load(withWrap: true);
        PublisherTable original = Table(document, 299);
        Assert.True(original.CanEditBounds, original.BoundsEditLimitation);
        document.SetTableBounds(299, 60, 60, original.Width, original.Height);
        OfficeDrawingRichText[] before = WrappedRegions(document);
        document.SetTableBounds(299, 500, 500, original.Width * 1.5, original.Height * 2);
        OfficeDrawingRichText[] after = WrappedRegions(document);
        Assert.NotEmpty(before); Assert.NotEmpty(after);
        Assert.NotEqual(string.Join(";", before.Select(region => $"{region.X},{region.Y},{region.Width},{region.Height}")),
            string.Join(";", after.Select(region => $"{region.X},{region.Y},{region.Width},{region.Height}")));
        PublisherTable current = Table(document, 299);
        Assert.Equal((500D, 500D, original.Width * 1.5, original.Height * 2), (current.X, current.Y, current.Width, current.Height));
        Assert.Equal(new OfficePoint(500, 500), current.PageTransform.TransformPoint(new OfficePoint(0, 0)));
        for (int i = 0; i < current.ColumnCount; i++) Assert.Equal(original.ColumnWidths[i] * 1.5, current.ColumnWidths[i], 6);
        for (int i = 0; i < current.RowCount; i++) Assert.Equal(original.RowHeights[i] * 2, current.RowHeights[i], 6);
        Assert.Equal(original.Cells.Select(cell => cell.Text), current.Cells.Select(cell => cell.Text));
        Assert.DoesNotContain(document.Pages.SelectMany(page => page.Shapes), shape => shape.Id == 299);
        Assert.Throws<ArgumentOutOfRangeException>(() => document.SetShapeBounds(299, 10, 20, 30, 40));
    }

    [Theory]
    [InlineData("hidden")]
    [InlineData("rotation")]
    [InlineData("gradient")]
    [InlineData("no-anchor")]
    [InlineData("unknown-preset")]
    [InlineData("custom-path")]
    public void Unsupported_placement_retains_grid_text_and_explicit_bounds_limitation(string variant) {
        PublisherDocument document = PublisherTableEditingFixture.Load(variant);
        PublisherTable table = Table(document, 299);
        Assert.False(table.CanEditBounds); Assert.NotNull(table.BoundsEditLimitation);
        Assert.Equal(6, table.Cells.Count); Assert.True(table.HasTextMapping);
        Assert.Equal("Table on page 2", table.Cells[0].Text);
        var pages = document.Pages;
        Assert.Throws<NotSupportedException>(() => document.SetTableBounds(299, 10, 20, 30, 40)); Assert.Same(pages, document.Pages);
        if (variant is "hidden" or "no-anchor") Assert.DoesNotContain("Table on page 2", document.ToSvg(1));
        if (variant == "no-anchor") Assert.False(table.HasPlacement);
        Assert.Contains(document.ReadReport.FidelityDiagnostics, diagnostic => diagnostic.Code == "PUB_TABLE_BOUNDS_EDIT_UNSUPPORTED");
    }

    [Fact]
    public void Host_shape_without_a_native_table_grid_remains_inspection_only() {
        PublisherDocument document = PublisherDocument.Load(PublisherDrawingFixture.Mutate(new Dictionary<ushort, uint>(), 201));
        PublisherShape shape = document.Pages.SelectMany(page => page.Shapes).Single(item => item.Id == 293);
        Assert.False(shape.CanEdit); Assert.NotNull(shape.EditLimitation);
        var pages = document.Pages;
        Assert.Throws<NotSupportedException>(() => document.SetShapeBounds(shape.Id, 10, 20, 30, 40));
        Assert.Same(pages, document.Pages);
    }

    [Fact]
    public void Table_creation_uses_shared_identity_order_current_counts_and_master_inheritance() {
        PublisherDocument document = PublisherShapeEditingFixture.Create();
        int source = document.ReadReport.SourceObjectCount, records = document.ReadReport.InspectedRecordCount, stories = document.TextStories.Count;
        PublisherPage oldPage = document.Pages[0];
        var columns = new List<double> { 40, 60 }; var rows = new List<double> { 24, 36 };
        PublisherTable first = document.AddTable(10, 20, 120, columns, rows);
        columns[0] = 1; rows.Clear();
        Assert.Equal(new[] { 40D, 60D }, first.ColumnWidths); Assert.Equal(new[] { 24D, 36D }, first.RowHeights);
        Assert.Equal(4, first.Cells.Count); Assert.All(first.Cells, cell => Assert.Empty(cell.Text));
        Assert.True(first.CanEditBounds); Assert.True(first.CanEditCellText);
        PublisherShape shape = document.AddShape(10, OfficeShape.Ellipse(20, 20), 240, 240);
        PublisherTable master = document.AddTable(20, 20, 220, new[] { 80D }, new[] { 40D });
        Assert.True(first.Id < first.StoryId && first.StoryId < shape.Id && shape.Id < master.Id && master.Id < master.StoryId);
        document.ReplaceTableCellParagraphs(first.Id, 0, 0, Paragraphs("Page cell"));
        document.ReplaceTableCellParagraphs(master.Id, 0, 0, Paragraphs("Master cell"));
        Assert.Equal(source + 3, document.ReadReport.CurrentObjectCount); Assert.Equal(source + 3, document.ReadReport.ProjectedObjectCount);
        Assert.Equal(records, document.ReadReport.InspectedRecordCount); Assert.Equal(source, document.ReadReport.SourceObjectCount);
        Assert.Equal(stories + 2, document.TextStories.Count);
        Assert.All(document.Pages, page => Assert.Contains("Master cell", document.ToSvg(document.Pages.ToList().IndexOf(page))));
        Assert.All(document.Pages, page => Assert.DoesNotContain(page.Tables, table => table.Id == master.Id));
        Assert.Empty(oldPage.Tables); Assert.Empty(first.Cells[0].Text);
        Assert.Equal(2, PdfReadDocument.Open(document.ToPdfBytes()).Pages.Count);
        Assert.DoesNotContain(document.ReadReport.FidelityDiagnostics, diagnostic => diagnostic.Location == "Contents/object/" + first.Id
            && diagnostic.Code == "PUB_TABLE_CELL_STYLE_UNASSESSED");
    }

    [Fact]
    public void Table_story_and_continuation_creation_interleave_without_identity_collisions() {
        PublisherDocument document = PublisherFrameCreationFixture.Load().Document;
        PublisherTextFrame tail = document.Pages.SelectMany(page => page.TextFrames).Single();
        PublisherTable table = document.AddTable(tail.PageId, 10, 10, new[] { 50D }, new[] { 30D });
        PublisherTextFrame next = document.AddLinkedTextFrame(tail.Id, tail.PageId, 10, 60, 50, 30);
        PublisherShape shape = document.AddShape(tail.PageId, OfficeShape.Rectangle(10, 10), 10, 110);
        Assert.Equal(4, new[] { table.Id, table.StoryId!.Value, next.Id, shape.Id }.Distinct().Count());
        Assert.Equal(table.Id + 1, table.StoryId.Value); Assert.Equal(table.StoryId.Value + 1, next.Id); Assert.Equal(next.Id + 1, shape.Id);
        document.ReplaceStoryParagraphs(tail.StoryId, Paragraphs("Ordinary story survives"));
        document.SetTextFrameBounds(next.Id, 20, 60, 70, 30);
        document.SetShapePaint(shape.Id, OfficeColor.Black, null);
        Assert.Equal(table.Id, Assert.Single(document.Pages.SelectMany(page => page.Tables)).Id);
        Assert.Empty(Table(document, table.Id).Cells[0].Text);
    }

    [Fact]
    public void Invalid_and_cancelled_edits_preserve_current_state_and_failed_creation_reuses_identifiers() {
        PublisherDocument document = PublisherShapeEditingFixture.Create();
        PublisherDocument control = PublisherShapeEditingFixture.Create();
        var pages = document.Pages; var stories = document.TextStories; var report = document.ReadReport;
        Assert.Throws<ArgumentOutOfRangeException>(() => document.AddTable(10, 0, 0, new[] { double.NaN }, new[] { 20D }));
        Assert.Throws<ArgumentException>(() => document.AddTable(10, 0, 0, Array.Empty<double>(), new[] { 20D }));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.AddTable(uint.MaxValue, 0, 0, new[] { 20D }, new[] { 20D }));
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        Assert.Throws<OperationCanceledException>(() => document.AddTable(10, 0, 0, new[] { 20D }, new[] { 20D }, cancelled.Token));
        Assert.Same(pages, document.Pages); Assert.Same(stories, document.TextStories); Assert.Same(report, document.ReadReport);
        PublisherTable actual = document.AddTable(10, 20, 120, new[] { 50D }, new[] { 30D });
        PublisherTable expected = control.AddTable(10, 20, 120, new[] { 50D }, new[] { 30D });
        Assert.Equal(expected.Id, actual.Id); Assert.Equal(expected.StoryId, actual.StoryId);
        pages = document.Pages; stories = document.TextStories; report = document.ReadReport;
        Assert.Throws<ArgumentException>(() => document.ReplaceTableCellParagraphs(actual.Id, 0, 0, Paragraphs("\uD800")));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.ReplaceTableCellParagraphs(actual.Id, 1, 0, Paragraphs("outside")));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.SetTableBounds(actual.Id, 0, 0, 0, 20));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.SetTableBounds(actual.Id, double.MaxValue, 0, 20, 20));
        Assert.Throws<OperationCanceledException>(() => document.ReplaceTableCellParagraphs(actual.Id, 0, 0, Paragraphs("cancelled"), cancelled.Token));
        Assert.Throws<OperationCanceledException>(() => document.SetTableBounds(actual.Id, 0, 0, 20, 20, cancelled.Token));
        Assert.Same(pages, document.Pages); Assert.Same(stories, document.TextStories); Assert.Same(report, document.ReadReport);
    }

    [Fact]
    public void Table_creation_and_text_enforce_model_and_projection_budgets_without_partial_commit() {
        PublisherDocument document = PublisherShapeEditingFixture.Create(options: new PublisherReadOptions {
            Limits = new OfficeLegacyImportLimits { MaxItems = 30, MaxTextCharacters = 100 }
        });
        PublisherTable first = document.AddTable(10, 20, 120, new[] { 80D }, new[] { 30D });
        var pages = document.Pages; var stories = document.TextStories; var report = document.ReadReport;
        Assert.Throws<InvalidDataException>(() => document.AddTable(10, 0, 0, Enumerable.Repeat(10D, 8).ToArray(), Enumerable.Repeat(10D, 8).ToArray()));
        // Story admission fits 100 characters; repeated master scene composition does not.
        PublisherTable master = document.AddTable(20, 10, 220, new[] { 150D }, new[] { 60D });
        pages = document.Pages; stories = document.TextStories; report = document.ReadReport;
        Assert.Throws<InvalidDataException>(() => document.ReplaceTableCellParagraphs(master.Id, 0, 0, Paragraphs(new string('x', 50))));
        Assert.Same(pages, document.Pages); Assert.Same(stories, document.TextStories); Assert.Same(report, document.ReadReport);
        Assert.Throws<ArgumentException>(() => document.ReplaceTableCellParagraphs(first.Id, 0, 0, Paragraphs(new string('x', 101))));
        Assert.Empty(Table(document, first.Id).Cells[0].Text);
        // This grid fits its own budget but exceeds cumulative table-model work.
        Assert.Throws<InvalidDataException>(() => document.AddTable(10, 0, 0, Enumerable.Repeat(10D, 4).ToArray(), Enumerable.Repeat(10D, 4).ToArray()));
        Assert.Same(pages, document.Pages); Assert.Same(stories, document.TextStories);
        PublisherTable afterFailure = document.AddTable(10, 200, 200, new[] { 20D }, new[] { 20D });
        Assert.Equal(master.StoryId!.Value + 1, afterFailure.Id);
        PublisherDocument exhausted = PublisherShapeEditingFixture.Create(exhaustedIds: true);
        Assert.Throws<InvalidOperationException>(() => exhausted.AddTable(10, 0, 0, new[] { 20D }, new[] { 20D }));
    }

    [Fact]
    public void Table_creation_requires_two_identifiers_object_capacity_and_qualified_page_registration() {
        PublisherDocument oneId = PublisherTableEditingFixture.Load("one-id");
        uint pageId = Table(oneId, 299).PageId;
        var pages = oneId.Pages;
        Assert.Throws<InvalidOperationException>(() => oneId.AddTable(pageId, 0, 0, new[] { 20D }, new[] { 20D }));
        Assert.Same(pages, oneId.Pages);
        PublisherShape last = oneId.AddShape(pageId, OfficeShape.Rectangle(10, 10), 0, 0);
        Assert.Equal(uint.MaxValue, last.Id);

        PublisherDocument full = PublisherShapeEditingFixture.Create(hiddenObjects: 125, options: new PublisherReadOptions {
            Limits = new OfficeLegacyImportLimits { MaxItems = 128 }
        });
        Assert.Equal(128, full.ReadReport.CurrentObjectCount);
        pages = full.Pages;
        Assert.Throws<InvalidDataException>(() => full.AddTable(10, 0, 0, new[] { 20D }, new[] { 20D })); Assert.Same(pages, full.Pages);

        PublisherDocument unqualified = PublisherShapeEditingFixture.Create("unqualified-master");
        var masters = unqualified.MasterPages;
        Assert.Throws<NotSupportedException>(() => unqualified.AddTable(20, 0, 0, new[] { 20D }, new[] { 20D }));
        Assert.Same(masters, unqualified.MasterPages);
        Assert.True(unqualified.AddTable(10, 0, 0, new[] { 20D }, new[] { 20D }).CanEditBounds);
    }

    private static PublisherTable Table(PublisherDocument document, uint id) => document.Pages.Concat(document.MasterPages)
        .SelectMany(page => page.Tables).Single(table => table.Id == id);
    private static OfficeRichTextParagraph[] Paragraphs(params string[] values) => values.Select(value =>
        new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(value, 9, OfficeColor.Black, fontFamily: "Arial") })).ToArray();
    private static OfficeDrawingRichText[] WrappedRegions(PublisherDocument document) => PublisherNativeTests.Elements(document.Pages[1].CreateDrawing())
        .OfType<OfficeDrawingRichText>().Where(region => region.SourceElementIds?.Contains("publisher-object-902") == true).ToArray();
    private static string VisibleSvgText(string svg) => NormalizeWhitespace(string.Join(" ", XDocument.Parse(svg)
        .Descendants(XName.Get("text", "http://www.w3.org/2000/svg")).SelectMany(element => element.DescendantNodes().OfType<XText>()).Select(node => node.Value)));
    private static string NormalizeWhitespace(string value) => Regex.Replace(value, @"\s+", " ").Trim();
}
