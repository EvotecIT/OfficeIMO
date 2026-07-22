using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.Reader;
using OfficeIMO.Word;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests;

public sealed class ReaderOfficeRichMappingTests {
    [Fact]
    public void DocumentReader_WordRichMapping_UsesInspectionBlocksTablesLinksAndProperties() {
        using var stream = new MemoryStream();
        using (WordDocument document = WordDocument.Create(stream)) {
            document.BuiltinDocumentProperties.Title = "Rich policy";
            document.BuiltinDocumentProperties.Creator = "OfficeIMO";
            document.AddParagraph("Policy").Style = WordParagraphStyles.Heading1;
            document.AddParagraph("Read ").AddHyperLink("the policy", new Uri("https://example.test/policy"));
            WordParagraph noteReference = document.AddParagraph("Supporting references");
            WordParagraph footnoteReference = noteReference.AddFootNote("Footnote detail ");
            footnoteReference.FootNote!.Paragraphs![1].AddHyperLink(
                "footnote source",
                new Uri("https://example.test/footnote"));
            noteReference.AddEndNote("Endnote detail");
            WordTable table = document.AddTable(2, 2);
            table.Title = "Inventory";
            table.RepeatAsHeaderRowAtTheTopOfEachPage = true;
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Name";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "Qty";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "Bandage";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "4";
            document.Save();
        }
        stream.Position = 0;

        OfficeDocumentReadResult result = OfficeIMO.Reader.Tests.ReaderTestReaders.Word().ReadDocument(stream, "policy.docx");

        Assert.Equal("Rich policy", result.Source.Title);
        Assert.Equal("OfficeIMO", result.Source.Author);
        Assert.Contains(result.Blocks, block => block.Kind == "heading" && block.Level == 1 && block.Location.HeadingPath == "Policy");
        Assert.Contains(result.Blocks, block => block.Kind == "paragraph" && block.Location.HeadingPath == "Policy");
        ReaderTable mapped = Assert.Single(result.Tables);
        Assert.Equal("Inventory", mapped.Title);
        Assert.Equal("Bandage", mapped.Rows[0][0]);
        Assert.Equal("https://example.test/policy", Assert.Single(result.Links, link => link.Uri == "https://example.test/policy").Uri);
        OfficeDocumentLink noteLink = Assert.Single(result.Links, link => link.Uri == "https://example.test/footnote");
        Assert.StartsWith("word-footnote-", noteLink.Location.BlockAnchor, StringComparison.Ordinal);
        Assert.Contains(result.Blocks, block => block.Kind == "footnote" && block.Text.Contains("Footnote detail", StringComparison.Ordinal));
        Assert.Contains(result.Blocks, block => block.Kind == "endnote" && block.Text.Contains("Endnote detail", StringComparison.Ordinal));
        Assert.Contains("officeimo.word.inspection-snapshot", result.CapabilitiesUsed);

        stream.Position = 0;
        OfficeDocumentReadResult withoutNotes = OfficeIMO.Reader.Tests.ReaderTestReaders.Word(includeFootnotes: false).ReadDocument(stream, "policy.docx");
        Assert.DoesNotContain(withoutNotes.Blocks, block => block.Kind == "footnote" || block.Kind == "endnote");
    }

    [Fact]
    public void DocumentReader_WordRichMapping_AppliesTableRowLimitToBlocksAndTables() {
        using var stream = new MemoryStream();
        using (WordDocument document = WordDocument.Create(stream)) {
            WordTable table = document.AddTable(4, 2);
            table.RepeatAsHeaderRowAtTheTopOfEachPage = true;
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Name";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "Qty";
            for (int row = 1; row < 4; row++) {
                table.Rows[row].Cells[0].Paragraphs[0].Text = "Row " + row;
                table.Rows[row].Cells[1].Paragraphs[0].Text = row.ToString();
            }
            document.Save();
        }
        stream.Position = 0;

        OfficeDocumentReadResult result = OfficeIMO.Reader.Tests.ReaderTestReaders.Word().ReadDocument(
            stream,
            "bounded.docx",
            new ReaderOptions { MaxTableRows = 1 });

        ReaderTable tableResult = Assert.Single(result.Tables);
        Assert.Single(tableResult.Rows);
        Assert.True(tableResult.Truncated);
        OfficeDocumentBlock tableBlock = Assert.Single(result.Blocks, block => block.Kind == "table");
        Assert.Contains("Row 1", tableBlock.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("Row 2", tableBlock.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("Row 3", tableBlock.Text, StringComparison.Ordinal);
        ReaderChunk tableChunk = Assert.Single(result.Chunks, chunk => chunk.Tables?.Count == 1);
        ReaderTable chunkTable = Assert.Single(tableChunk.Tables!);
        Assert.Equal(3, chunkTable.TotalRowCount);
        Assert.Single(chunkTable.Rows);
        Assert.True(chunkTable.Truncated);
        Assert.Contains("Row 1", tableChunk.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("Row 2", tableChunk.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("Row 3", tableChunk.Markdown!, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReader_WordChunksSplitTextAndMarkdownTogether() {
        string plainBoundaryRun = new string('p', 255) + "😀" + new string('q', 80);
        string formattedBoundaryRun = new string('a', 251) + "🧭" + new string('b', 180);
        string hyperlinkText = new string('λ', 420) + "🚀" + new string('ω', 180);
        string finalRun = new string('z', 360);
        string paragraphText = plainBoundaryRun + formattedBoundaryRun + hyperlinkText + finalRun;
        const string hyperlinkUri = "https://example.test/rich-chunk";
        using var stream = new MemoryStream();
        using (WordDocument document = WordDocument.Create(stream)) {
            WordParagraph paragraph = document.AddParagraph();
            paragraph.AddText(plainBoundaryRun);
            paragraph.AddText(formattedBoundaryRun).Bold = true;
            WordParagraph hyperlink = paragraph.AddHyperLink(hyperlinkText, new Uri(hyperlinkUri));
            hyperlink.Bold = true;
            hyperlink.Italic = true;
            paragraph.AddText(finalRun).Italic = true;
            document.Save();
        }
        stream.Position = 0;

        OfficeDocumentReadResult result = OfficeIMO.Reader.Tests.ReaderTestReaders.Word().ReadDocument(
            stream,
            "long-paragraph.docx",
            new ReaderOptions { MaxChars = 256 });
        ReaderChunk[] chunks = result.Chunks.Where(chunk =>
            chunk.Location.SourceBlockKind == "paragraph").ToArray();

        Assert.True(chunks.Length > 1);
        Assert.All(chunks, chunk => {
            Assert.True(chunk.Text.Length <= 256);
            Assert.True((chunk.Markdown?.Length ?? 0) <= 256);
            Assert.NotEqual(paragraphText, chunk.Text);
            AssertWellFormedUtf16(chunk.Text);
            AssertWellFormedUtf16(chunk.Markdown ?? string.Empty);
            Assert.Equal(
                CountOccurrences(chunk.Markdown ?? string.Empty, "["),
                CountOccurrences(chunk.Markdown ?? string.Empty, "](" + hyperlinkUri + ")"));
        });
        Assert.Equal(paragraphText, string.Concat(chunks.Select(chunk => chunk.Text)));
        Assert.Contains(chunks, chunk => (chunk.Markdown ?? string.Empty).Contains(hyperlinkUri, StringComparison.Ordinal));
    }

    [Fact]
    public void DocumentReader_WordRichMapping_PreservesHeaderlessFirstRow() {
        using var stream = new MemoryStream();
        using (WordDocument document = WordDocument.Create(stream)) {
            WordTable table = document.AddTable(2, 1);
            table.RepeatAsHeaderRowAtTheTopOfEachPage = false;
            table.Rows[0].Cells[0].Paragraphs[0].Text = "First value";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "Second value";
            document.Save();
        }
        stream.Position = 0;

        OfficeDocumentReadResult result = OfficeIMO.Reader.Tests.ReaderTestReaders.Word()
            .ReadDocument(stream, "headerless.docx");
        ReaderTable mapped = Assert.Single(result.Tables);

        Assert.Equal(new[] { "Column 1" }, mapped.Columns);
        Assert.Equal(2, mapped.TotalRowCount);
        Assert.Equal(new[] { "First value", "Second value" }, mapped.Rows.Select(row => row[0]));
        ReaderTable chunkTable = Assert.Single(Assert.Single(
            result.Chunks,
            chunk => chunk.Tables?.Count == 1).Tables!);
        Assert.Equal(new[] { "Column 1" }, chunkTable.Columns);
        Assert.Equal(2, chunkTable.TotalRowCount);
        Assert.Equal(new[] { "First value", "Second value" },
            chunkTable.Rows.Select(row => row[0]));
    }

    [Fact]
    public void DocumentReader_WordChunksDoNotDuplicateOversizedHyperlinkDestinations() {
        string destination = "https://example.test/" + new string('a', 3000);
        using var stream = new MemoryStream();
        using (WordDocument document = WordDocument.Create(stream)) {
            document.AddParagraph().AddHyperLink(new string('x', 1000), new Uri(destination));
            document.Save();
        }
        stream.Position = 0;

        OfficeDocumentReadResult result = OfficeIMO.Reader.Tests.ReaderTestReaders.Word().ReadDocument(
            stream,
            "bounded-hyperlink.docx",
            new ReaderOptions { MaxChars = 256 });

        Assert.All(result.Chunks, chunk => Assert.True((chunk.Markdown?.Length ?? 0) <= 256));
        Assert.DoesNotContain(result.Chunks, chunk => (chunk.Markdown ?? string.Empty).Contains(destination, StringComparison.Ordinal));
        Assert.Contains(result.Chunks.SelectMany(chunk => chunk.Warnings ?? Array.Empty<string>()),
            warning => warning.Contains("hyperlink destination", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void DocumentReader_ExcelRichMapping_UsesFormalTablesCellLinksAndProperties() {
        using var stream = new MemoryStream();
        using (ExcelDocument document = ExcelDocument.Create(stream)) {
            document.BuiltinDocumentProperties.Title = "Rich workbook";
            document.BuiltinDocumentProperties.Creator = "OfficeIMO";
            ExcelSheet sheet = document.AddWorksheet("Inventory");
            sheet.Cell(1, 1, "Name");
            sheet.Cell(1, 2, "Qty");
            sheet.Cell(2, 1, "Bandage");
            sheet.Cell(2, 2, 4);
            sheet.AddTable("A1:B2", hasHeader: true, name: "InventoryTable", style: TableStyle.TableStyleMedium2);
            sheet.SetHyperlink(2, 1, "https://example.test/bandage", display: "Bandage");
            sheet.Cell(1, 4, "Loose");
            sheet.Cell(1, 5, "Value");
            sheet.Cell(2, 4, "Unstructured");
            sheet.Cell(2, 5, 7);
            ExcelSheet rawSheet = document.AddWorksheet("Raw");
            rawSheet.Cell(1, 1, "Metric");
            rawSheet.Cell(1, 2, "Value");
            rawSheet.Cell(2, 1, "Retries");
            rawSheet.Cell(2, 2, 3);
            document.Save();
        }
        stream.Position = 0;

        OfficeDocumentReadResult result = OfficeIMO.Reader.Tests.ReaderTestReaders.Excel().ReadDocument(stream, "inventory.xlsx");

        Assert.Equal("Rich workbook", result.Source.Title);
        Assert.Equal("OfficeIMO", result.Source.Author);
        ReaderTable mapped = Assert.Single(result.Tables, table => table.Kind == "excel-table");
        Assert.Equal("InventoryTable", mapped.Title);
        Assert.Equal("excel-table", mapped.Kind);
        Assert.Equal("Bandage", mapped.Rows[0][0]);
        Assert.Contains(result.Tables, table => table.Location?.Sheet == "Raw");
        OfficeDocumentLink link = Assert.Single(result.Links);
        Assert.Equal("Inventory", link.Location.Sheet);
        Assert.Equal("A2", link.Location.A1Range);
        Assert.Equal("https://example.test/bandage", link.Uri);
        Assert.DoesNotContain(result.Tables, table => table.Kind != "excel-table" && table.Location?.Sheet == "Inventory");
        Assert.Contains(result.Blocks, block => block.Location.Sheet == "Inventory" && block.Text.Contains("Unstructured", StringComparison.Ordinal));
        OfficeDocumentPage inventoryPage = Assert.Single(result.Pages, page => page.Name == "Inventory");
        Assert.Contains(inventoryPage.Tables, table => ReferenceEquals(table, mapped));
        Assert.Single(inventoryPage.Tables);
        Assert.Contains("officeimo.excel.inspection-snapshot", result.CapabilitiesUsed);
    }

    [Fact]
    public void DocumentReader_ExcelRichMapping_HonorsSelectedRangeAcrossRichArtifacts() {
        using var stream = new MemoryStream();
        using (ExcelDocument document = ExcelDocument.Create(stream)) {
            ExcelSheet sheet = document.AddWorksheet("Inventory");
            sheet.Cell(1, 1, "Name");
            sheet.Cell(1, 2, "Qty");
            sheet.Cell(2, 1, "Bandage");
            sheet.CellFormula(2, 2, "1+1");
            sheet.Cell(3, 1, "Gauze");
            sheet.Cell(3, 2, 3);
            sheet.AddTable("A1:B3", hasHeader: true, name: "InventoryTable", style: TableStyle.TableStyleMedium2);
            sheet.SetHyperlink(2, 1, "https://example.test/inside", display: "Bandage");
            sheet.SetComment(2, 1, "Inside comment");
            sheet.CellFormula(2, 4, "2+2");
            sheet.SetHyperlink(2, 4, "https://example.test/outside", display: "Outside");
            sheet.SetComment(2, 4, "Outside comment");
            document.Save();
        }
        stream.Position = 0;

        OfficeDocumentReadResult result = OfficeIMO.Reader.Tests.ReaderTestReaders.Excel(a1Range: "A1:B2").ReadDocument(
            stream,
            "inventory.xlsx");

        OfficeDocumentLink link = Assert.Single(result.Links);
        Assert.Equal("https://example.test/inside", link.Uri);
        ReaderTable table = Assert.Single(result.Tables, candidate => candidate.Kind == "excel-table");
        Assert.Equal("A1:B2", table.Location!.A1Range);
        Assert.Single(table.Rows);
        Assert.Equal("1", Assert.Single(result.Metadata, item => item.Name == "FormulaCount").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, item => item.Name == "CommentCount").Value);
    }

    [Fact]
    public void DocumentReader_PowerPointRichMapping_UsesShapesTablesChartsLinksAndSlideGeometry() {
        using var stream = new MemoryStream();
        using (PowerPointPresentation presentation = PowerPointPresentation.Create(stream)) {
            presentation.BuiltinDocumentProperties.Title = "Rich deck";
            presentation.BuiltinDocumentProperties.Creator = "OfficeIMO";
            PowerPointSlide slide = presentation.AddSlide();
            slide.Notes.Text = "Speaker guidance";
            PowerPointTextBox title = slide.AddTextBox("Overview");
            title.AddParagraph("Summary");
            title.Paragraphs[0].Runs[0].SetHyperlink("https://example.test/deck");
            PowerPointTextBox hidden = slide.AddTextBox("Hidden guidance");
            hidden.Hidden = true;
            PowerPointTable table = slide.AddTable(2, 2);
            table.GetCell(0, 0).Text = "Name";
            table.GetCell(0, 1).Text = "Qty";
            table.GetCell(1, 0).Text = "Bandage";
            table.GetCell(1, 1).Text = "4";
            slide.AddChart(OfficeChartKind.ColumnClustered, new OfficeChartData(
                new[] { "Q1", "Q2" },
                new[] { new OfficeChartSeries("Sales", new[] { 1D, 2D }) }));
            presentation.Save();
        }
        stream.Position = 0;

        OfficeDocumentReadResult result = OfficeIMO.Reader.Tests.ReaderTestReaders.PowerPoint().ReadDocument(stream, "deck.pptx");

        Assert.Equal("Rich deck", result.Source.Title);
        Assert.Equal("OfficeIMO", result.Source.Author);
        OfficeDocumentPage page = Assert.Single(result.Pages);
        Assert.True(page.Width > 0);
        Assert.True(page.Height > 0);
        Assert.Contains(result.Blocks, block => block.Kind == "paragraph" && block.Text == "Overview" && block.Region != null);
        Assert.Equal("Bandage", Assert.Single(result.Tables).Rows[0][0]);
        Assert.Contains(result.Blocks, block => block.Text == "Hidden guidance");
        Assert.Equal("https://example.test/deck", Assert.Single(result.Links).Uri);
        ReaderVisual chart = Assert.Single(result.Visuals);
        Assert.Equal("chart", chart.Kind);
        Assert.Contains("Sales", chart.Content, StringComparison.Ordinal);
        Assert.Equal("4", Assert.Single(result.Metadata, item => item.Name == "ShapeCount").Value);
        OfficeDocumentBlock notes = Assert.Single(result.Blocks, block => block.Kind == "speaker-notes");
        Assert.Equal("Speaker guidance", notes.Text);
        Assert.Same(notes, Assert.Single(page.Blocks, block => block.Kind == "speaker-notes"));
        Assert.Contains("officeimo.powerpoint.chart-snapshot", result.CapabilitiesUsed);

        stream.Position = 0;
        OfficeDocumentReadResult withoutNotes = OfficeIMO.Reader.Tests.ReaderTestReaders.PowerPoint(includeNotes: false).ReadDocument(stream, "deck.pptx");
        Assert.DoesNotContain(withoutNotes.Blocks, block => block.Kind == "speaker-notes");
        Assert.DoesNotContain(Assert.Single(withoutNotes.Pages).Blocks, block => block.Kind == "speaker-notes");
    }

    [Fact]
    public void DocumentReader_PowerPointRichMapping_PreservesHeaderlessFirstRow() {
        using var stream = new MemoryStream();
        using (PowerPointPresentation presentation = PowerPointPresentation.Create(stream)) {
            PowerPointTable table = presentation.AddSlide().AddTable(2, 1);
            table.HeaderRow = false;
            table.GetCell(0, 0).Text = "First value";
            table.GetCell(1, 0).Text = "Second value";
            presentation.Save();
        }
        stream.Position = 0;

        ReaderTable mapped = Assert.Single(OfficeIMO.Reader.Tests.ReaderTestReaders.PowerPoint().ReadDocument(stream, "headerless.pptx").Tables);

        Assert.Equal(new[] { "Column 1" }, mapped.Columns);
        Assert.Equal(2, mapped.TotalRowCount);
        Assert.Equal(new[] { "First value", "Second value" }, mapped.Rows.Select(row => row[0]));
    }

    [Fact]
    public void DocumentReader_PowerPointRichMapping_HonorsHiddenShapeFilteringAcrossArtifacts() {
        byte[] png = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==");
        uint hiddenImageFillId;
        using var stream = new MemoryStream();
        using (PowerPointPresentation presentation = PowerPointPresentation.Create(stream)) {
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTextBox visible = slide.AddTextBox("Visible guidance");
            visible.Paragraphs[0].Runs[0].SetHyperlink("https://example.test/visible");
            PowerPointTextBox hiddenTitle = slide.AddTextBox("Hidden title");
            hiddenTitle.PlaceholderType = DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title;
            hiddenTitle.Hidden = true;
            hiddenTitle.Paragraphs[0].Runs[0].SetHyperlink("https://example.test/hidden");
            PowerPointTable hiddenTable = slide.AddTable(1, 1);
            hiddenTable.GetCell(0, 0).Text = "Hidden table";
            hiddenTable.Hidden = true;
            PowerPointChart hiddenChart = slide.AddChart(OfficeChartKind.ColumnClustered, new OfficeChartData(
                new[] { "Q1" },
                new[] { new OfficeChartSeries("Hidden series", new[] { 1D }) }));
            hiddenChart.Hidden = true;

            PowerPointPicture visiblePicture = slide.AddPicture(
                new MemoryStream(png, writable: false), OfficeIMO.PowerPoint.ImagePartType.Png);
            visiblePicture.AltText = "Visible direct picture";
            PowerPointPicture hiddenPicture = slide.AddPicture(
                new MemoryStream(png, writable: false), OfficeIMO.PowerPoint.ImagePartType.Png);
            hiddenPicture.AltText = "Hidden direct picture";
            hiddenPicture.Hidden = true;

            PowerPointAutoShape hiddenImageFill = slide.AddRectangle(0, 0, 914400, 914400);
            hiddenImageFill.Hidden = true;
            hiddenImageFillId = hiddenImageFill.Id ?? throw new InvalidOperationException("Hidden image-fill shape has no id.");

            PowerPointTextBox hiddenGroupText = slide.AddTextBox("Hidden group guidance");
            PowerPointPicture hiddenGroupPicture = slide.AddPicture(
                new MemoryStream(png, writable: false), OfficeIMO.PowerPoint.ImagePartType.Png);
            hiddenGroupPicture.AltText = "Hidden group picture";
            PowerPointGroupShape hiddenGroup = slide.GroupShapes(
                new PowerPointShape[] { hiddenGroupText, hiddenGroupPicture }, "Hidden group");
            hiddenGroup.Hidden = true;

            PowerPointTextBox visibleGroupText = slide.AddTextBox("Visible group guidance");
            PowerPointPicture visibleGroupPicture = slide.AddPicture(
                new MemoryStream(png, writable: false), OfficeIMO.PowerPoint.ImagePartType.Png);
            visibleGroupPicture.AltText = "Visible group picture";
            PowerPointPicture hiddenGroupChildPicture = slide.AddPicture(
                new MemoryStream(png, writable: false), OfficeIMO.PowerPoint.ImagePartType.Png);
            hiddenGroupChildPicture.AltText = "Hidden group child picture";
            hiddenGroupChildPicture.Hidden = true;
            slide.GroupShapes(
                new PowerPointShape[] { visibleGroupText, visibleGroupPicture, hiddenGroupChildPicture },
                "Visible group");
            presentation.Save();
        }

        stream.Position = 0;
        using (DocumentFormat.OpenXml.Packaging.PresentationDocument document =
               DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(stream, true)) {
            DocumentFormat.OpenXml.Packaging.SlidePart slidePart = document.PresentationPart!.SlideParts.Single();
            DocumentFormat.OpenXml.Packaging.ImagePart hiddenFillPart = slidePart.AddImagePart(
                DocumentFormat.OpenXml.Packaging.ImagePartType.Png);
            using (var hiddenFillStream = new MemoryStream(png, writable: false)) {
                hiddenFillPart.FeedData(hiddenFillStream);
            }
            string hiddenFillRelationshipId = slidePart.GetIdOfPart(hiddenFillPart);
            P.Shape hiddenFillShape = slidePart.Slide.Descendants<P.Shape>().Single(shape =>
                shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value == hiddenImageFillId);
            hiddenFillShape.ShapeProperties!.Append(new A.BlipFill(
                new A.Blip { Embed = hiddenFillRelationshipId },
                new A.Stretch(new A.FillRectangle())));
            slidePart.Slide.Save();
        }
        stream.Position = 0;

        OfficeDocumentReadResult result = OfficeIMO.Reader.Tests.ReaderTestReaders.PowerPoint(
            includeNotes: false,
            includeHiddenShapes: false).ReadDocument(stream, "visible-only.pptx");
        OfficeDocumentPage page = Assert.Single(result.Pages);

        Assert.Contains(result.Blocks, block => block.Text == "Visible guidance");
        Assert.Contains(result.Blocks, block => block.Text == "Visible group guidance");
        Assert.DoesNotContain(result.Blocks, block => block.Text.Contains("Hidden", StringComparison.Ordinal));
        Assert.DoesNotContain(page.Blocks, block => block.Text.Contains("Hidden", StringComparison.Ordinal));
        Assert.DoesNotContain(result.Chunks, chunk => chunk.Text.Contains("Hidden", StringComparison.Ordinal));
        Assert.Empty(result.Tables);
        Assert.Empty(result.Visuals);
        Assert.Equal("https://example.test/visible", Assert.Single(result.Links).Uri);
        Assert.Equal("Slide 1", page.Name);
        Assert.Equal(2, result.Assets.Count);
        Assert.Equal(2, page.Assets.Count);
        Assert.Equal(
            new[] { "Visible direct picture", "Visible group picture" },
            result.Assets.Select(asset => asset.AltText).OrderBy(static value => value, StringComparer.Ordinal).ToArray());
        Assert.Equal("5", Assert.Single(result.Metadata, item => item.Name == "ShapeCount").Value);
    }

    private static int CountOccurrences(string value, string marker) {
        int count = 0;
        int offset = 0;
        while ((offset = value.IndexOf(marker, offset, StringComparison.Ordinal)) >= 0) {
            count++;
            offset += marker.Length;
        }
        return count;
    }

    private static void AssertWellFormedUtf16(string value) {
        for (int index = 0; index < value.Length; index++) {
            if (char.IsHighSurrogate(value[index])) {
                Assert.True(index + 1 < value.Length && char.IsLowSurrogate(value[index + 1]));
                index++;
            } else {
                Assert.False(char.IsLowSurrogate(value[index]));
            }
        }
    }
}
