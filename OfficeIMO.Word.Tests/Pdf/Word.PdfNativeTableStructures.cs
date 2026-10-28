using DocumentFormat.OpenXml;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("start", false, false)]
    [InlineData("end", false, false)]
    [InlineData("start", true, false)]
    [InlineData("end", true, false)]
    [InlineData("start", false, true)]
    [InlineData("end", false, true)]
    [InlineData("start", true, true)]
    [InlineData("end", true, true)]
    public void SaveAsPdf_LogicalTableAlignmentUsesEffectiveDirection(string token, bool rightToLeft, bool inherited) {
        using WordDocument document = WordDocument.Create();
        WordTable left = AddNativeInputTable(document, "PhysicalLeft");
        left.Alignment = WordTableAlignment.Left;
        WordTable right = AddNativeInputTable(document, "PhysicalRight");
        right.Alignment = WordTableAlignment.Right;
        WordTable logical = AddNativeInputTable(document, "LogicalCell");
        logical._tableProperties!.AddChild(new W.BiDiVisual { Val = rightToLeft ? W.OnOffOnlyValues.On : W.OnOffOnlyValues.Off }, true);
        W.TableJustification declaration = new();
        declaration.SetAttribute(new OpenXmlAttribute("w", "val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", token));
        if (inherited) {
            document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(new W.Style(
                new W.StyleName { Val = "Logical alignment" },
                new W.BasedOn { Val = "TableGrid" },
                new W.StyleTableProperties(declaration)) {
                    Type = W.StyleValues.Table, StyleId = "LogicalAlignment"
                });
            logical._tableProperties!.TableStyle = new W.TableStyle { Val = "LogicalAlignment" };
            logical._tableProperties.TableJustification = null;
        } else {
            logical._tableProperties!.TableJustification = declaration;
            Assert.Equal((token == "end") != rightToLeft ? WordTableAlignment.Right : WordTableAlignment.Left,
                logical.Alignment);
        }

        using WordDocument loaded = WordDocument.Load(new MemoryStream(document.ToBytes()));
        using PdfPigDocument pdf = PdfPigDocument.Open(loaded.ToPdf(NativeInputTableOptions()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var page = Assert.Single(pdf.GetPages());
        double logicalX = Assert.Single(page.GetWords(), word => word.Text == "LogicalCell").BoundingBox.Left;
        bool alignRight = (token == "end") != rightToLeft;
        string referenceText = alignRight ? "PhysicalRight" : "PhysicalLeft";
        double referenceX = Assert.Single(page.GetWords(), word => word.Text == referenceText).BoundingBox.Left;
        Assert.Equal(referenceX, logicalX, 2);
        Assert.Equal(token, (inherited
            ? loaded._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Elements<W.Style>()
                .Single(style => style.StyleId == "LogicalAlignment").GetFirstChild<W.StyleTableProperties>()!.GetFirstChild<W.TableJustification>()!
            : loaded.Tables.Last()._tableProperties!.TableJustification!).Val!.InnerText);
    }

    [Theory]
    [InlineData("body", false)]
    [InlineData("header", false)]
    [InlineData("footer", false)]
    [InlineData("body", true)]
    [InlineData("header", true)]
    [InlineData("footer", true)]
    public void SaveAsPdf_ContentControlCellsRetainTableGridAndStoryText(string storyKind, bool wrappedRow) {
        using WordDocument document = WordDocument.Create();
        WordHeaderFooter? story = storyKind == "header" ? document.HeaderDefaultOrCreate :
            storyKind == "footer" ? document.FooterDefaultOrCreate : null;
        WordTable table = story == null ? document.AddTable(1, 2) : story.AddTable(1, 2);
        ConfigureNativeInputTable(table, 4000);
        table.GridColumnWidth = new List<int> { 1600, 2400 };
        table.Rows[0].Cells[0].Width = 1600;
        table.Rows[0].Cells[1].Width = 2400;
        table.Rows[0].Cells[0].Paragraphs[0].Text = "LeftCell";
        table.Rows[0].Cells[1].Paragraphs[0].Text = "ControlledCell";
        table.Rows[0].Cells[1].VerticalAlignment = WordTableVerticalAlignment.Bottom;
        W.TableCell controlled = table.Rows[0].Cells[1]._tableCell;
        controlled.Parent!.ReplaceChild(new W.SdtCell(new W.SdtProperties(new W.Tag { Val = "inert-cell" }),
            new W.SdtContentCell((W.TableCell)controlled.CloneNode(true))), controlled);
        if (wrappedRow) {
            W.TableRow row = table.Rows[0]._tableRow;
            row.Parent!.ReplaceChild(new W.SdtRow(new W.SdtProperties(new W.Tag { Val = "inert-row" }),
                new W.SdtContentRow((W.TableRow)row.CloneNode(true))), row);
        }
        document.AddParagraph("BodyMarker");

        using WordDocument loaded = WordDocument.Load(new MemoryStream(document.ToBytes()));
        WordTable loadedTable = storyKind == "header" ? loaded.Sections[0].ResolveDefaultHeader()!.Tables.Single() :
            storyKind == "footer" ? loaded.Sections[0].ResolveDefaultFooter()!.Tables.Single() : loaded.Tables.Single();
        Assert.Equal(2, Assert.Single(loadedTable.Rows).Cells.Count);
        Assert.Equal(new[] { "LeftCell", "ControlledCell" }, loadedTable.Rows[0].Cells.Select(cell => cell.Paragraphs[0].Text));
        using PdfPigDocument pdf = PdfPigDocument.Open(loaded.ToPdf(NativeInputTableOptions()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var page = Assert.Single(pdf.GetPages());
        var leftWord = Assert.Single(page.GetWords(), word => word.Text == "LeftCell");
        var controlledWord = Assert.Single(page.GetWords(), word => word.Text == "ControlledCell");
        Assert.InRange(controlledWord.BoundingBox.Left - leftWord.BoundingBox.Left, 79.9, 80.1);
        Assert.Contains("BodyMarker", page.Text);
        Assert.Single(loadedTable._table.Descendants<W.SdtCell>());
        Assert.Equal(wrappedRow ? 1 : 0, loadedTable._table.Elements<W.SdtRow>().Count());
    }

    [Fact]
    public void ContentControlCellSpansUseLogicalRowCellsAndExcludeNestedTables() {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(2, 2);
        ConfigureNativeInputTable(table, 4000);
        table.GridColumnWidth = new List<int> { 2000, 2000 };
        WordTableCell anchor = table.Rows[0].Cells[0];
        anchor.VerticalMerge = WordCellMerge.Restart;
        anchor.Paragraphs[0].Text = "MergedAnchor";
        table.Rows[1].Cells[0].VerticalMerge = WordCellMerge.Continue;
        foreach (WordTableRow row in table.Rows) {
            W.TableCell first = row.Cells[0]._tableCell;
            first.Parent!.ReplaceChild(new W.SdtCell(new W.SdtContentCell((W.TableCell)first.CloneNode(true))), first);
        }
        WordTable nested = table.Rows[0].Cells[1].AddTable(1, 3);
        nested.Rows[0].Cells[0].Paragraphs[0].Text = "NestedMarker";
        WordTableCell reloadedAnchor = table.Rows[0].Cells[0];
        Assert.Equal(2, reloadedAnchor.RowSpan);
        Assert.Equal(2, table.Rows[0].Cells.Count);
        Assert.Equal(3, Assert.Single(table.Rows[0].Cells[1].DirectNestedTables).Rows[0].Cells.Count);
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdf(NativeInputTableOptions()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        string text = string.Concat(pdf.GetPages().Select(page => page.Text));
        Assert.Equal(1, text.Split(new[] { "MergedAnchor" }, StringSplitOptions.None).Length - 1);
        Assert.Equal(1, text.Split(new[] { "NestedMarker" }, StringSplitOptions.None).Length - 1);
    }

    private static WordTable AddNativeInputTable(WordDocument document, string marker) {
        WordTable table = document.AddTable(1, 1);
        ConfigureNativeInputTable(table, 2000);
        table.Rows[0].Cells[0].Paragraphs[0].Text = marker;
        return table;
    }

    private static void ConfigureNativeInputTable(WordTable table, int width) {
        table.LayoutMode = WordTableLayoutMode.Fixed;
        table.WidthType = WordTableWidthUnit.Dxa;
        table.Width = width;
        table.GridColumnWidth = new List<int> { width };
        foreach (WordTableRow row in table.Rows) {
            foreach (WordTableCell cell in row.Cells) {
                cell.WidthType = WordTableWidthUnit.Dxa;
                cell.Width = width / row.Cells.Count;
                cell.MarginLeft = global::OfficeIMO.OfficeLength.Zero;
                cell.MarginRight = global::OfficeIMO.OfficeLength.Zero;
                cell.Paragraphs[0].FontFamily = "Arial";
                cell.Paragraphs[0].FontSize = 10;
                cell.Paragraphs[0].LineSpacingBefore = global::OfficeIMO.OfficeLength.Zero;
                cell.Paragraphs[0].LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
            }
        }
    }

    private static WordPdfExportOptions NativeInputTableOptions() => new() {
        IncludePageNumbers = false,
        PageSize = new PdfCore.PdfPageSize(OfficeIMO.OfficeLength.Points(400),OfficeIMO.OfficeLength.Points(500)),
        Margins = PdfCore.PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(40))
    };
}
