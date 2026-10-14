using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using System.Security.Cryptography;
using System.Text.Json;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkPagesTableWritingTests {
    [Fact]
    public void Apple_accepted_tables_retain_the_native_saved_edit_and_effective_layout_in_strict_Word() {
        string corpus = Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "native-exports");
        using var manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(corpus, "pages-created-tables-v15.4.json")));
        JsonElement artifacts = manifest.RootElement.GetProperty("artifacts");
        foreach (JsonProperty artifact in artifacts.EnumerateObject()) {
            string path = Path.Combine(corpus, artifact.Value.GetProperty("path").GetString()!);
            Assert.Equal(artifact.Value.GetProperty("sha256").GetString(),
                Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant());
        }
        Assert.Equal(File.ReadAllBytes(Asset("created")), Example().SaveBytes());
        var source = IWorkSourceDocument.Open(Asset("nativeEdited"));
        var projection = source.ReadPages();
        Assert.Empty(projection.Diagnostics); Assert.Empty(projection.SourceDeclarationIssues); Assert.Empty(projection.SourceReferenceIssues);
        Assert.Equal(2, projection.Tables.Count);
        IWorkTable values = projection.Tables[0], empty = projection.Tables[1];
        Assert.Equal(20.75, values.GetCell(2, 1)!.Value);
        Assert.Equal("A😀B\nZażółć", values.GetCell(1, 1)!.Value);
        Assert.Equal("=SUM(A1:A2)", values.GetCell(1, 2)!.Value);
        Assert.Equal(IWorkCellKind.Text, values.GetCell(1, 2)!.Kind);
        Assert.Equal(true, values.GetCell(2, 2)!.Value); Assert.Equal(false, values.GetCell(2, 3)!.Value);
        Assert.Empty(empty.Cells);
        foreach (IWorkTable table in projection.Tables) {
            for (int row = 1; row <= table.RowCount; row++) {
                for (int column = 1; column <= table.ColumnCount; column++) {
                    IWorkCellPadding padding = Assert.IsType<IWorkCellPadding>(table.GetPadding(row, column));
                    Assert.Equal(new double[] { 4, 4, 4, 4 },
                        new[] { padding.LeftPoints, padding.TopPoints, padding.RightPoints, padding.BottomPoints });
                    Assert.Equal(IWorkCellVerticalAlignment.Top, table.GetVerticalAlignment(row, column));
                }
            }
        }
        using var result = source.ToWordDocumentResult(new() { Mode = IWorkConversionMode.EditableOnly });
        result.RequireCompleteEditableReconstruction();
        using var output = new MemoryStream(); result.Value.Save(output); output.Position = 0;
        using var actual = WordprocessingDocument.Open(output, false);
        using var native = WordprocessingDocument.Open(Asset("nativeEditedDocx"), false);
        W.Table[] tables = actual.MainDocumentPart!.Document!.Body!.Elements<W.Table>().ToArray();
        W.Table[] reference = native.MainDocumentPart!.Document!.Body!.Elements<W.Table>().ToArray();
        Assert.Equal(2, tables.Length); Assert.Equal(2, reference.Length);
        for (int table = 0; table < tables.Length; table++) {
            Assert.Equal(reference[table].GetFirstChild<W.TableGrid>()!.Elements<W.GridColumn>().Select(column => column.Width!.Value),
                tables[table].GetFirstChild<W.TableGrid>()!.Elements<W.GridColumn>().Select(column => column.Width!.Value));
            Assert.Equal(reference[table].Elements<W.TableRow>().Count(), tables[table].Elements<W.TableRow>().Count());
        }
        Assert.Contains("20.75", tables[0].InnerText);
        Assert.Empty(new OpenXmlValidator().Validate(actual));
        string Asset(string key) => Path.Combine(corpus, artifacts.GetProperty(key).GetProperty("path").GetString()!);
    }

    [Fact]
    public void Inline_tables_snapshot_scalars_and_keep_body_order_through_strict_saved_Word() {
        var document = Example();
        Assert.Equal(new[] { typeof(IWorkPagesParagraphData), typeof(IWorkPagesTable), typeof(IWorkPagesTable), typeof(IWorkPagesParagraphData) },
            document.BodyBlocks.Select(block => block.GetType()));
        Assert.Equal(new[] { "Before 😀", "After" }, document.Paragraphs.Select(paragraph => paragraph.Text));
        Assert.Equal(new[] { "Values", "Empty" }, document.Tables.Select(table => table.Name));
        byte[] bytes = document.SaveBytes(); Assert.Equal(bytes, Example().SaveBytes());
        var projection = IWorkSourceDocument.Open(bytes).ReadPages();
        Assert.Empty(projection.Diagnostics); Assert.Empty(projection.SourceDeclarationIssues); Assert.Empty(projection.SourceReferenceIssues);
        Assert.Equal(2, projection.Tables.Count);
        IWorkTable table = projection.Tables[0];
        Assert.Equal("Values", table.Name); Assert.False(table.AutoResizeRows);
        Assert.Equal(3, table.RowCount); Assert.Equal(3, table.ColumnCount);
        Assert.Equal(0, table.Geometry!.LeftPoints); Assert.Equal(0, table.Geometry.TopPoints);
        Assert.Equal(324, table.Geometry.WidthPoints); Assert.Equal(144, table.Geometry.HeightPoints);
        Assert.Equal(48, table.GetRowHeight(3)); Assert.Equal(108, table.GetColumnWidth(3));
        Assert.Equal("A😀B\nZażółć", Cell(1, 1).Value);
        Assert.Equal("=SUM(A1:A2)", Cell(1, 2).Value); Assert.Equal(IWorkCellKind.Text, Cell(1, 2).Kind);
        Assert.Equal(-1234.5, Cell(2, 1).Value); Assert.Equal(true, Cell(2, 2).Value); Assert.Equal(false, Cell(2, 3).Value);
        Assert.DoesNotContain(table.Cells, cell => cell.Row == 3 && cell.Column == 3);
        Assert.Equal(4, table.GetPadding(3, 3)!.BottomPoints);
        Assert.Equal(IWorkCellVerticalAlignment.Top, table.GetVerticalAlignment(3, 3));
        Assert.Equal("003366", table.GetParagraphStyle(3, 3)!.TextStyle.Color!.RgbHex);
        using var result = WordIWorkConverter.ConvertPagesToWordResult(new MemoryStream(bytes),
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly });
        result.RequireCompleteEditableReconstruction();
        using var output = new MemoryStream(); result.Value.Save(output); output.Position = 0;
        using var saved = WordprocessingDocument.Open(output, false);
        Assert.Empty(new OpenXmlValidator().Validate(saved));
        var body = saved.MainDocumentPart!.Document!.Body!;
        W.Table[] tables = body.Elements<W.Table>().ToArray(); Assert.Equal(2, tables.Length);
        Assert.Equal("Before 😀", body.FirstChild!.InnerText);
        Assert.IsType<W.Paragraph>(tables[1].PreviousSibling());
        Assert.Equal(new[] { "2160", "2160", "2160" }, tables[0].GetFirstChild<W.TableGrid>()!
            .Elements<W.GridColumn>().Select(column => column.Width!.Value));
        Assert.Equal(new uint[] { 880, 880, 880 }, tables[0].Elements<W.TableRow>()
            .Select(row => row.TableRowProperties!.GetFirstChild<W.TableRowHeight>()!.Val!.Value));
        Assert.Contains(body.Elements<W.Paragraph>(), paragraph => paragraph.InnerText == "After");
        Assert.Contains("=SUM(A1:A2)", tables[0].InnerText);
        Assert.Equal(4, tables[1].Descendants<W.TableCell>().Count());
        IWorkTableCell Cell(int row, int column) => table.Cells.Single(cell => cell.Row == row && cell.Column == column);
    }

    [Fact]
    public void A_table_only_document_is_editable_and_snapshots_input_without_exposing_mutable_collections() {
        var document = IWorkPagesDocument.Create();
        var data = IWorkTableData.Create(2, 2).SetText(1, 1, "original");
        var table = document.AddTable(data);
        byte[] bytes = document.SaveBytes();
        data.SetText(1, 1, "changed").SetBoolean(2, 2, true);
        Assert.Equal(bytes, document.SaveBytes()); Assert.Equal("original", Assert.Single(table.Cells).Value);
        Assert.Empty(document.Paragraphs); Assert.Same(table, Assert.Single(document.BodyBlocks));
        Assert.Throws<NotSupportedException>(() => ((IList<IWorkPagesBodyBlock>)document.BodyBlocks).Clear());
        Assert.Throws<NotSupportedException>(() => ((IList<IWorkPagesTable>)document.Tables).Clear());
        Assert.Throws<NotSupportedException>(() => ((IList<IWorkTableCell>)table.Cells).Clear());
        using var converted = WordIWorkConverter.ConvertPagesToWordResult(new MemoryStream(bytes),
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly });
        converted.RequireCompleteEditableReconstruction();
        using var output = new MemoryStream(); converted.Value.Save(output); output.Position = 0;
        using var saved = WordprocessingDocument.Open(output, false);
        Assert.IsType<W.Table>(saved.MainDocumentPart!.Document!.Body!.FirstChild);
        Assert.Empty(new OpenXmlValidator().Validate(saved));
    }

    [Theory]
    [InlineData("tables")]
    [InlineData("cells")]
    [InlineData("paragraphs")]
    [InlineData("runs")]
    [InlineData("characters")]
    public void Table_budgets_fail_before_mutating_the_caller_stream(string boundary) {
        var limits = new IWorkPagesWriteOptions();
        switch (boundary) {
            case "tables": limits.MaximumTables = 1; break;
            case "cells": limits.MaximumTableCells = 12; break;
            case "paragraphs": limits.MaximumParagraphs = 3; break;
            case "runs": limits.MaximumTextRuns = 3; break;
            case "characters": limits.MaximumTextCharacters = 20; break;
        }
        using var output = new MemoryStream(new byte[] { 9, 8, 7 });
        Assert.Throws<InvalidDataException>(() => Example().Save(output, limits));
        Assert.Equal(new byte[] { 9, 8, 7 }, output.ToArray());
    }

    [Fact]
    public void Limits_count_unset_cells_table_names_and_UTF16_attachment_text_at_the_exact_boundary() {
        var document = IWorkPagesDocument.Create();
        document.AddTable(IWorkTableData.Create(2, 3).SetText(1, 1, "x😀"), "T");
        Assert.Throws<InvalidDataException>(() => document.SaveBytes(new() { MaximumTextCharacters = 5 }));
        Assert.Throws<InvalidDataException>(() => document.SaveBytes(new() { MaximumTableCells = 5 }));
        var projection = IWorkSourceDocument.Open(document.SaveBytes(new() {
            MaximumTextCharacters = 6, MaximumTableCells = 6, MaximumTables = 1, MaximumParagraphs = 1, MaximumTextRuns = 1
        })).ReadPages();
        Assert.Equal("x😀", Assert.Single(Assert.Single(projection.Tables).Cells).Value);
    }

    [Fact]
    public void Invalid_table_names_dimensions_and_limits_preserve_the_document() {
        var document = IWorkPagesDocument.Create(); var data = IWorkTableData.Create(2, 2);
        document.AddTable(data, "Values"); byte[] bytes = document.SaveBytes();
        Assert.Throws<ArgumentException>(() => document.AddTable(data, "VALUES"));
        Assert.Throws<ArgumentException>(() => document.AddTable(data, "bad\nname"));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.AddTable(data, "Other", columnWidthPoints: float.NaN));
        Assert.Throws<ArgumentException>(() => document.AddTable(data, "Other", columnWidthPoints: 250));
        Assert.Throws<ArgumentException>(() => document.AddTable(data, "Other", rowHeightPoints: 350));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.SaveBytes(new() { MaximumTables = 0 }));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.SaveBytes(new() { MaximumTableCells = 0 }));
        Assert.Equal(bytes, document.SaveBytes()); Assert.Single(document.Tables); Assert.Single(document.BodyBlocks);
    }

    internal static IWorkPagesDocument Example() {
        var document = IWorkPagesDocument.Create();
        document.AddParagraph("Before 😀");
        var data = IWorkTableData.Create(3, 3).SetText(1, 1, "A😀B\r\nZażółć")
            .SetText(1, 2, "=SUM(A1:A2)").SetText(1, 3, "")
            .SetNumber(2, 1, -1234.5).SetBoolean(2, 2, true).SetBoolean(2, 3, false);
        document.AddTable(data, "Values", 108, 48, "Arial", 12, "003366");
        document.AddTable(IWorkTableData.Create(2, 2), "Empty", 100, 28);
        document.AddParagraph("After", new IWorkPagesParagraphLayout(1.5f, 6, 9));
        return document;
    }
}
