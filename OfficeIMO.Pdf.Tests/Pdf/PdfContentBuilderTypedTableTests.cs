using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO;
using OfficeIMO.Data;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfContentBuilderTypedTableTests {
    [Fact]
    public void TypedProjectionEnumeratesOnceAndSavesSelectedOrderedFormattedColumns() {
        var source = new SinglePass<Row>(new[] { new Row("Ada", 12.5m), new Row("Grace", 7m) });
        var options = new PdfTableOptions { Name = "Team totals", HeaderTransform = name => "Header " + name };
        options.Exclude.Add(nameof(Row.Unused));
        options.Order.Add(nameof(Row.Amount));
        options.HeaderRenames[nameof(Row.Amount)] = "Total";
        options.Formatters[nameof(Row.Amount)] = value => ((decimal)value!).ToString("0.00", System.Globalization.CultureInfo.InvariantCulture);
        options.Alignments.Add(OfficeAlignment.Right);
        var style = new PdfTableStyle { HeaderRowCount = 0, Caption = "Original" };
        options.Style = style;
        using var document = PdfDocument.Create(builder => builder.Content(content => content.Table(source, options)));
        var block = Assert.IsType<TableBlock>(Assert.Single(document.Blocks));
        Assert.Equal(new[] { "Total", "Header Name" }, block.Rows[0]);
        Assert.Equal(new[] { "12.50", "Ada" }, block.Rows[1]);
        Assert.Equal(PdfColumnAlign.Right, block.Style!.Alignments![0]);
        Assert.Equal(1, block.Style.HeaderRowCount);
        Assert.Equal(0, style.HeaderRowCount);
        Assert.Equal("Original", style.Caption);
        using var readback = PdfPigDocument.Open(document.ToBytes());
        string text = string.Join(" ", readback.GetPages().Select(page => page.Text));
        Assert.Contains("Team totals", text);
        Assert.Contains("12.50", text);
        Assert.Contains("Grace", text);
        Assert.Equal(1, source.Enumerations);
    }

    [Fact]
    public void HeaderlessDictionaryProjectionKeepsSparseCellsAndCommonSelection() {
        var rows = new[] {
            new Dictionary<string, object?> { ["Name"] = "Ada", ["Amount"] = 3, ["Ignored"] = "omit" },
            new Dictionary<string, object?> { ["Name"] = "Grace", ["Ignored"] = "omit" }
        };
        var options = new OfficeTableOptions { IncludeHeaders = false };
        options.Include.UnionWith(new[] { "Name", "Amount" });
        options.Order.Add("Amount");
        using var document = PdfDocument.Create(builder => builder.Content(content => content.Table(rows, options)));
        var block = Assert.IsType<TableBlock>(Assert.Single(document.Blocks));
        Assert.Equal(new[] { "3", "Ada" }, block.Rows[0]);
        Assert.Equal(new[] { "", "Grace" }, block.Rows[1]);
        Assert.Equal(0, block.Style!.HeaderRowCount);
        Assert.Equal(0, block.Style.RepeatHeaderRowCount);
        using var readback = PdfPigDocument.Open(document.ToBytes());
        Assert.DoesNotContain("Ignored", string.Join(" ", readback.GetPages().Select(page => page.Text)));
    }

    [Fact]
    public void EmptyTypedRowsKeepHeadersAndExistingArrayOverloadsKeepTheirCells() {
        using var document = PdfDocument.Create(builder => builder.Content(content => content
            .Table(Array.Empty<Row>(), new OfficeTableOptions { Include = { nameof(Row.Name) } })
            .Table(new[] { new[] { "Raw", "Text" } })
            .Table(new[] { new[] { new PdfTableCell("Explicit") } })));
        var blocks = document.Blocks.Cast<TableBlock>().ToArray();
        Assert.Equal(new[] { "Name" }, Assert.Single(blocks[0].Rows));
        Assert.Equal(new[] { "Raw", "Text" }, Assert.Single(blocks[1].Rows));
        Assert.Equal("Explicit", Assert.Single(Assert.Single(blocks[2].Rows)));
        Assert.Throws<ArgumentException>(() => PdfDocument.Create(builder => builder.Content(content => content.Table(
            Array.Empty<Row>(), new OfficeTableOptions { IncludeHeaders = false }))));
        using var readback = PdfPigDocument.Open(document.ToBytes());
        Assert.Single(readback.GetPages());
    }

    [Fact]
    public void NestedExpansionUsesSharedOptionsAndEnforcesExpandedCellBudget() {
        var source = new[] { new { Name = "Ada", Tags = new[] { "one", "two" } } };
        int scalarCalls = 0;
        var options = new PdfTableOptions { Projection = new ObjectFlattenerOptions { CollectionMode = CollectionMode.ExpandRows } };
        options.Formatters["Tags"] = value => "[" + value + "]";
        options.Formatters["Name"] = value => { scalarCalls++; return value?.ToString()?.ToUpperInvariant(); };
        using var document = PdfDocument.Create(builder => builder.Content(content => content.Table(source, options)));
        var block = Assert.IsType<TableBlock>(Assert.Single(document.Blocks));
        Assert.Equal(new[] { "ADA", "[one]" }, block.Rows[1]);
        Assert.Equal(new[] { "ADA", "[two]" }, block.Rows[2]);
        Assert.Equal(1, scalarCalls);
        Assert.Throws<InvalidDataException>(() => PdfDocument.Create(builder => builder.Content(content => content.Table(source,
            new PdfTableOptions { Projection = new ObjectFlattenerOptions { CollectionMode = CollectionMode.ExpandRows, MaxCells = 4 } }))));
        Assert.Throws<InvalidDataException>(() => PdfDocument.Create(builder => builder.Content(content => content.Table(source,
            new PdfTableOptions { Projection = new ObjectFlattenerOptions { CollectionMode = CollectionMode.ExpandRows, MaxRows = 1 } }))));
    }

    [Theory]
    [InlineData("rows")]
    [InlineData("columns")]
    [InlineData("cells")]
    public void SharedProjectionRejectsMaterializationBeyondConfiguredLimits(string limit) {
        var rows = new[] { new Row("Ada", 1), new Row("Grace", 2) };
        var projection = new ObjectFlattenerOptions();
        if (limit == "rows") projection.MaxRows = 1;
        if (limit == "columns") projection.MaxColumns = 1;
        if (limit == "cells") projection.MaxCells = 5;
        Assert.Throws<InvalidDataException>(() => PdfDocument.Create(builder => builder.Content(content => content.Table(rows,
            new PdfTableOptions { Exclude = { nameof(Row.Unused) }, Projection = projection }))));
    }

    [Fact]
    public void CommonOrderingPrecedesAdvancedColumnsAndFormatterFailureDoesNotAppendATable() {
        var options = new PdfTableOptions {
            Projection = new ObjectFlattenerOptions { Columns = new[] { "Amount", "Name" } }
        };
        options.Order.Add("Name");
        using var document = PdfDocument.Create(builder => builder.Content(content => content.Table(
            new[] { new { Name = "Ada", Amount = 12 } }, options)));
        Assert.Equal(new[] { "Name", "Amount" }, Assert.IsType<TableBlock>(Assert.Single(document.Blocks)).Rows[0]);
        Assert.Equal(new[] { "Amount", "Name" }, options.Projection.Columns);

        var rejected = new PdfTableOptions { Projection = new ObjectFlattenerOptions { CollectionMode = CollectionMode.ExpandRows } };
        rejected.Formatters["Tags"] = _ => throw new InvalidOperationException("Cannot format this cell");
        using var afterFailure = PdfDocument.Create(builder => builder.Content(content => {
            Assert.Throws<InvalidOperationException>(() => content.Table(new[] { new { Tags = new[] { "one" } } }, rejected));
            content.Paragraph(paragraph => paragraph.Text("Still usable"));
        }));
        Assert.IsType<RichParagraphBlock>(Assert.Single(afterFailure.Blocks));
    }

    [Fact]
    public void ExhaustedExplicitProjectionNeverDiscoversUnselectedGetters() {
        var options = new PdfTableOptions { Projection = new ObjectFlattenerOptions { Columns = new[] { nameof(WhitelistedRow.Name) } } };
        options.Order.Add(nameof(WhitelistedRow.Name));
        options.Exclude.Add(nameof(WhitelistedRow.Name));
        var row = new WhitelistedRow();
        using var document = PdfDocument.Create(builder => builder.Content(content => {
            Assert.Throws<ArgumentException>(() => content.Table(new[] { row }, options));
            content.Paragraph(paragraph => paragraph.Text("Selection remains empty"));
        }));
        Assert.Equal(0, row.SecretReads);
        Assert.IsType<RichParagraphBlock>(Assert.Single(document.Blocks));
        Assert.Equal(new[] { nameof(WhitelistedRow.Name) }, options.Projection.Columns);
    }

    private sealed class WhitelistedRow {
        public string Name => "Ada";
        public int SecretReads { get; private set; }
        public string Secret { get { SecretReads++; throw new InvalidOperationException("An unselected getter was read"); } }
    }

    private sealed record Row(string Name, decimal Amount) {
        public string Unused => throw new InvalidOperationException("Excluded properties must not be evaluated.");
    }

    private sealed class SinglePass<T>(IEnumerable<T> rows) : IEnumerable<T> {
        internal int Enumerations { get; private set; }
        public IEnumerator<T> GetEnumerator() {
            if (++Enumerations != 1) throw new InvalidOperationException("Source was enumerated twice.");
            return rows.GetEnumerator();
        }
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
