using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Fact]
    public void Table_NormalizesCellsAndSnapshotsInputRowsBeforeRendering() {
        var body = new[] { "Original", (string)null! };
        var rows = new[] {
            new[] { "Name", "Value" },
            body
        };

        var doc = PdfDocument.Create()
            .Table(rows);

        body[0] = "Mutated";
        body[1] = "AlsoMutated";

        byte[] bytes = doc.ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = string.Concat(pdf.GetPage(1).Letters.Select(letter => letter.Value));

        Assert.Contains("Original", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Mutated", text, StringComparison.Ordinal);
        Assert.DoesNotContain("AlsoMutated", text, StringComparison.Ordinal);
    }

    [Fact]
    public void TableBlock_SnapshotsRowsStyleAndLinksIntoReadOnlyModel() {
        var body = new[] { "Original", (string)null! };
        var style = TableStyles.Minimal();
        style.BorderWidth = 2;
        style.CellPaddingX = 7;
        style.RowSeparatorColor = new PdfColor(0.11, 0.22, 0.33);
        style.RowSeparatorWidth = 0.8;
        style.HeaderSeparatorColor = new PdfColor(0.44, 0.55, 0.66);
        style.HeaderSeparatorWidth = 1.2;
        style.FooterSeparatorColor = new PdfColor(0.22, 0.33, 0.44);
        style.FooterSeparatorWidth = 1.4;
        style.MaxWidth = 160;
        style.LeftIndent = 18;

        var block = new TableBlock(new[] {
            new[] { "Name", "Value" },
            body
        }, PdfAlign.Left, style);

        block.AddLink((1, 0), "https://evotec.xyz");
        body[0] = "Mutated";
        body[1] = "AlsoMutated";
        style.BorderWidth = 5;
        style.CellPaddingX = 20;
        style.RowSeparatorColor = PdfColor.Black;
        style.RowSeparatorWidth = 2;
        style.HeaderSeparatorColor = PdfColor.White;
        style.HeaderSeparatorWidth = 3;
        style.FooterSeparatorColor = PdfColor.White;
        style.FooterSeparatorWidth = 4;
        style.MaxWidth = 220;
        style.LeftIndent = 30;

        Assert.False(block.Rows is List<string[]>);
        Assert.False(block.Cells is List<IReadOnlyList<PdfTableCell>>);
        Assert.False(block.Links is Dictionary<(int Row, int Col), string>);
        Assert.Equal("Original", block.Rows[1][0]);
        Assert.Equal(string.Empty, block.Rows[1][1]);
        Assert.Equal(2, block.ColumnCount);
        Assert.Equal("Original", block.Cells[1][0].Text);
        Assert.Equal(1, block.Cells[1][0].ColumnSpan);
        Assert.Equal(1, block.Cells[1][0].RowSpan);
        Assert.Equal(2, block.Style!.BorderWidth);
        Assert.Equal(7, block.Style.CellPaddingX);
        Assert.Equal(new PdfColor(0.11, 0.22, 0.33), block.Style.RowSeparatorColor);
        Assert.Equal(0.8, block.Style.RowSeparatorWidth);
        Assert.Equal(new PdfColor(0.44, 0.55, 0.66), block.Style.HeaderSeparatorColor);
        Assert.Equal(1.2, block.Style.HeaderSeparatorWidth);
        Assert.Equal(new PdfColor(0.22, 0.33, 0.44), block.Style.FooterSeparatorColor);
        Assert.Equal(1.4, block.Style.FooterSeparatorWidth);
        Assert.Equal(160, block.Style.MaxWidth);
        Assert.Equal(18, block.Style.LeftIndent);
        Assert.True(block.Links.TryGetValue((1, 0), out string? uri));
        Assert.Equal("https://evotec.xyz", uri);
    }

    [Fact]
    public void PdfDocument_DefaultTableStyleAppliesToFollowingTablesAndSnapshotsInput() {
        var options = new PdfOptions {
            PageWidth = 300,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };
        var style = TableStyles.Minimal();
        style.CellPaddingX = 22;

        byte[] bytes = PdfDocument.Create(options)
            .DefaultTableStyle(style)
            .Table(new[] {
                new[] { "DefaultPad", "Value" },
                new[] { "Row", "1" }
            })
            .ToBytes();

        style.CellPaddingX = 0;

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double markerX = FindWordStartX(page, "DefaultPad");

        Assert.True(markerX - options.MarginLeft >= 20, $"Expected fluent default table style padding to affect following tables and snapshot caller input. Marker x: {markerX:0.##}, margin: {options.MarginLeft:0.##}.");
    }

    [Fact]
    public void PdfDocument_ExplicitTableStyleOverridesDefaultTableStyle() {
        var options = new PdfOptions {
            PageWidth = 300,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };
        var defaultStyle = TableStyles.Minimal();
        defaultStyle.CellPaddingX = 24;
        var explicitStyle = TableStyles.Minimal();
        explicitStyle.CellPaddingX = 2;

        byte[] bytes = PdfDocument.Create(options)
            .DefaultTableStyle(defaultStyle)
            .Table(new[] {
                new[] { "ExplicitPad", "Value" },
                new[] { "Row", "1" }
            }, style: explicitStyle)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double markerX = FindWordStartX(page, "ExplicitPad");

        Assert.True(markerX - options.MarginLeft <= 6, $"Expected explicit table style padding to override the document default. Marker x: {markerX:0.##}, margin: {options.MarginLeft:0.##}.");
    }

    [Fact]
    public void PdfDocument_DefaultTableStyleRejectsInvalidInputs() {
        Assert.Throws<ArgumentNullException>(() => PdfDocument.Create().DefaultTableStyle((PdfTableStyle)null!));
        Assert.Throws<ArgumentNullException>(() => PdfDocument.Create().DefaultTableStyle((string)null!));
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().DefaultTableStyle("Missing Table Style"));
    }

    [Fact]
    public void TableWithLinks_SnapshotsInputLinkDictionaryBeforeRendering() {
        var links = new Dictionary<(int Row, int Col), string> {
            [(1, 0)] = "https://evotec.xyz"
        };

        var doc = PdfDocument.Create()
            .TableWithLinks(
                new[] { new[] { "Name" }, new[] { "OfficeIMO" } },
                links);

        links[(1, 0)] = "https://example.com";

        byte[] bytes = doc.ToBytes();
        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("(https://evotec.xyz)", content);
        Assert.DoesNotContain("(https://example.com)", content);
    }


}
