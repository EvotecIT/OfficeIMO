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
        var style = PdfTableStyles.Minimal();
        style.BorderWidth = global::OfficeIMO.OfficeLength.Points(2);
        style.CellPaddingX = global::OfficeIMO.OfficeLength.Points(7);
        style.RowSeparatorColor = global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.11, 0.22, 0.33);
        style.RowSeparatorWidth = global::OfficeIMO.OfficeLength.Points(0.8);
        style.HeaderSeparatorColor = global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.44, 0.55, 0.66);
        style.HeaderSeparatorWidth = global::OfficeIMO.OfficeLength.Points(1.2);
        style.FooterSeparatorColor = global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.22, 0.33, 0.44);
        style.FooterSeparatorWidth = global::OfficeIMO.OfficeLength.Points(1.4);
        style.MaxWidth = global::OfficeIMO.OfficeLength.Points(160);
        style.LeftIndent = global::OfficeIMO.OfficeLength.Points(18);

        var block = new TableBlock(new[] {
            new[] { "Name", "Value" },
            body
        }, OfficeAlignment.Left, style);

        body[0] = "Mutated";
        body[1] = "AlsoMutated";
        style.BorderWidth = global::OfficeIMO.OfficeLength.Points(5);
        style.CellPaddingX = global::OfficeIMO.OfficeLength.Points(20);
        style.RowSeparatorColor = global::OfficeIMO.OfficeColor.Black;
        style.RowSeparatorWidth = global::OfficeIMO.OfficeLength.Points(2);
        style.HeaderSeparatorColor = global::OfficeIMO.OfficeColor.White;
        style.HeaderSeparatorWidth = global::OfficeIMO.OfficeLength.Points(3);
        style.FooterSeparatorColor = global::OfficeIMO.OfficeColor.White;
        style.FooterSeparatorWidth = global::OfficeIMO.OfficeLength.Points(4);
        style.MaxWidth = global::OfficeIMO.OfficeLength.Points(220);
        style.LeftIndent = global::OfficeIMO.OfficeLength.Points(30);

        Assert.False(block.Rows is List<string[]>);
        Assert.False(block.Cells is List<IReadOnlyList<PdfTableCell>>);
        Assert.Equal("Original", block.Rows[1][0]);
        Assert.Equal(string.Empty, block.Rows[1][1]);
        Assert.Equal(2, block.ColumnCount);
        Assert.Equal("Original", block.Cells[1][0].Text);
        Assert.Equal(1, block.Cells[1][0].ColumnSpan);
        Assert.Equal(1, block.Cells[1][0].RowSpan);
        Assert.Equal(2, block.Style!.BorderWidth.ToPoints());
        Assert.Equal(7, block.Style.CellPaddingX.ToPoints());
        Assert.Equal(global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.11, 0.22, 0.33), block.Style.RowSeparatorColor);
        Assert.Equal(0.8, block.Style.RowSeparatorWidth.ToPoints());
        Assert.Equal(global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.44, 0.55, 0.66), block.Style.HeaderSeparatorColor);
        Assert.Equal(1.2, block.Style.HeaderSeparatorWidth.ToPoints());
        Assert.Equal(global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.22, 0.33, 0.44), block.Style.FooterSeparatorColor);
        Assert.Equal(1.4, block.Style.FooterSeparatorWidth.ToPoints());
        Assert.Equal(160, block.Style.MaxWidth?.ToPoints());
        Assert.Equal(18, block.Style.LeftIndent.ToPoints());
    }

    [Fact]
    public void TableBlock_SnapshotsMutableCellImageStyle() {
        var image = new PdfTableCellImage(
            PdfPngTestImages.CreateRgbaPng(24, 48, 72, 255),
            global::OfficeIMO.OfficeLength.Points(12),
            global::OfficeIMO.OfficeLength.Points(12),
            new PdfImageStyle { AltText = "Original logo" });
        PdfTableCell cell = PdfTableCell.WithImages("Logo", new[] { image });
        var block = new TableBlock(new[] { new[] { cell } }, OfficeAlignment.Left, null);

        cell.Images[0].Style!.AltText = "Changed logo";

        Assert.Equal("Original logo", block.Cells[0][0].Images[0].Style!.AltText);
    }

    [Fact]
    public void PdfDocument_DefaultTableStyleAppliesToFollowingTablesAndSnapshotsInput() {
        var options = new PdfCreateOptions {
            PageSetup = {
                Size = new PdfPageSize(global::OfficeIMO.OfficeLength.Points(300), global::OfficeIMO.OfficeLength.Points(220)),
                Margins = { Left = global::OfficeIMO.OfficeLength.Points(30), Right = global::OfficeIMO.OfficeLength.Points(30), Top = global::OfficeIMO.OfficeLength.Points(30), Bottom = global::OfficeIMO.OfficeLength.Points(30) }
            },
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
        };
        var style = PdfTableStyles.Minimal();
        style.CellPaddingX = global::OfficeIMO.OfficeLength.Points(22);

        byte[] bytes = PdfDocument.Create(options)
            .DefaultTableStyle(style)
            .Table(new[] {
                new[] { "DefaultPad", "Value" },
                new[] { "Row", "1" }
            })
            .ToBytes();

        style.CellPaddingX = global::OfficeIMO.OfficeLength.Points(0);

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double markerX = FindWordStartX(page, "DefaultPad");

        Assert.True(markerX - options.PageSetup.Margins.Left.ToPoints() >= 20, $"Expected fluent default table style padding to affect following tables and snapshot caller input. Marker x: {markerX:0.##}, margin: {options.PageSetup.Margins.Left.ToPoints():0.##}.");
    }

    [Fact]
    public void PdfDocument_ExplicitTableStyleOverridesDefaultTableStyle() {
        var options = new PdfCreateOptions {
            PageSetup = {
                Size = new PdfPageSize(global::OfficeIMO.OfficeLength.Points(300), global::OfficeIMO.OfficeLength.Points(220)),
                Margins = { Left = global::OfficeIMO.OfficeLength.Points(30), Right = global::OfficeIMO.OfficeLength.Points(30), Top = global::OfficeIMO.OfficeLength.Points(30), Bottom = global::OfficeIMO.OfficeLength.Points(30) }
            },
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
        };
        var defaultStyle = PdfTableStyles.Minimal();
        defaultStyle.CellPaddingX = global::OfficeIMO.OfficeLength.Points(24);
        var explicitStyle = PdfTableStyles.Minimal();
        explicitStyle.CellPaddingX = global::OfficeIMO.OfficeLength.Points(2);

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

        Assert.True(markerX - options.PageSetup.Margins.Left.ToPoints() <= 6, $"Expected explicit table style padding to override the document default. Marker x: {markerX:0.##}, margin: {options.PageSetup.Margins.Left.ToPoints():0.##}.");
    }

    [Fact]
    public void PdfDocument_DefaultTableStyleRejectsInvalidInputs() {
        Assert.Throws<ArgumentNullException>(() => PdfDocument.Create().DefaultTableStyle((PdfTableStyle)null!));
        Assert.Throws<ArgumentNullException>(() => PdfDocument.Create().DefaultTableStyle((string)null!));
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().DefaultTableStyle("Missing Table Style"));
    }


}
