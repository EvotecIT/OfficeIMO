using System.Globalization;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Pdf.Tests;

/// <summary>
/// Tables read from the rules a page paints, built the way spreadsheet and report exports draw them, with no
/// structure tags to fall back on.
/// </summary>
public sealed class PdfRuleLatticeTableTests {
    private static readonly string[][] Market = {
        new[] { "ID", "Market", "Region", "Status", "Amount" },
        new[] { "1", "Alpha Futures", "EU", "Open", "10.50" },
        new[] { "2", "Beta Options", "US", "Closed", "7.25" },
        new[] { "3", "Gamma Swaps", "APAC", "Open", "99.00" },
        new[] { "4", "Delta Futures", "EU", "Hold", "1.75" },
        new[] { "5", "Epsilon Spreads", "US", "Open", "64.10" },
        new[] { "6", "Zeta Options", "APAC", "Closed", "3.30" }
    };

    [Fact]
    public void SpreadsheetBordersPaintedAsColouredBarsOverFillsReadAsOneTable() {
        // Excel-style paint: cell fills first, then borders as 0.6 pt filled bars in several colours, column
        // dividers drawn only in the header (white on the header fill), a logo image elsewhere on the page, a
        // wrapped cell and an empty one.
        string[][] expected = Market.Select(row => row.ToArray()).ToArray();
        expected[3][1] = "Gamma Swaps and Options";
        expected[4][2] = string.Empty;
        double[] x = { 72, 112, 272, 332, 392, 472 };
        double[] y = { 700, 680, 664, 648, 622, 606, 590, 574 };
        var content = new StringBuilder();
        content.Append("q 60 0 0 24 72 740 cm BI /W 1 /H 1 /CS /G /BPC 8 /F /AHx ID 80> EI Q\n");
        Fill(content, "0.44 0.77 0.91 rg", x[0], y[1], x[5] - x[0], y[0] - y[1]);
        for (int row = 2; row < y.Length - 1; row += 2) Fill(content, "0.85 g", x[0], y[row + 1], x[5] - x[0], y[row] - y[row + 1]);
        string[][] painted = expected.Select(row => row.ToArray()).ToArray();
        painted[3][1] = string.Empty;
        Text(content, painted, x, y, "1 g");
        content.Append("BT /F1 9 Tf 0 g 1 0 0 1 115 ").Append(Number(y[4] + 14)).Append(" Tm (Gamma Swaps and) Tj ET\n");
        content.Append("BT /F1 9 Tf 0 g 1 0 0 1 115 ").Append(Number(y[4] + 4)).Append(" Tm (Options) Tj ET\n");
        for (int row = 0; row < y.Length; row++) {
            // Each row rule is painted per colour run: black over the first columns, grey over the rest.
            Fill(content, "0 g", x[0], y[row] - 0.3, x[2] - x[0], 0.6);
            Fill(content, "0.65 g", x[2], y[row] - 0.3, x[5] - x[2], 0.6);
        }
        Fill(content, "0 g", x[0] - 0.3, y[y.Length - 1], 0.6, y[0] - y[y.Length - 1]);
        Fill(content, "0 g", x[5] - 0.3, y[y.Length - 1], 0.6, y[0] - y[y.Length - 1]);
        Fill(content, "0.65 g", x[1] - 0.3, y[y.Length - 1], 0.6, y[0] - y[y.Length - 1]);
        for (int column = 2; column < 5; column++) Fill(content, "1 g", x[column] - 0.3, y[1], 0.6, y[0] - y[1]);

        PdfLogicalTable table = Assert.Single(Read(content, "Helvetica").Tables);

        Assert.Contains(table.Evidence, evidence => evidence.Code == "table-ruling-lattice");
        Assert.Equal(expected, table.Rows.Select(row => row.ToArray()));
    }

    [Fact]
    public void RunStretchedAcrossADrawnDividerSplitsBetweenItsWords() {
        // One string per row; word spacing stretches the space so the second word starts in the next column.
        string[][] rows = {
            new[] { "Code", "Item" },
            new[] { "A-01", "Widget" },
            new[] { "B-02", "Sprocket" },
            new[] { "C-03", "Flange" }
        };
        double[] x = { 72, 162, 262 };
        double[] y = Rows(700, 16, 16, rows.Length - 1);
        var content = new StringBuilder();
        content.Append("0 G 0.5 w\n");
        foreach (double row in y) Line(content, x[0], row, x[x.Length - 1], row);
        foreach (double column in x) Line(content, column, y[0], column, y[y.Length - 1]);
        for (int row = 0; row < rows.Length; row++) {
            // Courier is 5.4 pt per character at 9 pt.
            double spacing = x[1] + 3 - (x[0] + 3 + rows[row][0].Length * 5.4) - 5.4;
            content.Append("BT /F1 9 Tf ").Append(Number(spacing)).Append(" Tw 0 g 1 0 0 1 ").Append(Number(x[0] + 3)).Append(' ')
                .Append(Number(y[row + 1] + 5)).Append(" Tm (").Append(rows[row][0]).Append(' ').Append(rows[row][1]).Append(") Tj ET\n");
        }

        PdfLogicalTable table = Assert.Single(Read(content, "Courier").Tables);

        Assert.Equal(rows, table.Rows.Select(row => row.ToArray()));
    }

    [Fact]
    public void RunOverflowingPastADividerDrawnOnlyInTheHeaderStaysInItsCell() {
        // The divider between the second and third columns is drawn through the header only; body text that
        // runs past it overflows, as spreadsheets let it, rather than starting a new cell.
        string[][] rows = {
            new[] { "Code", "Item", "Note" },
            new[] { "A-01", "Widget assembly", "" },
            new[] { "B-02", "Sprocket housing", "" },
            new[] { "C-03", "Flange bracket", "" }
        };
        double[] x = { 72, 162, 222, 300 };
        double[] y = Rows(700, 16, 16, rows.Length - 1);
        var content = new StringBuilder();
        content.Append("0 G 0.5 w\n");
        foreach (double row in y) Line(content, x[0], row, x[x.Length - 1], row);
        foreach (double column in new[] { x[0], x[1], x[3] }) Line(content, column, y[0], column, y[y.Length - 1]);
        Line(content, x[2], y[0], x[2], y[1]);
        for (int column = 0; column < 3; column++) {
            content.Append("BT /F1 9 Tf 0 g 1 0 0 1 ").Append(Number(x[column] + 3)).Append(' ').Append(Number(y[1] + 5))
                .Append(" Tm (").Append(rows[0][column]).Append(") Tj ET\n");
        }
        for (int row = 1; row < rows.Length; row++) {
            // One string per row; word spacing carries the item across the first divider.
            double spacing = x[1] + 3 - (x[0] + 3 + rows[row][0].Length * 5.4) - 5.4;
            content.Append("BT /F1 9 Tf ").Append(Number(spacing)).Append(" Tw 0 g 1 0 0 1 ").Append(Number(x[0] + 3)).Append(' ')
                .Append(Number(y[row + 1] + 5)).Append(" Tm (").Append(rows[row][0]).Append(' ')
                .Append(rows[row][1].Replace(" ", "\\040")).Append(") Tj 0 Tw ET\n");
        }

        PdfLogicalTable table = Assert.Single(Read(content, "Courier").Tables);

        Assert.Equal(rows, table.Rows.Select(row => row.ToArray()));
    }

    [Fact]
    public void WhiteBordersOfDarkCellsPaintedWithOneOperatorReadAsAGrid() {
        // Each cell is one `re B`: a dark fill with a white stroke over it. A logo image keeps the strict stroke
        // reader out, and each row is one string, so only the white borders can split the rows into cells.
        string[][] rows = {
            new[] { "Code", "Item" },
            new[] { "A-01", "Widget" },
            new[] { "B-02", "Sprocket" },
            new[] { "C-03", "Flange" }
        };
        double[] x = { 72, 162, 262 };
        double[] y = Rows(700, 16, 16, rows.Length - 1);
        var content = new StringBuilder();
        content.Append("q 60 0 0 24 72 740 cm BI /W 1 /H 1 /CS /G /BPC 8 /F /AHx ID 80> EI Q\n");
        content.Append("0.2 0.3 0.5 rg 1 G 0.75 w\n");
        for (int row = 0; row + 1 < y.Length; row++) {
            for (int column = 0; column + 1 < x.Length; column++) {
                content.Append(Number(x[column])).Append(' ').Append(Number(y[row + 1])).Append(' ')
                    .Append(Number(x[column + 1] - x[column])).Append(' ').Append(Number(y[row] - y[row + 1])).Append(" re B\n");
            }
        }
        for (int row = 0; row < rows.Length; row++) {
            double spacing = x[1] + 3 - (x[0] + 3 + rows[row][0].Length * 5.4) - 5.4;
            content.Append("BT /F1 9 Tf ").Append(Number(spacing)).Append(" Tw 1 g 1 0 0 1 ").Append(Number(x[0] + 3)).Append(' ')
                .Append(Number(y[row + 1] + 5)).Append(" Tm (").Append(rows[row][0]).Append(' ').Append(rows[row][1]).Append(") Tj ET\n");
        }

        PdfLogicalTable table = Assert.Single(Read(content, "Courier").Tables);

        Assert.Contains(table.Evidence, evidence => evidence.Code == "table-ruling-lattice");
        Assert.Equal(rows, table.Rows.Select(row => row.ToArray()));
    }

    [Fact]
    public void StencilImagePaintedOverAGridLeavesItsBordersVisible() {
        // A stencil paints only where its samples say; the rest of its box shows the grid beneath it.
        PdfLogicalTable table = Assert.Single(Read(GridUnderImage(out string[][] rows,
            "q 200 0 0 80 60 630 cm BI /W 1 /H 1 /IM true /BPC 1 /F /AHx ID 00> EI Q\n"), "Courier").Tables);

        Assert.Contains(table.Evidence, evidence => evidence.Code == "table-ruling-lattice");
        Assert.Equal(rows, table.Rows.Select(row => row.ToArray()));
    }

    [Fact]
    public void OpaqueImageClippedOverAGridHidesItsBorders() {
        // The image's box reaches above the clip; the clip alone, over the table, decides what shows.
        StringBuilder content = GridUnderImage(out _,
            "q 60 620 240 90 re W n 240 0 0 120 60 600 cm BI /W 1 /H 1 /CS /G /BPC 8 /F /AHx ID 80> EI Q\n");

        Assert.DoesNotContain(Read(content, "Courier").Tables,
            table => table.Evidence.Any(evidence => evidence.Code == "table-ruling-lattice"));
    }

    [Fact]
    public void TaggedTableThatFailsItsChecksLeavesThePageToItsPaintedBorders() {
        // The structure tree's Table holds a paragraph instead of rows, so the reader rejects it and reads the
        // page's geometry, painted borders included, as it would without tags.
        StringBuilder content = GridUnderImage(out string[][] rows, "q 60 0 0 24 72 740 cm BI /W 1 /H 1 /CS /G /BPC 8 /F /AHx ID 80> EI Q\n");
        content.Insert(0, "/P << /MCID 0 >> BDC\n").Append("EMC\n");

        PdfLogicalPage page = Read(content, "Courier", tagged: true, out PdfDocumentReadResult read);

        Assert.Contains(read.TaggedTableDiagnostics, diagnostic => diagnostic.Reason == PdfTaggedTableFallbackReason.UnsupportedStructure);
        PdfLogicalTable table = Assert.Single(page.Tables);
        Assert.Contains(table.Evidence, evidence => evidence.Code == "table-ruling-lattice");
        Assert.Equal(rows, table.Rows.Select(row => row.ToArray()));
    }

    [Fact]
    public void UntaggedGridBesideATaggedTableStillReadsFromItsBorders() {
        // A tagged two-column table at the top right passes its checks; the bordered grid below it is marked as a
        // paragraph. The geometry finds the grid as a table from its columns, but one row is a single string whose
        // stretched space crosses the divider: only the border splits it into its two cells.
        string[][] rows = {
            new[] { "Code", "Item" },
            new[] { "A-01", "Widget" },
            new[] { "B-02", "Sprocket" },
            new[] { "C-03", "Flange" },
            new[] { "D-04", "Gasket" }
        };
        double[] x = { 72, 162, 262 };
        double[] y = Rows(700, 16, 16, rows.Length - 1);
        var content = new StringBuilder("/P << /MCID 6 >> BDC\n0 G 0.75 w\n");
        foreach (double top in y) Line(content, x[0], top, x[x.Length - 1], top);
        foreach (double left in x) Line(content, left, y[0], left, y[y.Length - 1]);
        content.Append("q 60 0 0 24 72 740 cm BI /W 1 /H 1 /CS /G /BPC 8 /F /AHx ID 80> EI Q\n");
        for (int row = 0; row < rows.Length; row++) {
            string baseline = Number(y[row + 1] + 5);
            if (row == 2) {
                double spacing = x[1] + 3 - (x[0] + 3 + rows[row][0].Length * 5.4) - 5.4;
                content.Append("BT /F1 9 Tf ").Append(Number(spacing)).Append(" Tw 0 g 1 0 0 1 ").Append(Number(x[0] + 3)).Append(' ').Append(baseline)
                    .Append(" Tm (").Append(rows[row][0]).Append(' ').Append(rows[row][1]).Append(") Tj ET\n");
                continue;
            }
            for (int column = 0; column < 2; column++) {
                content.Append("BT /F1 9 Tf 0 Tw 0 g 1 0 0 1 ").Append(Number(x[column] + 3)).Append(' ').Append(baseline)
                    .Append(" Tm (").Append(rows[row][column]).Append(") Tj ET\n");
            }
        }
        content.Append("EMC\n");
        string[][] tagged = { new[] { "Name", "Qty" }, new[] { "Bolt", "4" }, new[] { "Nut", "9" } };
        var structure = new List<string> { "<< /Type /StructTreeRoot /K [7 0 R 17 0 R] /ParentTree << /Nums [0 [11 0 R 12 0 R 13 0 R 14 0 R 15 0 R 16 0 R 17 0 R]] >> >>",
            "<< /Type /StructElem /S /Table /P 6 0 R /K [8 0 R 9 0 R 10 0 R] /Pg 3 0 R >>" };
        for (int row = 0; row < tagged.Length; row++) {
            structure.Add("<< /Type /StructElem /S /TR /P 7 0 R /K [" + (11 + row * 2) + " 0 R " + (12 + row * 2) + " 0 R] /Pg 3 0 R >>");
            for (int column = 0; column < 2; column++) {
                content.Append("/TD << /MCID ").Append(row * 2 + column).Append(" >> BDC BT /F1 9 Tf 0 g 1 0 0 1 ").Append(320 + column * 100).Append(' ')
                    .Append(770 - row * 14).Append(" Tm (").Append(tagged[row][column]).Append(") Tj ET EMC\n");
            }
        }
        for (int cell = 0; cell < 6; cell++) structure.Add("<< /Type /StructElem /S /TD /P " + (8 + cell / 2) + " 0 R /K " + cell + " /Pg 3 0 R >>");
        structure.Add("<< /Type /StructElem /S /P /P 6 0 R /K 6 /Pg 3 0 R >>");

        PdfLogicalPage page = Read(content, "Courier", tagged: true, out PdfDocumentReadResult read, structure: structure);

        Assert.Empty(read.TaggedTableDiagnostics);
        Assert.Contains(page.Tables, table => table.Rows.Select(row => row.ToArray()).SequenceEqual(tagged, ArrayComparer));
        Assert.Contains(page.Tables, table => table.Rows.Select(row => row.ToArray()).SequenceEqual(rows, ArrayComparer));
    }

    [Fact]
    public void NearlyTransparentBordersAreNotReadAsAGrid() {
        // Black at 1% opacity on a white page looks white, so it splits nothing.
        StringBuilder content = GridUnderImage(out _, "q 60 0 0 24 72 740 cm BI /W 1 /H 1 /CS /G /BPC 8 /F /AHx ID 80> EI Q\n");
        content.Insert(0, "/GS1 gs\n");

        PdfLogicalPage page = Read(content, "Courier", tagged: false, out _, resources: "/ExtGState << /GS1 << /CA 0.01 >> >>");

        Assert.DoesNotContain(page.Tables, table => table.Evidence.Any(evidence => evidence.Code == "table-ruling-lattice"));
    }

    [Fact]
    public void NonzeroFilledBoxInsideABoxHidesTheGridUnderIt() {
        // Two boxes wound the same way and filled with the nonzero rule paint the whole outer box: the inner one
        // is not a hole, so the grid it surrounds is painted over.
        StringBuilder content = GridUnderImage(out _,
            "q 1 g 40 600 m 300 600 l 300 720 l 40 720 l h 60 620 m 280 620 l 280 710 l 60 710 l h f Q\n" +
            "q 60 0 0 24 72 740 cm BI /W 1 /H 1 /CS /G /BPC 8 /F /AHx ID 80> EI Q\n");

        Assert.DoesNotContain(Read(content, "Courier").Tables,
            table => table.Evidence.Any(evidence => evidence.Code == "table-ruling-lattice"));
    }

    [Fact]
    public void SlantedImageCoversOnlyItsOwnStripNotItsBoundingBox() {
        // A thin opaque image slanted below the grid: its bounding box spans the grid, the image itself does not.
        StringBuilder content = GridUnderImage(out string[][] rows,
            "q 600 400 -1.11 1.66 0 400 cm BI /W 1 /H 1 /CS /G /BPC 8 /F /AHx ID 80> EI Q\n");

        PdfLogicalTable table = Assert.Single(Read(content, "Courier").Tables);
        Assert.Contains(table.Evidence, evidence => evidence.Code == "table-ruling-lattice");
        Assert.Equal(rows, table.Rows.Select(row => row.ToArray()));
    }

    private static readonly IEqualityComparer<string[]> ArrayComparer = new SequenceComparer();

    private sealed class SequenceComparer : IEqualityComparer<string[]> {
        public bool Equals(string[]? x, string[]? y) => x is null ? y is null : y is not null && x.SequenceEqual(y);
        public int GetHashCode(string[] obj) => obj.Length;
    }

    // A stroked grid whose rows are one string each, so only the borders split them into cells, with an image
    // painted after it. The image also keeps the strict stroke reader out.
    private static StringBuilder GridUnderImage(out string[][] rows, string image) {
        rows = new[] {
            new[] { "Code", "Item" },
            new[] { "A-01", "Widget" },
            new[] { "B-02", "Sprocket" },
            new[] { "C-03", "Flange" }
        };
        double[] x = { 72, 162, 262 };
        double[] y = Rows(700, 16, 16, rows.Length - 1);
        var content = new StringBuilder("0 G 0.75 w\n");
        foreach (double top in y) Line(content, x[0], top, x[x.Length - 1], top);
        foreach (double left in x) Line(content, left, y[0], left, y[y.Length - 1]);
        content.Append(image);
        for (int row = 0; row < rows.Length; row++) {
            double spacing = x[1] + 3 - (x[0] + 3 + rows[row][0].Length * 5.4) - 5.4;
            content.Append("BT /F1 9 Tf ").Append(Number(spacing)).Append(" Tw 0 g 1 0 0 1 ").Append(Number(x[0] + 3)).Append(' ')
                .Append(Number(y[row + 1] + 5)).Append(" Tm (").Append(rows[row][0]).Append(' ').Append(rows[row][1]).Append(") Tj ET\n");
        }
        return content;
    }

    [Fact]
    public void ClosedGridWithAWrappedHeaderKeepsItsBodyRows() {
        // Every header cell wraps to the same three lines; the body rows below are each ruled. The shared
        // header lines are wrapping, not data rows the producer left unruled.
        string[][] rows = {
            new[] { "Annual sales volume", "Average unit price", "Regional market share" },
            new[] { "1200", "4.50", "12%" },
            new[] { "900", "5.25", "9%" }
        };
        double[] x = { 72, 152, 232, 312 };
        double[] y = { 700, 664, 648, 632 };
        string[][] header = {
            new[] { "Annual", "sales", "volume" },
            new[] { "Average", "unit", "price" },
            new[] { "Regional", "market", "share" }
        };
        var content = new StringBuilder();
        content.Append("0 G 0.5 w\n");
        foreach (double row in y) Line(content, x[0], row, x[x.Length - 1], row);
        foreach (double column in x) Line(content, column, y[0], column, y[y.Length - 1]);
        for (int column = 0; column < header.Length; column++) {
            for (int line = 0; line < 3; line++) {
                content.Append("BT /F1 9 Tf 0 g 1 0 0 1 ").Append(Number(x[column] + 3)).Append(' ').Append(Number(y[0] - 11 - line * 10))
                    .Append(" Tm (").Append(header[column][line]).Append(") Tj ET\n");
            }
        }
        for (int row = 1; row < rows.Length; row++) {
            for (int column = 0; column < 3; column++) {
                content.Append("BT /F1 9 Tf 0 g 1 0 0 1 ").Append(Number(x[column] + 3)).Append(' ').Append(Number(y[row + 1] + 5))
                    .Append(" Tm (").Append(rows[row][column]).Append(") Tj ET\n");
            }
        }

        PdfLogicalTable table = Assert.Single(Read(content, "Helvetica").Tables);

        Assert.Equal(rows, table.Rows.Select(row => row.ToArray()));
    }

    [Fact]
    public void BoxWithoutRowRulesKeepsEveryRowSeparate() {
        // Column dividers and a header rule, but no rules between body rows: one ruled row holds every data
        // line, so the grid must not fold them into a single row.
        double[] x = { 72, 112, 272, 332, 392, 472 };
        double[] frame = { 700, 680, 584 };
        var content = new StringBuilder();
        Text(content, Market, x, Rows(700, 20, 16, Market.Length - 1), "0 g");
        content.Append("0 G 0.5 w\n");
        foreach (double rule in frame) Line(content, x[0], rule, x[5], rule);
        foreach (double divider in x) Line(content, divider, frame[0], divider, frame[2]);

        PdfLogicalTable table = Assert.Single(Read(content, "Helvetica").Tables);

        Assert.DoesNotContain(table.Evidence, evidence => evidence.Code == "table-ruling-lattice");
        Assert.Equal(Market, table.Rows.Select(row => row.ToArray()));
    }

    [Fact]
    public void LabelCentredInAVerticallyMergedCellBelongsToItsTopRow() {
        string[][] rows = {
            new[] { "Group", "Item", "Qty", "Price" },
            new[] { "Hardware", "Widget", "12", "1.50" },
            new[] { "", "Sprocket", "7", "2.25" },
            new[] { "Services", "Install", "1", "90.00" }
        };
        double[] x = { 72, 172, 272, 332, 412 };
        double[] y = Rows(700, 16, 16, rows.Length - 1);
        var content = new StringBuilder();
        var cells = rows.Select(row => row.ToArray()).ToArray();
        cells[1][0] = string.Empty;
        Text(content, cells, x, y, "0 g");
        // The merged label sits halfway down the two rows it spans, just inside the lower one.
        content.Append("BT /F1 9 Tf 0 g 1 0 0 1 75 ").Append(Number(y[2] - 6)).Append(" Tm (Hardware) Tj ET\n");
        content.Append("0 G 0.5 w\n");
        for (int row = 0; row < y.Length; row++) {
            // The rule between the two merged rows skips the first column.
            Line(content, row == 2 ? x[1] : x[0], y[row], x[x.Length - 1], y[row]);
        }
        foreach (double divider in x) Line(content, divider, y[0], divider, y[y.Length - 1]);

        PdfLogicalTable table = Assert.Single(Read(content, "Helvetica").Tables);

        Assert.Equal(rows, table.Rows.Select(row => row.ToArray()));
    }

    private static double[] Rows(double top, double header, double body, int bodyRows) {
        var y = new double[bodyRows + 2];
        y[0] = top;
        y[1] = top - header;
        for (int row = 2; row < y.Length; row++) y[row] = y[row - 1] - body;
        return y;
    }

    private static void Text(StringBuilder content, string[][] rows, double[] x, double[] y, string headerColor) {
        for (int row = 0; row < rows.Length; row++) {
            for (int column = 0; column < rows[row].Length; column++) {
                if (rows[row][column].Length == 0) continue;
                content.Append("BT /F1 9 Tf ").Append(row == 0 ? headerColor : "0 g").Append(" 1 0 0 1 ")
                    .Append(Number(x[column] + 3)).Append(' ').Append(Number(y[row + 1] + 5))
                    .Append(" Tm (").Append(rows[row][column]).Append(") Tj ET\n");
            }
        }
    }

    private static void Fill(StringBuilder content, string color, double left, double bottom, double width, double height) =>
        content.Append(color).Append(' ').Append(Number(left)).Append(' ').Append(Number(bottom)).Append(' ')
            .Append(Number(width)).Append(' ').Append(Number(height)).Append(" re f\n");

    private static void Line(StringBuilder content, double x1, double y1, double x2, double y2) =>
        content.Append(Number(x1)).Append(' ').Append(Number(y1)).Append(" m ")
            .Append(Number(x2)).Append(' ').Append(Number(y2)).Append(" l S\n");

    private static string Number(double value) => value.ToString("0.###", CultureInfo.InvariantCulture);

    private static PdfLogicalPage Read(StringBuilder content, string font) => Read(content, font, tagged: false, out _);

    // Tagged: a marked catalog whose Table element holds one paragraph over marked content 0.
    // Structure, when given, replaces it from object 6 (the StructTreeRoot) on.
    private static PdfLogicalPage Read(StringBuilder content, string font, bool tagged, out PdfDocumentReadResult read,
        string resources = "", IReadOnlyList<string>? structure = null) {
        string stream = content.ToString();
        var objects = new List<string> {
            tagged ? "<< /Type /Catalog /Pages 2 0 R /MarkInfo << /Marked true >> /StructTreeRoot 6 0 R >>" : "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 5 0 R >> " + resources + " >> /Contents 4 0 R" +
                (tagged ? " /StructParents 0" : string.Empty) + " >>",
            "<< /Length " + Encoding.ASCII.GetByteCount(stream).ToString(CultureInfo.InvariantCulture) + " >>\nstream\n" + stream + "endstream",
            // Producers write glyph widths into the font; Courier's are all 600.
            "<< /Type /Font /Subtype /Type1 /BaseFont /" + font + (font == "Courier"
                ? " /FirstChar 32 /LastChar 126 /Widths [" + string.Join(" ", Enumerable.Repeat("600", 95)) + "]" : string.Empty) + " >>"
        };
        if (tagged && structure is not null) {
            objects.AddRange(structure);
        } else if (tagged) {
            objects.Add("<< /Type /StructTreeRoot /K [7 0 R] /ParentTree << /Nums [0 [8 0 R]] >> >>");
            objects.Add("<< /Type /StructElem /S /Table /P 6 0 R /K [8 0 R] /Pg 3 0 R >>");
            objects.Add("<< /Type /StructElem /S /P /P 7 0 R /K 0 /Pg 3 0 R >>");
        }
        var builder = new StringBuilder("%PDF-1.7\n");
        var offsets = new List<int>();
        for (int index = 0; index < objects.Count; index++) {
            offsets.Add(Encoding.ASCII.GetByteCount(builder.ToString()));
            builder.Append(index + 1).Append(" 0 obj\n").Append(objects[index]).Append("\nendobj\n");
        }
        int xref = Encoding.ASCII.GetByteCount(builder.ToString());
        builder.Append("xref\n0 ").Append(objects.Count + 1).Append("\n0000000000 65535 f \n");
        foreach (int offset in offsets) builder.Append(offset.ToString("D10", CultureInfo.InvariantCulture)).Append(" 00000 n \n");
        builder.Append("trailer\n<< /Root 1 0 R /Size ").Append(objects.Count + 1).Append(" >>\nstartxref\n").Append(xref.ToString(CultureInfo.InvariantCulture)).Append("\n%%EOF\n");
        using PdfDocument document = PdfDocument.Load(Encoding.ASCII.GetBytes(builder.ToString()));
        read = document.Read();
        return Assert.Single(read.Pages);
    }
}
