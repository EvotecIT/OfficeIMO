using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using PdfCore = OfficeIMO.Pdf;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Adjacent_header_and_footer_separators_require_matching_edge_declarations(bool conflict) {
        using var package = DimensionPackage(IWorkDocumentKind.Keynote, rowCount: 2,
            firstBucketPayload: Message(VarintField(1, 1), BytesField(2, DimensionHeader(0, 22)), BytesField(2, DimensionHeader(1, 22))),
            secondBucketPayload: Message(VarintField(1, 1)), tableStyleReference: ReferenceField(3, 40),
            modelDeclarations: Message(VarintField(9, 1), VarintField(11, 1)),
            styleRecords: new[] { ArchiveRecord(40, 6003, BytesField(11, Message(
                BytesField(46, BorderStroke(2)), BytesField(54, BorderStroke(conflict ? 3 : 2))))) });
        var projection = IWorkSourceDocument.Open(package).ReadKeynote();
        if (conflict) {
            Assert.Contains(projection.Diagnostics, diagnostic => diagnostic.Code == "IWORK_TABLE_BORDERS_UNSUPPORTED");
            Assert.False(projection.HasEditableContent);
        } else {
            AssertBorder(projection.Slides[0].Tables[0].GetBorders(1, 1).Bottom, "000000", 2);
            AssertBorder(projection.Slides[0].Tables[0].GetBorders(2, 1).Top, "000000", 2);
        }
    }

    [Fact]
    public void Footer_only_table_retains_its_complete_outer_perimeter() {
        using var package = DimensionPackage(IWorkDocumentKind.Keynote, rowCount: 1,
            firstBucketPayload: Message(VarintField(1, 0), BytesField(2, DimensionHeader(0, 22))),
            secondBucketPayload: Message(VarintField(1, 1)), tableStyleReference: ReferenceField(3, 40),
            modelDeclarations: Message(VarintField(10, 1), VarintField(11, 1)),
            styleRecords: new[] { ArchiveRecord(40, 6003, BytesField(11, Message(BytesField(55, BorderStroke(2))))) });
        var table = IWorkSourceDocument.Open(package).ReadKeynote().Slides[0].Tables[0];
        for (int column = 1; column <= 2; column++) {
            AssertBorder(table.GetBorders(1, column).Top, "000000", 2);
            AssertBorder(table.GetBorders(1, column).Bottom, "000000", 2);
        }
        AssertBorder(table.GetBorders(1, 1).Left, "000000", 2);
        AssertBorder(table.GetBorders(1, 2).Right, "000000", 2);
    }

    public static IEnumerable<object[]> MergedBorderDestinations() {
        foreach (var kind in new[] { IWorkDocumentKind.Pages, IWorkDocumentKind.Numbers, IWorkDocumentKind.Keynote })
            foreach (var span in new[] { (1, 2), (3, 1), (3, 2) }) yield return new object[] { kind, span.Item1, span.Item2 };
    }

    [Theory]
    [MemberData(nameof(MergedBorderDestinations))]
    public void Saved_merged_cells_retain_perimeter_edges_instead_of_interior_grid(IWorkDocumentKind kind, int rows, int columns) {
        using var package = MergedBorderPackage(kind, rows, columns);
        var source = IWorkSourceDocument.Open(package);
        using var saved = new MemoryStream();
        if (kind == IWorkDocumentKind.Pages) {
            using var result = source.ToWordDocumentResult(); result.Report.RequireCompleteEditableReconstruction();
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = WordDocument.Load(saved); var table = reopened.Tables[0];
            Assert.Equal(columns == 2 ? 16u : 8u, table.Rows[0].Cells[0].Borders.RightSize);
            Assert.Equal(rows == 3 ? 16u : 8u, table.Rows[rows - 1].Cells[0].Borders.BottomSize);
            Assert.Equal(16u, table.Rows[0].Cells[0].Borders.TopSize);
            if (rows > 1) Assert.Equal(WordBorderStyle.Nil, table.Rows[0].Cells[0].Borders.BottomStyle);
        } else if (kind == IWorkDocumentKind.Numbers) {
            using var result = source.ToExcelDocumentResult(); result.Report.RequireCompleteEditableReconstruction();
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = ExcelDocument.Load(saved); var borders = reopened.Sheets[0].GetCellStyle(1, 1).Border!;
            Assert.Equal(columns == 2 ? "medium" : "thin", borders.Right!.Style);
            Assert.Equal(rows == 3 ? "medium" : "thin", borders.Bottom!.Style);
        } else {
            using var result = source.ToPowerPointDocumentResult(); result.Report.RequireCompleteEditableReconstruction();
            // Border evidence does not depend on the host's minimum font line box.
            var authoredTable = result.Value.Slides[0].Tables.First();
            for (int row = 0; row < authoredTable.Rows; row++) authoredTable.SetRowHeightPoints(row, 60D);
            authoredTable.Height = OfficeLength.Points(authoredTable.Rows * 60D);
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = PowerPointDocument.Load(saved); Assert.Empty(reopened.ValidateDocument());
            var pdf = reopened.ToPdf();
            var rendered = pdf.Render.Pages("1", new PdfCore.PdfPageRenderOptions {
                Format = global::OfficeIMO.Drawing.OfficeImageExportFormat.Svg, ContinueOnError = false
            }).Single();
            Assert.True(rendered.Succeeded);
            string svg = System.Text.Encoding.UTF8.GetString(rendered.Bytes!);
            var target = reopened.Slides[0].Tables.First();
            double right = target.Left.ToPoints() + Enumerable.Range(0, columns).Sum(target.GetColumnWidthPoints);
            double bottom = target.Top.ToPoints() + Enumerable.Range(0, rows).Sum(target.GetRowHeightPoints);
            AssertRenderedBorder(svg, right, target.Top.ToPoints(), right, bottom, columns == 2 ? 2d : 1d);
            AssertRenderedBorder(svg, target.Left.ToPoints(), bottom, right, bottom, rows == 3 ? 2d : 1d);
            saved.Position = 0; using var zip = new ZipArchive(saved, ZipArchiveMode.Read, leaveOpen: true);
            using var xmlStream = zip.Entries.Single(entry => entry.FullName.StartsWith("ppt/slides/slide", StringComparison.Ordinal)
                && entry.FullName.EndsWith(".xml", StringComparison.Ordinal)).Open();
            var xml = System.Xml.Linq.XDocument.Load(xmlStream); System.Xml.Linq.XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            var anchor = xml.Descendants(a + "tcPr").First();
            Assert.Equal(columns == 2 ? "25400" : "12700", (string?)anchor.Element(a + "lnR")!.Attribute("w"));
            Assert.Equal(rows == 3 ? "25400" : "12700", (string?)anchor.Element(a + "lnB")!.Attribute("w"));
        }
    }

    private static void AssertRenderedBorder(string svg, double x1, double y1, double x2, double y2, double width) {
        var document = System.Xml.Linq.XDocument.Parse(svg);
        Assert.True(document.Descendants().Any(element => {
            if (element.Name.LocalName != "line" || !double.TryParse((string?)element.Attribute("stroke-width"),
                System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double actualWidth)
                || Math.Abs(actualWidth - width) > 0.001) return false;
            var values = new[] { "x1", "y1", "x2", "y2" }.Select(name =>
                double.Parse((string)element.Attribute(name)!, System.Globalization.CultureInfo.InvariantCulture)).ToArray();
            bool At(int index, double x, double y) => Math.Abs(values[index] - x) < 0.001 && Math.Abs(values[index + 1] - y) < 0.001;
            return At(0, x1, y1) && At(2, x2, y2) || At(0, x2, y2) && At(2, x1, y1);
        }), $"Missing border ({x1},{y1})-({x2},{y2}) at width {width}.\n{svg}");
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Varying_merged_perimeter_requires_partial_policy_and_reports_approximation(IWorkDocumentKind kind) {
        using var package = MergedBorderPackage(kind, 1, 2, varyingTop: true);
        var source = IWorkSourceDocument.Open(package);
        var strict = new IWorkConversionOptions { Mode = IWorkConversionMode.EditableOnly };
        Assert.Contains("merged", Assert.Throws<InvalidDataException>(() => {
            if (kind == IWorkDocumentKind.Pages) { using var result = source.ToWordDocumentResult(strict); }
            else if (kind == IWorkDocumentKind.Numbers) { using var result = source.ToExcelDocumentResult(strict); }
            else { using var result = source.ToPowerPointDocumentResult(strict); }
        }).Message, StringComparison.OrdinalIgnoreCase);
        var policy = new IWorkConversionOptions { AllowPartialEditableReconstruction = true };
        IWorkConversionReport report;
        if (kind == IWorkDocumentKind.Pages) { using var result = source.ToWordDocumentResult(policy); report = result.Report; }
        else if (kind == IWorkDocumentKind.Numbers) { using var result = source.ToExcelDocumentResult(policy); report = result.Report; }
        else { using var result = source.ToPowerPointDocumentResult(policy); report = result.Report; }
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code.EndsWith("MERGED_TABLE_BORDERS_APPROXIMATED", StringComparison.Ordinal)
            && diagnostic.LossKind == OfficeConversionLossKind.Approximation);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Saved_Excel_merge_clears_interior_border_when_its_perimeter_is_undeclared(bool horizontalEdge) {
        using var package = MergedBorderPackage(IWorkDocumentKind.Numbers, horizontalEdge ? 3 : 1, horizontalEdge ? 1 : 2,
            missingOuterVertical: !horizontalEdge, missingOuterHorizontal: horizontalEdge);
        using var result = IWorkSourceDocument.Open(package).ToExcelDocumentResult();
        result.Report.RequireCompleteEditableReconstruction();
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        var borders = reopened.Sheets[0].GetCellStyle(1, 1).Border!;
        if (horizontalEdge) {
            Assert.True(borders.Bottom?.Style is null or "none");
            Assert.Equal("medium", borders.Left!.Style);
        } else {
            Assert.True(borders.Right?.Style is null or "none");
            Assert.Equal("medium", borders.Top!.Style);
        }
    }

    private static MemoryStream MergedBorderPackage(IWorkDocumentKind kind, int rows, int columns, bool varyingTop = false,
        bool missingOuterVertical = false, bool missingOuterHorizontal = false) {
        var records = new List<byte[]> { ArchiveRecord(40, 6003, BytesField(11, Message(
            missingOuterHorizontal ? Array.Empty<byte>() : BytesField(58, BorderStroke(2)),
            missingOuterVertical ? Array.Empty<byte>() : BytesField(59, BorderStroke(2)),
            BytesField(60, BorderStroke(1)), BytesField(61, BorderStroke(1))))) };
        if (varyingTop) {
            records.Add(ArchiveRecord(60, 6305, Message(VarintField(1, 1), VarintField(2, 2), VarintField(3, 3), ReferenceField(6, 61))));
            records.Add(ArchiveRecord(61, 6306, Message(VarintField(1, 0), BytesField(2, Message(
                VarintField(1, 1), VarintField(2, 1), BytesField(3, BorderStroke(3, 1, 0, 0)), VarintField(4, 1))))));
        }
        return DimensionPackage(kind, tableStyleReference: ReferenceField(3, 40), styleRecords: records.ToArray(),
            modelDeclarations: Message(MergeTestOwner(MergeTestPair(0, 0, (ulong)rows - 1, (ulong)columns - 1)),
                varyingTop ? ReferenceField(49, 60) : Array.Empty<byte>()));
    }
}
