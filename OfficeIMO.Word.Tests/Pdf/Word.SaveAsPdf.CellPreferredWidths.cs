using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void SaveAsPdf_AutomaticMergedCellGrowsBeyondAbsoluteCellsInAnotherRow() {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateAutomaticWidthControl(document, 1440, 1440);
        table.LayoutMode = WordTableLayoutMode.AutoFit;
        table._tableProperties!.TableWidth = new W.TableWidth { Type = W.TableWidthUnitValues.Auto, Width = "0" };
        foreach (WordTableCell cell in table.Rows[0].Cells) {
            cell.Paragraphs[0].Text = "S";
            ConfigurePartialColumnParagraph(cell.Paragraphs[0]);
        }
        WordTableRow row = table.AddRow();
        row.Cells[0].MergeHorizontally(1);
        row.Cells[0]._tableCellProperties!.TableCellWidth = new W.TableCellWidth { Type = W.TableWidthUnitValues.Auto, Width = "0" };
        row.Cells[0].Paragraphs[0].Text = string.Join(" ", Enumerable.Repeat("AAA BBB CCCC", 8));
        ConfigurePartialColumnParagraph(row.Cells[0].Paragraphs[0]);
        var fonts = new OfficeIMO.Pdf.PdfCreateOptions();
        fonts.RegisterNamedFontFamily(new OfficeIMO.Pdf.PdfEmbeddedFontFamily("Arial",
            OfficeIMO.TestAssets.ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 600, Enumerable.Range(32, 95).ToArray())));
        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false, PdfOptions = fonts }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var letters = pdf.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
        Assert.Equal("SS" + string.Concat(Enumerable.Repeat("AAABBBCCCC", 8)), string.Concat(letters.Select(letter => letter.Value)));
        Assert.Equal(3, letters.Select(letter => Math.Round(letter.StartBaseLine.Y, 2)).Distinct().Count());
    }

    [Theory]
    [InlineData(WordTableLayoutMode.Fixed, 40, 102.351D)]
    [InlineData(WordTableLayoutMode.Fixed, 120, 91.054D)]
    [InlineData(WordTableLayoutMode.AutoFit, 40, 113D)]
    [InlineData(WordTableLayoutMode.AutoFit, 120, 121D)]
    public void SaveAsPdf_SpacedCellWidthsRetainPhysicalMinimumsAndAbsoluteNoWrapAllocation(
        WordTableLayoutMode layout, short spacing, double expectedGap) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateAutomaticWidthControl(document, 2160, 720);
        table.LayoutMode = layout;
        table._tableProperties!.TableWidth = new W.TableWidth { Type = W.TableWidthUnitValues.Dxa, Width = "2880" };
        table.StyleDetails!.CellSpacing = spacing;
        for (int column = 0; column < 2; column++) {
            WordTableCell cell = table.Rows[0].Cells[column];
            cell.Paragraphs[0].Text = column == 0 ? "AAA BBB" : "CCCC DDDD EEEE";
            ConfigurePartialColumnParagraph(cell.Paragraphs[0]);
            cell.WrapText = false;
        }
        var fonts = new OfficeIMO.Pdf.PdfCreateOptions();
        fonts.RegisterNamedFontFamily(new OfficeIMO.Pdf.PdfEmbeddedFontFamily("Arial",
            OfficeIMO.TestAssets.ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 600, Enumerable.Range(32, 95).ToArray())));
        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false, PdfOptions = fonts }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var letters = pdf.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
        Assert.Equal("AAABBBCCCCDDDDEEEE", string.Concat(letters.Select(letter => letter.Value)));
        double gap = letters.First(letter => letter.Value == "C").StartBaseLine.X - letters[0].StartBaseLine.X;
        Assert.InRange(gap, expectedGap - .05D, expectedGap + .05D);
    }

    [Theory]
    [InlineData(WordTableLayoutMode.AutoFit, 3)]
    [InlineData(WordTableLayoutMode.Fixed, 10)]
    public void SaveAsPdf_MergedZeroCellWidthDiscardsUnconstrainedCachedGrid(WordTableLayoutMode layout, int expectedLines) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateAutomaticWidthControl(document, 1440, 1440);
        table.LayoutMode = layout;
        table._tableProperties!.TableWidth = new W.TableWidth { Type = W.TableWidthUnitValues.Auto, Width = "0" };
        table.Rows[0].Cells[0].MergeHorizontally(1);
        WordParagraph paragraph = table.Rows[0].Cells[0].Paragraphs[0];
        paragraph.Text = "AAA BBB CCCC";
        ConfigurePartialColumnParagraph(paragraph);
        using var input = new MemoryStream();
        byte[] bytes = document.ToBytes(new WordSaveOptions { Format = WordFileFormat.Docx });
        input.Write(bytes, 0, bytes.Length);
        input.Position = 0;
        using (var source = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(input, true)) {
            source.MainDocumentPart!.Document.Body!.Elements<W.Table>().Single()
                .Elements<W.TableRow>().Single().Elements<W.TableCell>().Single().TableCellProperties!.TableCellWidth =
                new W.TableCellWidth { Type = W.TableWidthUnitValues.Dxa, Width = "0" };
            source.MainDocumentPart.Document.Save();
        }
        input.Position = 0;
        using WordDocument loaded = WordDocument.Load(input);
        var fonts = new OfficeIMO.Pdf.PdfCreateOptions();
        fonts.RegisterNamedFontFamily(new OfficeIMO.Pdf.PdfEmbeddedFontFamily("Arial",
            OfficeIMO.TestAssets.ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 600, Enumerable.Range(32, 95).ToArray())));
        using var pdf = PdfPigDocument.Open(loaded.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false, PdfOptions = fonts }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var letters = pdf.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
        Assert.Equal("AAABBBCCCC", string.Concat(letters.Select(letter => letter.Value)));
        Assert.Equal(expectedLines, letters.Select(letter => Math.Round(letter.StartBaseLine.Y, 2)).Distinct().Count());
    }

    [Theory]
    [InlineData(null, "2160", "720", false, 108D, 3)]
    [InlineData("dxa", "2160", "720", false, 108D, 3)]
    [InlineData("pct", "3750", "1250", false, 61.201D, 1)]
    [InlineData("pct", "75%", "25%", false, 61.201D, 1)]
    [InlineData("dxa", "75%", "25%", false, 61.201D, 1)]
    [InlineData(null, "75%", "25%", false, 61.201D, 1)]
    [InlineData("dxa", "0", "0", true, 32.401D, 3)]
    [InlineData(null, "0", "0", true, 32.401D, 3)]
    [InlineData("auto", null, null, true, 32.401D, 3)]
    [InlineData("nil", "2160", "720", false, 72D, 3)]
    [InlineData("nil", "0", "0", false, 72D, 3)]
    public void SaveAsPdf_NoWrapCellWidthsUseEffectiveUnitsAndRetainAbsoluteAllocation(
        string? type, string? leftWidth, string? rightWidth, bool automaticFrame, double expectedWidth, int expectedLines) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateAutomaticWidthControl(document, 1440, 1440);
        table.LayoutMode = WordTableLayoutMode.AutoFit;
        table._tableProperties!.TableWidth = new W.TableWidth {
            Type = automaticFrame ? W.TableWidthUnitValues.Auto : W.TableWidthUnitValues.Dxa,
            Width = automaticFrame ? "0" : "2880"
        };
        for (int column = 0; column < 2; column++) {
            WordTableCell cell = table.Rows[0].Cells[column];
            cell.Paragraphs[0].Text = column == 0 ? "AAA BBB" : "CCCC DDDD EEEE";
            ConfigurePartialColumnParagraph(cell.Paragraphs[0]);
            cell.WrapText = false;
        }
        using var input = new MemoryStream();
        byte[] bytes = document.ToBytes(new WordSaveOptions { Format = WordFileFormat.Docx });
        input.Write(bytes, 0, bytes.Length);
        input.Position = 0;
        using (var source = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(input, true)) {
            var cells = source.MainDocumentPart!.Document.Body!.Elements<W.Table>().Single()
                .Elements<W.TableRow>().First().Elements<W.TableCell>().ToArray();
            for (int column = 0; column < cells.Length; column++) {
                var width = new W.TableCellWidth { Width = column == 0 ? leftWidth : rightWidth };
                if (type != null) width.Type = type switch {
                    "pct" => W.TableWidthUnitValues.Pct, "auto" => W.TableWidthUnitValues.Auto,
                    "nil" => W.TableWidthUnitValues.Nil, _ => W.TableWidthUnitValues.Dxa
                };
                cells[column].TableCellProperties!.TableCellWidth = width;
            }
            source.MainDocumentPart.Document.Save();
        }
        input.Position = 0;
        using WordDocument loaded = WordDocument.Load(input);
        var fonts = new OfficeIMO.Pdf.PdfCreateOptions();
        fonts.RegisterNamedFontFamily(new OfficeIMO.Pdf.PdfEmbeddedFontFamily("Arial",
            OfficeIMO.TestAssets.ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 600, Enumerable.Range(32, 95).ToArray())));
        using var pdf = PdfPigDocument.Open(loaded.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false, PdfOptions = fonts }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var letters = pdf.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
        Assert.Equal("AAABBBCCCCDDDDEEEE", string.Concat(letters.Select(letter => letter.Value)));
        double actual = letters.First(letter => letter.Value == "C").StartBaseLine.X - letters[0].StartBaseLine.X;
        Assert.InRange(actual, expectedWidth - .05D, expectedWidth + .05D);
        Assert.Equal(expectedLines, letters.Select(letter => Math.Round(letter.StartBaseLine.Y, 2)).Distinct().Count());
    }

    [Theory]
    [InlineData(WordTableLayoutMode.AutoFit, 2, 75.603D)]
    [InlineData(WordTableLayoutMode.Fixed, 2, 35.4D)]
    [InlineData(WordTableLayoutMode.AutoFit, 3, 25.201D)]
    [InlineData(WordTableLayoutMode.Fixed, 3, 11.8D)]
    public void SaveAsPdf_MergedPercentageCellsRetainIndependentColumnPreferences(WordTableLayoutMode layout, int columns, double expectedGap) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateAutomaticWidthControl(document, Enumerable.Repeat(1440, columns).ToArray());
        table.LayoutMode = layout;
        table._tableProperties!.TableWidth = new W.TableWidth { Type = W.TableWidthUnitValues.Auto, Width = "0" };
        int[] percentages = columns == 2 ? new[] { 75, 25 } : new[] { 25, 50, 25 };
        for (int column = 0; column < columns; column++) {
            WordTableCell cell = table.Rows[0].Cells[column];
            cell.Paragraphs[0].Text = "S" + (char)('A' + column);
            ConfigurePartialColumnParagraph(cell.Paragraphs[0]);
            cell._tableCellProperties!.TableCellWidth = new W.TableCellWidth {
                Type = W.TableWidthUnitValues.Pct, Width = (percentages[column] * 50).ToString()
            };
        }
        WordTableRow mergedRow = table.AddRow();
        mergedRow.Cells[0].MergeHorizontally(1);
        for (int column = 0; column < mergedRow.Cells.Count; column++) {
            WordTableCell cell = mergedRow.Cells[column];
            cell.Paragraphs[0].Text = column == 0 ? "Top" : "SB";
            ConfigurePartialColumnParagraph(cell.Paragraphs[0]);
            cell._tableCellProperties!.TableCellWidth = new W.TableCellWidth {
                Type = W.TableWidthUnitValues.Pct, Width = column == 0 ? columns == 2 ? "5000" : "3750" : "1250"
            };
        }
        var fonts = new OfficeIMO.Pdf.PdfCreateOptions();
        fonts.RegisterNamedFontFamily(new OfficeIMO.Pdf.PdfEmbeddedFontFamily("Arial",
            OfficeIMO.TestAssets.ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 600, Enumerable.Range(32, 95).ToArray())));
        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false, PdfOptions = fonts }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var starts = pdf.GetPage(1).Letters.TakeWhile(letter => letter.Value != "T")
            .Where(letter => letter.Value == "S").ToArray();
        Assert.Equal(columns, starts.Length);
        Assert.InRange(starts[1].StartBaseLine.X - starts[0].StartBaseLine.X, expectedGap - .05D, expectedGap + .05D);
    }

    [Theory]
    [InlineData(60D)]
    [InlineData(120D)]
    [InlineData(240D)]
    public void SaveAsPdf_NarrowPercentageFramesPreservePaddingWhileCompressingUnbreakableText(double frameWidth) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateAutomaticWidthControl(document, 1440, 1440);
        table.LayoutMode = WordTableLayoutMode.AutoFit;
        table._tableProperties!.TableWidth = new W.TableWidth { Type = W.TableWidthUnitValues.Auto, Width = "0" };
        string left = new('A', 27), right = new('B', 24);
        for (int column = 0; column < 2; column++) {
            WordTableCell cell = table.Rows[0].Cells[column];
            cell.Paragraphs[0].Text = column == 0 ? left : right;
            ConfigurePartialColumnParagraph(cell.Paragraphs[0]);
            cell._tableCellProperties!.TableCellWidth = new W.TableCellWidth {
                Type = W.TableWidthUnitValues.Pct, Width = column == 0 ? "3750" : "1250"
            };
        }
        var fonts = new OfficeIMO.Pdf.PdfCreateOptions();
        fonts.RegisterNamedFontFamily(new OfficeIMO.Pdf.PdfEmbeddedFontFamily("Arial",
            OfficeIMO.TestAssets.ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 600, Enumerable.Range(32, 95).ToArray())));
        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, PdfOptions = fonts,
            PageSize = new OfficeIMO.Pdf.PdfPageSize(OfficeIMO.OfficeLength.Points(frameWidth + 144D),OfficeIMO.OfficeLength.Points(792D)), Margins = OfficeIMO.Pdf.PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(72D))
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var letters = pdf.GetPage(1).Letters.ToArray();
        Assert.Equal(left + right, string.Concat(letters.Select(letter => letter.Value)));
        // Word's outer half-borders reserve half a point. Cell padding remains
        // 5.4 points on each side while only the text advance is compressed.
        double expected = 10.8D + left.Length / (double)(left.Length + right.Length) * (frameWidth - .5D - 21.6D);
        double actual = letters.First(letter => letter.Value == "B").StartBaseLine.X - letters.First(letter => letter.Value == "A").StartBaseLine.X;
        Assert.InRange(actual, expected - .05D, expected + .05D);
    }

    public static IEnumerable<object[]> CellPreferredWidthCases() {
        var forms = new (string? Type, string? Left, string? Right, string Kind)[] {
            (null, "2160", "720", "absolute"),
            ("nil", "0", "0", "grid"), ("nil", "2160", "720", "grid"),
            ("auto", "0", "0", "content"), ("auto", "2160", "720", "content"),
            ("omit", null, null, "content"), (null, null, null, "zero"),
            (null, "0", "0", "zero"), ("dxa", "0", "0", "zero"),
            ("dxa", null, null, "zero"), ("auto", null, null, "zero"),
            ("pct", "3750", "1250", "percentage"),
            ("pct", "75%", "25%", "percentage"), ("dxa", "75%", "25%", "percentage")
        };
        foreach (WordTableLayoutMode layout in new[] { WordTableLayoutMode.AutoFit, WordTableLayoutMode.Fixed }) {
            foreach (bool automaticFrame in new[] { false, true }) {
                foreach (var form in forms) {
                    double expected = form.Kind == "absolute" ? 108D : form.Kind == "grid" ? 72D :
                        form.Kind == "zero" && automaticFrame && layout == WordTableLayoutMode.Fixed ? 11.8D :
                        form.Kind == "percentage" ? automaticFrame
                            ? layout == WordTableLayoutMode.Fixed ? 35.4D : 75.603D : 108D :
                        automaticFrame && layout == WordTableLayoutMode.AutoFit ? 25.201D : 72D;
                    int lines = (form.Kind == "percentage" || form.Kind == "zero") && automaticFrame && layout == WordTableLayoutMode.Fixed ? 2 : 1;
                    yield return new object[] { layout, automaticFrame, form.Type!, form.Left!, form.Right!, expected, lines };
                }
            }
            foreach (var ratio in new[] { (25, 75), (50, 50), (50, 25), (25, 25), (75, 75) }) {
                double minimum = layout == WordTableLayoutMode.Fixed ? 11.8D : 25.201D;
                double width = Math.Max(2D * minimum, minimum / (Math.Min(ratio.Item1, ratio.Item2) / 100D));
                double first = ratio.Item1 + ratio.Item2 > 100 ? minimum : width * ratio.Item1 / (ratio.Item1 + ratio.Item2);
                yield return new object[] { layout, true, "pct", (ratio.Item1 * 50).ToString(), (ratio.Item2 * 50).ToString(),
                    first, layout == WordTableLayoutMode.Fixed ? 2 : 1 };
            }
        }
    }

    [Theory]
    [MemberData(nameof(CellPreferredWidthCases))]
    public void SaveAsPdf_CellPreferredWidthsPreserveTypedMeasurementAndAutomaticFrames(
        WordTableLayoutMode layout, bool automaticFrame, string? type, string? leftWidth, string? rightWidth,
        double expectedColumnWidth, int expectedLines) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateAutomaticWidthControl(document, 1440, 1440);
        table.LayoutMode = layout;
        table._tableProperties!.TableWidth = new W.TableWidth {
            Type = automaticFrame ? W.TableWidthUnitValues.Auto : W.TableWidthUnitValues.Dxa,
            Width = automaticFrame ? "0" : "2880"
        };
        for (int index = 0; index < 2; index++) {
            WordTableCell cell = table.Rows[0].Cells[index];
            WordParagraph paragraph = cell.Paragraphs[0];
            paragraph.Text = index == 0 ? "SA" : "SB";
            ConfigurePartialColumnParagraph(paragraph);
        }
        var fonts = new OfficeIMO.Pdf.PdfCreateOptions();
        fonts.RegisterNamedFontFamily(new OfficeIMO.Pdf.PdfEmbeddedFontFamily("Arial",
            OfficeIMO.TestAssets.ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 600, Enumerable.Range(32, 95).ToArray())));
        using var input = new MemoryStream();
        byte[] bytes = document.ToBytes(new WordSaveOptions { Format = WordFileFormat.Docx });
        input.Write(bytes, 0, bytes.Length);
        input.Position = 0;
        // Model an independently produced input after the author's online-grid normalization.
        using (var source = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(input, true)) {
            var cells = source.MainDocumentPart!.Document.Body!.Elements<W.Table>().Single()
                .Elements<W.TableRow>().First().Elements<W.TableCell>().ToArray();
            for (int index = 0; index < cells.Length; index++) {
                if (type == "omit") cells[index].TableCellProperties!.TableCellWidth?.Remove();
                else {
                    var width = new W.TableCellWidth { Width = index == 0 ? leftWidth : rightWidth };
                    if (type != null) width.Type = type switch {
                        "nil" => W.TableWidthUnitValues.Nil, "auto" => W.TableWidthUnitValues.Auto,
                        "pct" => W.TableWidthUnitValues.Pct, _ => W.TableWidthUnitValues.Dxa
                    };
                    cells[index].TableCellProperties!.TableCellWidth = width;
                }
            }
            source.MainDocumentPart.Document.Save();
        }
        input.Position = 0;
        using WordDocument loaded = WordDocument.Load(input);
        using PdfPigDocument pdf = PdfPigDocument.Open(loaded.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, PdfOptions = fonts
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(1, pdf.NumberOfPages);
        var letters = pdf.GetPage(1).Letters.ToArray();
        Assert.Equal("SASB", string.Concat(letters.Select(letter => letter.Value)));
        var starts = letters.Where(letter => letter.Value == "S").ToArray();
        Assert.Equal(2, starts.Length);
        Assert.InRange(starts[1].StartBaseLine.X - starts[0].StartBaseLine.X,
            expectedColumnWidth - .05D, expectedColumnWidth + .05D);
        Assert.Equal(expectedLines, letters.Select(letter => Math.Round(letter.StartBaseLine.Y, 2)).Distinct().Count());
    }
}
