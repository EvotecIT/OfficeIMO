using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    public static IEnumerable<object[]> TablePreferredWidthDefaultCases() {
        var widths = new (string? Type, string? Value, bool Omitted, double ColumnWidth)[] {
            (null, null, true, 72), ("auto", "0", false, 72), ("nil", "0", false, 72),
            ("dxa", "4320", false, 108), ("dxa", "0", false, 72), ("pct", "0", false, 72),
            (null, "0", false, 72), (null, "4320", false, 108),
            ("nil", "4320", false, 72), ("auto", "4320", false, 72),
            (null, null, false, 72), ("dxa", null, false, 72), ("pct", null, false, 72),
            ("dxa", "50%", false, 116.906), ("pct", "50%", false, 116.906),
            ("auto", "50%", false, 116.906), ("nil", "50%", false, 116.906),
            (null, "50%", false, 116.906), ("pct", "0%", false, 72)
        };
        foreach (WordTableLayoutMode? layout in new WordTableLayoutMode?[] { null, WordTableLayoutMode.AutoFit, WordTableLayoutMode.Fixed }) {
            foreach (var width in widths) {
                yield return new object[] { layout!, width.Type!, width.Value!, width.Omitted, width.ColumnWidth };
            }
        }
    }

    [Theory]
    [MemberData(nameof(TablePreferredWidthDefaultCases))]
    public void SaveAsPdf_TablePreferredWidthDefaultsPreserveTheSourceFrame(
        WordTableLayoutMode? layout, string? widthType, string? widthValue, bool omitWidth, double expectedColumnWidth) {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(1, 2);
        table.GridColumnWidth = new List<int> { 1440, 1440 };
        if (layout.HasValue) table.LayoutMode = layout.Value;
        else table._tableProperties!.TableLayout?.Remove();
        if (omitWidth) table._tableProperties!.TableWidth?.Remove();
        else {
            var width = new W.TableWidth { Width = widthValue };
            if (widthType != null) {
                width.Type = widthType switch {
                    "auto" => W.TableWidthUnitValues.Auto,
                    "nil" => W.TableWidthUnitValues.Nil,
                    "pct" => W.TableWidthUnitValues.Pct,
                    _ => W.TableWidthUnitValues.Dxa
                };
            }
            table._tableProperties!.TableWidth = width;
        }
        for (int index = 0; index < 2; index++) {
            WordTableCell cell = table.Rows[0].Cells[index];
            cell.Width = 1440;
            cell.WidthType = WordTableWidthUnit.Dxa;
            cell.MarginLeftWidth = cell.MarginRightWidth = 108;
            cell.MarginTopWidth = cell.MarginBottomWidth = 0;
            WordParagraph paragraph = cell.Paragraphs[0];
            paragraph.Text = index == 0 ? "Left" : "Right";
            ConfigurePartialColumnParagraph(paragraph);
        }
        using var input = new MemoryStream(document.ToBytes(new WordSaveOptions { Format = WordFileFormat.Docx }));
        using WordDocument loaded = WordDocument.Load(input);
        using PdfPigDocument pdf = PdfPigDocument.Open(loaded.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, PageSize = new OfficeIMO.Pdf.PdfPageSize(612, 792),
            Margins = OfficeIMO.Pdf.PdfPageMargins.Uniform(72)
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(1, pdf.NumberOfPages);
        var words = pdf.GetPage(1).GetWords().ToArray();
        Assert.Equal(2, words.Length);
        var left = Assert.Single(words, word => word.Text == "Left");
        var right = Assert.Single(words, word => word.Text == "Right");
        Assert.InRange(right.BoundingBox.Left - left.BoundingBox.Left, expectedColumnWidth - .1, expectedColumnWidth + .1);
    }
}
