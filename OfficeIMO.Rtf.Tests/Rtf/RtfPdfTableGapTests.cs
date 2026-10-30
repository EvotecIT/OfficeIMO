using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfPdfTableGapTests {
    [Fact]
    public void BorderlessExplicitGapTableUsesSourceInsetsAndPreservesAllText() {
        using PdfPigDocument zero = PdfPigDocument.Open(ReadNativeGapFixture(0).ToPdf(Options()).ToBytes());
        using PdfPigDocument positive = PdfPigDocument.Open(ReadNativeGapFixture(300).ToPdf(Options()).ToBytes());
        Assert.Equal(1, zero.NumberOfPages);
        Assert.Equal(1, positive.NumberOfPages);
        // Clipping paths alone do not paint table borders.
        Assert.DoesNotContain(zero.GetPage(1).Paths, path => path.IsStroked || path.IsFilled);
        Assert.DoesNotContain(positive.GetPage(1).Paths, path => path.IsStroked || path.IsFilled);
        var zeroWords = zero.GetPage(1).GetWords().ToArray();
        var positiveWords = positive.GetPage(1).GetWords().ToArray();
        var zeroStart = Assert.Single(zeroWords, word => word.Text == "TABLE-MIN-A");
        var positiveStart = Assert.Single(positiveWords, word => word.Text == "TABLE-MIN-A");
        Assert.InRange(zeroStart.Letters.First().StartBaseLine.X, 71.99D, 72.01D);
        Assert.InRange(positiveStart.Letters.First().StartBaseLine.X - zeroStart.Letters.First().StartBaseLine.X, 14.99D, 15.01D);
        foreach (string marker in new[] { "TABLE-BEFORE", "TABLE-AFTER" }) {
            var zeroMarker = Assert.Single(zeroWords, word => word.Text == marker);
            var positiveMarker = Assert.Single(positiveWords, word => word.Text == marker);
            Assert.InRange(positiveMarker.Letters.First().StartBaseLine.X - zeroMarker.Letters.First().StartBaseLine.X, -0.01D, 0.01D);
        }
        // This portable profile deliberately uses Helvetica. Native Arial line
        // fitting is qualified separately with the matching font resources.
        const string expected = "TABLE-BEFORETABLE-MIN-ATABLE-MIN-TALLEND-TABLE-MINTABLE-MIN-BTABLE-EXACT-ATABLE-EXACT-TALLEND-TABLE-EXACTTABLE-EXACT-BTABLE-AFTER";
        Assert.Equal(expected, WithoutWhiteSpace(string.Concat(zero.GetPage(1).Letters.Select(letter => letter.Value))));
        Assert.Equal(expected, WithoutWhiteSpace(string.Concat(positive.GetPage(1).Letters.Select(letter => letter.Value))));
    }

    [Fact]
    public void ExplicitCellBorderAndPaddingOverridePositiveGap() {
        const int gapTwips = 300;
        RtfDocument source = ReadTable(@"\clbrdrl\brdrs\brdrw8\brdrcf1\clpadl240\clpadfl3", gapTwips);
        PdfCore.PdfDocument mapped = source.ToPdf(Options());
        PdfCore.TableBlock table = Assert.Single(mapped.Blocks.OfType<PdfCore.TableBlock>());
        Assert.Null(table.Style!.BorderColor);
        Assert.Equal(12D, table.Style.CellPaddings![(0, 0)].Left);
        PdfCore.PdfCellBorder border = table.Style.CellBorders![(0, 0)];
        Assert.True(border.Left);
        Assert.Equal(global::OfficeIMO.OfficeColor.FromRgb(255, 0, 0), border.LeftBorder!.Color);
        Assert.True(border.LeftBorder.Width > 0D);

        using PdfPigDocument pdf = PdfPigDocument.Open(mapped.ToBytes());
        Assert.Contains(pdf.GetPage(1).Paths, path => path.IsStroked || path.IsFilled);
        Assert.Single(pdf.GetPage(1).Letters, letter => letter.Value == "A");
    }

    [Fact]
    public void CallerTableStyleOverridesPositiveGapWithoutMutation() {
        const int gapTwips = 300;
        var caller = new PdfCore.PdfCreateOptions {
            DefaultFont = PdfCore.PdfStandardFont.Helvetica,
            DefaultFontSize = global::OfficeIMO.OfficeLength.Points(12D),
            DefaultTableStyle = new PdfCore.PdfTableStyle {
                BorderColor = global::OfficeIMO.OfficeColor.FromRgb(0, 0, 255),
                BorderWidth = 1D,
                CellPaddingX = 7D,
                CellPaddingY = 3D
            }
        };
        RtfToPdfOptions options = Options();
        options.PdfOptions = caller;
        PdfCore.PdfDocument mapped = ReadTable(gapTwips: gapTwips).ToPdf(options);
        PdfCore.TableBlock table = Assert.Single(mapped.Blocks.OfType<PdfCore.TableBlock>());
        Assert.Equal(global::OfficeIMO.OfficeColor.FromRgb(0, 0, 255), table.Style!.BorderColor);
        Assert.Equal(1D, table.Style.BorderWidth);
        Assert.Equal(7D, table.Style.CellPaddingX);
        Assert.Equal(3D, table.Style.CellPaddingY);
        Assert.Equal(7D, caller.DefaultTableStyle!.CellPaddingX);
        Assert.Equal(3D, caller.DefaultTableStyle.CellPaddingY);
        Assert.Null(caller.DefaultTableStyle.CellPaddings);
        Assert.Null(caller.DefaultTableStyle.CellBorders);

        using PdfPigDocument pdf = PdfPigDocument.Open(mapped.ToBytes());
        Assert.Contains(pdf.GetPage(1).Paths, path => path.IsStroked || path.IsFilled);
        Assert.Single(pdf.GetPage(1).Letters, letter => letter.Value == "A");
    }

    [Theory]
    [InlineData(false, 6D)]
    [InlineData(true, 12D)]
    public void RowAndCellPaddingOverridePositiveGap(bool cellOverride, double expectedLeft) {
        RtfDocument source = ReadTable(gapTwips: 300);
        RtfTableRow row = Assert.Single(source.Blocks.OfType<RtfTable>()).Rows[0];
        row.SetPadding(leftTwips: 120);
        if (cellOverride) row.Cells[0].SetPadding(leftTwips: 240);
        PdfCore.PdfDocument mapped = source.ToPdf(Options());
        PdfCore.TableBlock table = Assert.Single(mapped.Blocks.OfType<PdfCore.TableBlock>());
        Assert.Equal(expectedLeft, table.Style!.CellPaddings![(0, 0)].Left);
        using PdfPigDocument pdf = PdfPigDocument.Open(mapped.ToBytes());
        var marker = Assert.Single(pdf.GetPage(1).Letters, letter => letter.Value == "A");
        Assert.InRange(marker.StartBaseLine.X, 72D + expectedLeft - 0.01D, 72D + expectedLeft + 0.01D);
    }

    private static RtfDocument ReadTable(string cellControls = "", int gapTwips = 0) => RtfDocument.Read(
        @"{\rtf1\ansi\deff0{\fonttbl{\f0\fnil Helvetica;}}{\colortbl;\red255\green0\blue0;}" +
        @"\paperw12240\paperh15840\margl1440\margr1440\margt1440\margb1440" +
        @"\trowd\trgaph" + gapTwips.ToString(System.Globalization.CultureInfo.InvariantCulture) +
        @"\trleft0" + cellControls + @"\cellx2880\pard\intbl\f0\fs24 A\cell\row}").Document;

    private static RtfDocument ReadNativeGapFixture(int gapTwips) => RtfDocument.Load(Path.Combine(
        AppContext.BaseDirectory, "Documents", "RtfTableGapNative",
        "table-cell-minimum-exact-gap" + gapTwips.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".rtf"));

    private static string WithoutWhiteSpace(string value) => new string(value.Where(character => !char.IsWhiteSpace(character)).ToArray());

    private static RtfToPdfOptions Options(bool includePdfOptions = true) => new RtfToPdfOptions {
        ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic(),
        TextFallbacks = PdfCore.PdfTextFallbackFeatures.None,
        PdfOptions = includePdfOptions ? new PdfCore.PdfCreateOptions {
            DefaultFont = PdfCore.PdfStandardFont.Helvetica,
            DefaultFontSize = global::OfficeIMO.OfficeLength.Points(12D)
        } : null
    };
}
