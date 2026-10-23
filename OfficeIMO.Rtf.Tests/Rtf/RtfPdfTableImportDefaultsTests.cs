using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfPdfTableImportDefaultsTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void BorderlessZeroGapTableHasNoGridAndStartsAtNativeMargin(bool unrelatedPdfOptions) {
        RtfDocument source = ReadTable();
        PdfCore.PdfDocument mapped = source.ToPdfDocument(Options(unrelatedPdfOptions));
        PdfCore.TableBlock table = Assert.Single(mapped.Blocks.OfType<PdfCore.TableBlock>());
        byte[] bytes = mapped.ToBytes();
        Assert.True(bytes.Take(5).SequenceEqual(System.Text.Encoding.ASCII.GetBytes("%PDF-")));
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.DoesNotContain(pdf.GetPage(1).Paths, path => path.IsStroked || path.IsFilled);
        var marker = Assert.Single(pdf.GetPage(1).Letters, letter => letter.Value == "A");
        Assert.InRange(marker.StartBaseLine.X, 71.99D, 72.01D);
        Assert.Null(table.Style!.BorderColor);
        Assert.Equal(0D, table.Style.BorderWidth);
        Assert.Equal(0D, table.Style.CellPaddingX);
        Assert.Equal(0D, table.Style.CellPaddingY);
    }

    [Fact]
    public void ExplicitCellBorderAndPaddingRemainVisible() {
        RtfDocument source = ReadTable(@"\clbrdrl\brdrs\brdrw8\brdrcf1\clpadl240\clpadfl3");
        PdfCore.PdfDocument mapped = source.ToPdfDocument(Options());
        PdfCore.TableBlock table = Assert.Single(mapped.Blocks.OfType<PdfCore.TableBlock>());
        Assert.Null(table.Style!.BorderColor);
        Assert.Equal(12D, table.Style.CellPaddings![(0, 0)].Left);
        PdfCore.PdfCellBorder border = table.Style.CellBorders![(0, 0)];
        Assert.True(border.Left);
        Assert.Equal(PdfCore.PdfColor.FromRgb(255, 0, 0), border.LeftBorder!.Color);
        Assert.True(border.LeftBorder.Width > 0D);

        using PdfPigDocument pdf = PdfPigDocument.Open(mapped.ToBytes());
        Assert.Contains(pdf.GetPage(1).Paths, path => path.IsStroked || path.IsFilled);
        Assert.Single(pdf.GetPage(1).Letters, letter => letter.Value == "A");
    }

    [Fact]
    public void CallerTableStyleOverridesImportFallbackWithoutMutation() {
        var caller = new PdfCore.PdfOptions {
            DefaultFont = PdfCore.PdfStandardFont.Helvetica,
            DefaultFontSize = 12D,
            DefaultTableStyle = new PdfCore.PdfTableStyle {
                BorderColor = PdfCore.PdfColor.FromRgb(0, 0, 255),
                BorderWidth = 1D,
                CellPaddingX = 7D,
                CellPaddingY = 3D
            }
        };
        RtfToPdfOptions options = Options();
        options.PdfOptions = caller;
        PdfCore.PdfDocument mapped = ReadTable().ToPdfDocument(options);
        PdfCore.TableBlock table = Assert.Single(mapped.Blocks.OfType<PdfCore.TableBlock>());
        Assert.Equal(PdfCore.PdfColor.FromRgb(0, 0, 255), table.Style!.BorderColor);
        Assert.Equal(1D, table.Style.BorderWidth);
        Assert.Equal(7D, table.Style.CellPaddingX);
        Assert.Equal(3D, table.Style.CellPaddingY);
        Assert.Equal(7D, caller.DefaultTableStyle!.CellPaddingX);
        Assert.Equal(3D, caller.DefaultTableStyle.CellPaddingY);
        Assert.Null(caller.DefaultTableStyle.CellPaddings);
        Assert.Null(caller.DefaultTableStyle.CellBorders);

        using PdfPigDocument pdf = PdfPigDocument.Open(mapped.ToBytes());
        Assert.NotEmpty(pdf.GetPage(1).Paths);
        Assert.Single(pdf.GetPage(1).Letters, letter => letter.Value == "A");
    }

    private static RtfDocument ReadTable(string cellControls = "") => RtfDocument.Read(
        @"{\rtf1\ansi\deff0{\fonttbl{\f0\fnil Helvetica;}}{\colortbl;\red255\green0\blue0;}" +
        @"\paperw12240\paperh15840\margl1440\margr1440\margt1440\margb1440" +
        @"\trowd\trgaph0\trleft0" + cellControls + @"\cellx2880\pard\intbl\f0\fs24 A\cell\row}").Document;

    private static RtfToPdfOptions Options(bool includePdfOptions = true) => new RtfToPdfOptions {
        ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic(),
        TextFallbacks = PdfCore.PdfTextFallbackFeatures.None,
        PdfOptions = includePdfOptions ? new PdfCore.PdfOptions {
            DefaultFont = PdfCore.PdfStandardFont.Helvetica,
            DefaultFontSize = 12D
        } : null
    };
}
