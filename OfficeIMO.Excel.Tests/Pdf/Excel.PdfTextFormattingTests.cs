using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelPdfTextFormattingTests {
    [Fact]
    public void NativeCellFormattingProjectsToTypedPdfRuns() {
        using ExcelDocument document = ExcelDocument.Create();
        document.AddWorksheet("Text").CellAt(1, 1)
            .SetValue("Styled")
            .SetBold()
            .SetItalic()
            .SetUnderline(ExcelUnderlineStyle.Double)
            .SetStrikethrough()
            .SetSuperscript()
            .SetFontName("Aptos")
            .SetFontSize(14)
            .SetFontColor("336699");

        PdfCore.PdfDocument pdf = document.ToPdfDocument(new ExcelPdfSaveOptions {
            IncludeSheetHeadings = false,
            WorksheetLayout = ExcelPdfWorksheetLayoutMode.FlowTable
        });
        PdfCore.PageBlock page = Assert.IsType<PdfCore.PageBlock>(Assert.Single(pdf.Blocks));
        PdfCore.TableBlock table = Assert.Single(page.Blocks.OfType<PdfCore.TableBlock>());
        PdfCore.PdfTextRun run = Assert.Single(table.Cells[0][0].Runs);

        Assert.Equal("Styled", run.Text);
        Assert.True(run.Bold);
        Assert.True(run.Italic);
        Assert.Equal(OfficeTextDecorationStyle.Double, run.UnderlineStyle);
        Assert.Equal(OfficeTextDecorationStyle.Single, run.StrikeStyle);
        Assert.Equal(PdfCore.PdfTextBaseline.Superscript, run.Baseline);
        Assert.Equal(PdfCore.PdfColor.FromRgb(51, 102, 153), run.Color);
        Assert.Equal("Aptos", run.FontFamily);
        Assert.Equal(14D, run.FontSize);
    }
}
