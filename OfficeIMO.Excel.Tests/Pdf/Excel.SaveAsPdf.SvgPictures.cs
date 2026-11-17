using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Fact]
    public void SaveAsPdf_DrawsSvgPictureAsVectorsInsteadOfItsRasterFallback() {
        byte[] svg = Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' width='40' height='20'>" +
            "<rect width='40' height='20' fill='#ff0000'/></svg>");
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Data");
        for (int row = 1; row <= 6; row++) sheet.Cell(row, 1).SetValue(row);
        sheet.AddImage(2, 2, svg, new ExcelImageOptions { ContentType = "image/svg+xml",
            Width = global::OfficeIMO.OfficeLength.Points(80), Height = global::OfficeIMO.OfficeLength.Points(40) });

        byte[] pdf = document.ToPdf().ToBytes();

        Assert.Empty(PdfCore.PdfDocument.Load(pdf).Reader.Images());
        Assert.Single(Regex.Matches(Encoding.GetEncoding("ISO-8859-1").GetString(pdf), @"/Subtype /Form /FormType 1 /BBox \[0 0 40 20\]"));
    }
}
