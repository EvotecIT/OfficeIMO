using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Word;
using System.IO;
using System.IO.Compression;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void Test_WordDocument_SaveAsPdf_TableStyles() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfStyledTable.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfStyledTable.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(1, 1);
            WordTableCell cell = table.Rows[0].Cells[0];
            cell.Paragraphs[0].Text = "Styled";
            cell.ShadingFillColorHex = "FF0000";
            cell.Borders.TopStyle = BorderValues.Single;
            cell.Borders.BottomStyle = BorderValues.Single;
            cell.Borders.LeftStyle = BorderValues.Single;
            cell.Borders.RightStyle = BorderValues.Single;
            cell.Borders.TopColorHex = "0000FF";
            cell.Borders.BottomColorHex = "0000FF";
            cell.Borders.LeftColorHex = "0000FF";
            cell.Borders.RightColorHex = "0000FF";
            cell.Borders.TopSize = 8;
            cell.Borders.BottomSize = 8;
            cell.Borders.LeftSize = 8;
            cell.Borders.RightSize = 8;
            document.Save();
            document.SaveAsPdf(pdfPath);
        }

        byte[] bytes = File.ReadAllBytes(pdfPath);
        byte[] startPattern = Encoding.ASCII.GetBytes("stream\n");
        byte[] endPattern = Encoding.ASCII.GetBytes("\nendstream");
        int start = IndexOf(bytes, startPattern, 0);
        Assert.True(start >= 0, "stream marker not found");
        start += startPattern.Length;
        int end = IndexOf(bytes, endPattern, start);
        Assert.True(end >= 0, "endstream marker not found");
        int length = end - start;
        int deflateLength = length - 6;
        byte[] deflateData = new byte[deflateLength];
        Array.Copy(bytes, start + 2, deflateData, 0, deflateLength);
        using MemoryStream ms = new MemoryStream(deflateData);
        using DeflateStream ds = new DeflateStream(ms, CompressionMode.Decompress);
        using StreamReader reader = new StreamReader(ds, Encoding.GetEncoding("ISO-8859-1"));
        string content = reader.ReadToEnd();
        Assert.Contains("1 0 0 rg", content);
        Assert.Contains("0 0 1 RG", content);
    }
}
