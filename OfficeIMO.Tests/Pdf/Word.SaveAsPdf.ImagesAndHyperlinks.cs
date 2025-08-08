using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System;
using System.IO;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void Test_WordDocument_SaveAsPdf_ImagesAndHyperlinks() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfImagesLinks.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfImagesLinks.pdf");
        string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph().AddImage(imagePath, 50, 50);
            document.AddHyperLink("OfficeIMO", new Uri("https://evotec.xyz"), addStyle: true);
            document.Save();
            document.SaveAsPdf(pdfPath);
        }

        Assert.True(File.Exists(pdfPath));
        string pdfContent = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
        Assert.Contains("/Subtype /Image", pdfContent);
        Assert.Contains("/URI (https://evotec.xyz", pdfContent);
    }
}
