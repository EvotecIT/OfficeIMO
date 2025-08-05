using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void SaveAsPdf_Uses_DefaultPageSettings() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfDefaultSettings.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfDefaultSettings.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Hello World");
                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    DefaultOrientation = PageOrientationValues.Landscape,
                    DefaultPageSize = WordPageSize.A4
                });
            }

            Assert.True(File.Exists(pdfPath));

            string pdfContent = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
            Match mediaBox = Regex.Match(pdfContent, @"/MediaBox\s*\[\s*0\s+0\s+(?<w>[0-9\.]+)\s+(?<h>[0-9\.]+)\s*\]");
            Assert.True(mediaBox.Success, "MediaBox not found");
            double width = double.Parse(mediaBox.Groups["w"].Value, CultureInfo.InvariantCulture);
            double height = double.Parse(mediaBox.Groups["h"].Value, CultureInfo.InvariantCulture);
            Assert.True(width > height);
        }
    }
}
