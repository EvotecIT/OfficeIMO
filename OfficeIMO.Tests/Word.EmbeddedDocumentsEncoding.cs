using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void Test_EmbeddedHtmlFragmentWithUnicode() {
        string filePath = Path.Combine(_directoryWithFiles, "EmbeddedUnicodeFragment.docx");
        const string phrase = "Zażółć gęślą jaźń";
        string html = $"<html><body><p>{phrase}</p></body></html>";

        using (var document = WordDocument.Create(filePath)) {
            document.AddEmbeddedFragment(html, WordAlternativeFormatImportPartType.Html);
            document.Save(false);
        }

        using (var document = WordDocument.Load(filePath)) {
            Assert.Single(document.EmbeddedDocuments);
            AlternativeFormatImportPart part = document._document.MainDocumentPart.AlternativeFormatImportParts.First();
            using var reader = new StreamReader(part.GetStream(), Encoding.UTF8);
            string content = reader.ReadToEnd();
            Assert.Contains(phrase, content);
        }
    }
}
