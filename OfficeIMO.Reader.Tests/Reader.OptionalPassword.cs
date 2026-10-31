using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Reader;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderOptionalPasswordTests {
    [Theory]
    [InlineData("docx", false)]
    [InlineData("docx", true)]
    [InlineData("xlsx", false)]
    [InlineData("xlsx", true)]
    [InlineData("pptx", false)]
    [InlineData("pptx", true)]
    public void OptionalPasswordPreservesPlainOfficeInputs(string extension, bool streamInput) {
        using var encoded = new MemoryStream();
        if (extension == "docx") {
            using var word = WordDocument.Create(); word.AddParagraph("Plain source"); word.Save(encoded);
        } else if (extension == "xlsx") {
            using var workbook = ExcelDocument.Create(); workbook.AddSheet("Data").CellValue(1, 1, "Plain source"); workbook.Save(encoded);
        } else {
            using var presentation = PowerPointDocument.Create(); presentation.AddSlide().AddTextBox("Plain source"); presentation.Save(encoded);
        }
        byte[] bytes = encoded.ToArray();
        var options = new ReaderOptions { OpenPassword = "Synthetic-batch-password" };
        OfficeDocumentReadResult result;
        if (streamInput) {
            encoded.Position = 0;
            result = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadDocument(encoded, "source." + extension, options);
            Assert.True(encoded.CanRead);
            Assert.Equal(bytes, encoded.ToArray());
        } else {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + "." + extension);
            try {
                File.WriteAllBytes(path, bytes);
                result = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadDocument(path, options);
                Assert.Equal(bytes, File.ReadAllBytes(path));
            } finally { File.Delete(path); }
        }
        Assert.Contains("Plain source", result.GetText());
    }
}
