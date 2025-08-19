using System.IO;
using System.Linq;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void Test_AddEmbeddedDocument_MissingFile_CleansTemporaryPart() {
        string filePath = Path.Combine(_directoryWithFiles, "MissingFileEmbed.docx");
        using var document = WordDocument.Create(filePath);
        string missingPath = Path.Combine(_directoryWithFiles, "does-not-exist.rtf");

        Assert.Throws<FileNotFoundException>(() => document.AddEmbeddedDocument(missingPath));
        Assert.NotNull(document._document.MainDocumentPart);
        Assert.Empty(document._document.MainDocumentPart!.AlternativeFormatImportParts);
    }
}
