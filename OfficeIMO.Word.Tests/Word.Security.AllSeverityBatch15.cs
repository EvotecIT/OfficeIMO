using System.Xml.Linq;
using DocumentFormat.OpenXml.CustomXmlDataProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordAllSeverityBatch15SecurityTests {
    [Fact]
    public void NestedTablePreservesNonTextParagraphContentByDefault() {
        using WordDocument document = WordDocument.Create();
        WordTable outer = document.AddTable(1, 1);
        WordTableCell cell = outer.Rows[0].Cells[0];
        Paragraph paragraph = cell.Paragraphs[0]._paragraph;
        paragraph.Append(new BookmarkStart { Name = "marker", Id = "1" });
        paragraph.Append(new BookmarkEnd { Id = "1" });

        cell.AddTable(1, 1);

        Assert.Single(cell._tableCell.Descendants<BookmarkStart>());
        Assert.Single(cell._tableCell.Descendants<BookmarkEnd>());
    }

    [Fact]
    public void CoverPagePropertiesRejectExternalEntities() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-cover-page-" + Guid.NewGuid().ToString("N"));
        string documentPath = Path.Combine(root, "cover.docx");
        string secretPath = Path.Combine(root, "secret.txt");
        Directory.CreateDirectory(root);
        File.WriteAllText(secretPath, "should-not-be-expanded");
        try {
            using (WordDocument document = WordDocument.Create(documentPath)) {
                document.CoverPageProperties.Abstract = "safe";
                document.Save();
            }

            using (WordprocessingDocument package = WordprocessingDocument.Open(documentPath, true)) {
                CustomXmlPart part = package.MainDocumentPart!.CustomXmlParts.Single(customPart =>
                    string.Equals(customPart.CustomXmlPropertiesPart?.DataStoreItem?.ItemId?.Value,
                        WordCoverPageProperties.CoverPagePropsStoreItemId,
                        StringComparison.OrdinalIgnoreCase));
                using Stream output = part.GetStream(FileMode.Create, FileAccess.Write);
                using var writer = new StreamWriter(output);
                writer.Write("<!DOCTYPE CoverPageProperties [<!ENTITY xxe SYSTEM \"");
                writer.Write(new Uri(secretPath).AbsoluteUri);
                writer.Write("\">]><CoverPageProperties xmlns=\"http://schemas.microsoft.com/office/2006/coverPageProps\"><Abstract>&xxe;</Abstract></CoverPageProperties>");
            }

            using WordDocument loaded = WordDocument.Load(documentPath);
            Assert.DoesNotContain("should-not-be-expanded", loaded.CoverPageProperties.Abstract, StringComparison.Ordinal);
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, true);
        }
    }

    [Fact]
    public void WideTableOnlineNormalizationCompletesWithoutRepeatedCellMaterialization() {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(2, 256);
        for (int column = 0; column < 256; column++) {
            table.Rows[0].Cells[column].Width = 100;
        }

        Exception? exception = Record.Exception(table.NormalizeForOnline);

        Assert.Null(exception);
        Assert.Equal(256, table.Rows[0].Cells.Count);
    }
}
