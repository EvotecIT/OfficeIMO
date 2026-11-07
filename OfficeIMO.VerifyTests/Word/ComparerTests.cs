using OfficeIMO.Word;
using System.Threading.Tasks;
using VerifyTests;
using VerifyXunit;
using Xunit;

namespace OfficeIMO.VerifyTests.Word;

public class ComparerTests {

    [Fact]
    public async Task CompareListDocument() {
        using var sourcePath = new TempFile("docx");
        using var targetPath = new TempFile("docx");

        using (var document = WordDocument.Create(sourcePath)) {
            var list = document.AddList(WordListStyle.Numbered);
            list.AddItem("Item 1");
            document.Save();
        }

        using (var document = WordDocument.Create(targetPath)) {
            var list = document.AddList(WordListStyle.Numbered);
            list.AddItem("Item 1 updated");
            document.Save();
        }

        using (var result = WordDocumentComparer.Compare(sourcePath, targetPath)) {
            await Verifier.Verify(result);
        }
    }

    [Fact]
    public async Task CompareInsertedTableCellDocument() {
        using var sourcePath = new TempFile("docx");
        using var targetPath = new TempFile("docx");

        using (var document = WordDocument.Create(sourcePath)) {
            var table = document.AddTable(1, 1, WordTableStyle.OfficeIMOModern);
            table.Rows[0].Cells[0].Paragraphs[0].SetText("Left");
            document.Save();
        }

        using (var document = WordDocument.Create(targetPath)) {
            var table = document.AddTable(1, 2, WordTableStyle.OfficeIMOModern);
            table.Rows[0].Cells[0].Paragraphs[0].SetText("Left");
            table.Rows[0].Cells[1].Paragraphs[0].SetText("Right");
            document.Save();
        }

        using (var result = WordDocumentComparer.Compare(sourcePath, targetPath)) {
            await Verifier.Verify(result);
        }
    }
}
