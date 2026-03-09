using OfficeIMO.Word;
using System;
using System.IO;
using System.Threading.Tasks;
using VerifyXunit;
using Xunit;

namespace OfficeIMO.VerifyTests.Word;

public class ComparerTests : VerifyTestBase {
    private static async Task DoTest(string sourcePath, string targetPath) {
        using var result = WordDocumentComparer.Compare(sourcePath, targetPath);
        var verifyResult = await ToVerifyResult(result._wordprocessingDocument);
        await Verifier.Verify(verifyResult, GetSettings());
    }

    private static string CreateTempDocxPath() {
        return Path.Combine(Path.GetTempPath(), "OfficeIMO.VerifyTests." + Guid.NewGuid().ToString("N") + ".docx");
    }

    [Fact]
    public async Task CompareListDocument() {
        var sourcePath = CreateTempDocxPath();
        var targetPath = CreateTempDocxPath();

        try {
            using (var document = WordDocument.Create(sourcePath)) {
                var list = document.AddList(WordListStyle.Numbered);
                list.AddItem("Item 1");
                document.Save(false);
            }

            using (var document = WordDocument.Create(targetPath)) {
                var list = document.AddList(WordListStyle.Numbered);
                list.AddItem("Item 1 updated");
                document.Save(false);
            }

            await DoTest(sourcePath, targetPath);
        } finally {
            if (File.Exists(sourcePath)) {
                File.Delete(sourcePath);
            }

            if (File.Exists(targetPath)) {
                File.Delete(targetPath);
            }
        }
    }

    [Fact]
    public async Task CompareInsertedTableCellDocument() {
        var sourcePath = CreateTempDocxPath();
        var targetPath = CreateTempDocxPath();

        try {
            using (var document = WordDocument.Create(sourcePath)) {
                var table = document.AddTable(1, 1);
                table.Rows[0].Cells[0].Paragraphs[0].SetText("Left");
                document.Save(false);
            }

            using (var document = WordDocument.Create(targetPath)) {
                var table = document.AddTable(1, 2);
                table.Rows[0].Cells[0].Paragraphs[0].SetText("Left");
                table.Rows[0].Cells[1].Paragraphs[0].SetText("Right");
                document.Save(false);
            }

            await DoTest(sourcePath, targetPath);
        } finally {
            if (File.Exists(sourcePath)) {
                File.Delete(sourcePath);
            }

            if (File.Exists(targetPath)) {
                File.Delete(targetPath);
            }
        }
    }
}
