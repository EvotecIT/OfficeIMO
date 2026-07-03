using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_LoadMissingFile_ThrowsWithPath() {
            string filePath = Path.Combine(_directoryWithFiles, "missing.docx");
            var ex = Assert.Throws<FileNotFoundException>(() => WordDocument.Load(filePath));
            Assert.Equal($"File '{filePath}' doesn't exist.", ex.Message);
        }

        [Fact]
        public async Task Test_LoadAsyncMissingFile_ThrowsWithPath() {
            string filePath = Path.Combine(_directoryWithFiles, "missingAsync.docx");
            var ex = await Assert.ThrowsAsync<FileNotFoundException>(() => WordDocument.LoadAsync(filePath, cancellationToken: CancellationToken.None));
            Assert.Equal($"File '{filePath}' doesn't exist.", ex.Message);
        }

        [Fact]
        public void Test_AddMacroMissingFile_ThrowsWithPath() {
            using var document = WordDocument.Create(Path.Combine(_directoryWithFiles, "macro.docm"));
            string macroPath = Path.Combine(_directoryWithFiles, "missingMacro.bin");
            var ex = Assert.Throws<FileNotFoundException>(() => WordMacro.AddMacro(document, macroPath));
            Assert.Equal($"File '{macroPath}' doesn't exist.", ex.Message);
        }

        [Fact]
        public void Test_EmbedFontMissingFile_ThrowsWithPath() {
            using var document = WordDocument.Create(Path.Combine(_directoryWithFiles, "font.docx"));
            string fontPath = Path.Combine(_directoryWithFiles, "missing.ttf");
            var ex = Assert.Throws<FileNotFoundException>(() => document.EmbedFont(fontPath));
            Assert.Equal($"Font file '{fontPath}' doesn't exist.", ex.Message);
        }
    }
}
