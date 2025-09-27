using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public async Task Test_WordSaveLoadAsync() {
            var filePath = Path.Combine(_directoryWithFiles, "AsyncWord.docx");
            if (File.Exists(filePath)) File.Delete(filePath);

            await using (var document = WordDocument.Create(filePath)) {
                document.AddParagraph("Async");
                await document.SaveAsync();
            }

            Assert.True(File.Exists(filePath));

            await using (var document = await WordDocument.LoadAsync(filePath, cancellationToken: CancellationToken.None)) {
                Assert.Single(document.Paragraphs);
            }

            File.Delete(filePath);
        }

        [Fact]
        public async Task Test_WordCreateAsync() {
            var filePath = Path.Combine(_directoryWithFiles, "AsyncCreate.docx");
            if (File.Exists(filePath)) File.Delete(filePath);

            await using (var document = await WordDocument.CreateAsync(filePath, cancellationToken: CancellationToken.None)) {
                document.AddParagraph("Created");
                await document.SaveAsync();
            }

            Assert.True(File.Exists(filePath));

            await using (var document = await WordDocument.LoadAsync(filePath, cancellationToken: CancellationToken.None)) {
                Assert.Single(document.Paragraphs);
            }

            File.Delete(filePath);
        }

        [Fact]
        public async Task Test_WordCreateAsync_CanBeCancelled() {
            var filePath = Path.Combine(_directoryWithFiles, "AsyncCreateCancelled.docx");
            var cts = new CancellationTokenSource();
            cts.Cancel();

            await Assert.ThrowsAsync<TaskCanceledException>(() => WordDocument.CreateAsync(filePath, cancellationToken: cts.Token));
        }

        [Fact]
        public async Task Test_WordLoadAsync_CanBeCancelled() {
            var filePath = Path.Combine(_directoryWithFiles, "AsyncLoadCancelled.docx");
            if (File.Exists(filePath)) File.Delete(filePath);

            await using (var document = WordDocument.Create(filePath)) {
                document.AddParagraph("Cancelled");
                document.Save();
            }

            var cts = new CancellationTokenSource();
            cts.Cancel();

            await Assert.ThrowsAsync<TaskCanceledException>(() => WordDocument.LoadAsync(filePath, cancellationToken: cts.Token));

            File.Delete(filePath);
        }

        [Fact]
        public async Task Test_WordSaveAsync_CanBeCancelled() {
            var sourcePath = Path.Combine(_directoryWithFiles, "AsyncWordCancelSource.docx");
            if (File.Exists(sourcePath)) File.Delete(sourcePath);

            var destinationPath = Path.Combine(_directoryWithFiles, "AsyncWordCancelDestination.docx");
            if (File.Exists(destinationPath)) File.Delete(destinationPath);

            await using (var document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Cancelled");

                using var cts = new CancellationTokenSource();
                cts.Cancel();

                await Assert.ThrowsAsync<OperationCanceledException>(() => document.SaveAsync(destinationPath, openWord: false, cancellationToken: cts.Token));
            }

            Assert.False(File.Exists(destinationPath));

            if (File.Exists(sourcePath)) {
                File.Delete(sourcePath);
            }
        }
    }
}
