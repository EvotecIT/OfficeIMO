using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Text;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public async Task ToHtmlAsync_EqualsSync() {
            using var doc = WordDocument.Create();
            doc.AddParagraph("Async test");
            string sync = doc.ToHtml();
            string asyncResult = await doc.ToHtmlAsync();
            Assert.Equal(sync, asyncResult);
        }

        [Fact]
        public async Task LoadFromHtmlAsync_EqualsSync() {
            string html = "<p>Hello</p>";
            using var syncDoc = html.LoadFromHtml();
            using var asyncDoc = await html.LoadFromHtmlAsync();
            Assert.Equal(syncDoc.Paragraphs.Count, asyncDoc.Paragraphs.Count);
            Assert.Equal(syncDoc.Paragraphs.First().Text, asyncDoc.Paragraphs.First().Text);
        }

        [Fact]
        public async Task SaveAsHtmlAsync_EqualsSync() {
            using var doc = WordDocument.Create();
            doc.AddParagraph("Save test");
            string dir = Path.Combine(AppContext.BaseDirectory, "HtmlAsync");
            Directory.CreateDirectory(dir);
            string syncPath = Path.Combine(dir, "sync.html");
            string asyncPath = Path.Combine(dir, "async.html");
            if (File.Exists(syncPath)) File.Delete(syncPath);
            if (File.Exists(asyncPath)) File.Delete(asyncPath);

            doc.SaveAsHtml(syncPath);
            await doc.SaveAsHtmlAsync(asyncPath);

            string syncHtml = File.ReadAllText(syncPath);
            string asyncHtml = File.ReadAllText(asyncPath);
            Assert.Equal(syncHtml, asyncHtml);
        }

        [Fact]
        public async Task AddHtmlHeaderFooterAsync_EqualsSync() {
            using var docSync = WordDocument.Create();
            using var docAsync = WordDocument.Create();
            string fragment = "<p>Header</p>";
            docSync.AddHtmlToHeader(fragment);
            await docAsync.AddHtmlToHeaderAsync(fragment);
            Assert.Equal(docSync.Header.Default.Paragraphs[0].Text, docAsync.Header.Default.Paragraphs[0].Text);

            string footerFrag = "<p>Footer</p>";
            docSync.AddHtmlToFooter(footerFrag);
            await docAsync.AddHtmlToFooterAsync(footerFrag);
            Assert.Equal(docSync.Footer.Default.Paragraphs[0].Text, docAsync.Footer.Default.Paragraphs[0].Text);
        }

        [Fact]
        public async Task AsyncMethods_CanBeCancelled() {
            using var doc = WordDocument.Create();
            var cts = new CancellationTokenSource();
            cts.Cancel();
            await Assert.ThrowsAsync<OperationCanceledException>(() => doc.ToHtmlAsync(cancellationToken: cts.Token));
            await Assert.ThrowsAsync<OperationCanceledException>(() => "<p>a</p>".LoadFromHtmlAsync(cancellationToken: cts.Token));
            await Assert.ThrowsAsync<OperationCanceledException>(() => doc.SaveAsHtmlAsync("foo.html", cancellationToken: cts.Token));
            await Assert.ThrowsAsync<OperationCanceledException>(() => doc.AddHtmlToHeaderAsync("<p>h</p>", cancellationToken: cts.Token));
            await Assert.ThrowsAsync<OperationCanceledException>(() => doc.AddHtmlToFooterAsync("<p>f</p>", cancellationToken: cts.Token));
        }

        [Fact]
        public async Task SaveAsHtmlAsync_CancellationDoesNotCreateFile() {
            using var doc = WordDocument.Create();
            doc.AddParagraph("cancel");
            string path = Path.Combine(AppContext.BaseDirectory, "CancelFile.html");
            if (File.Exists(path)) {
                File.Delete(path);
            }
            using var cts = new CancellationTokenSource();
            cts.Cancel();
            await Assert.ThrowsAsync<OperationCanceledException>(() => doc.SaveAsHtmlAsync(path, cancellationToken: cts.Token));
            Assert.False(File.Exists(path));
        }

        [Fact]
        public async Task LoadFromHtmlAsync_StreamCancellationLeavesPosition() {
            using var stream = new MemoryStream(Encoding.UTF8.GetBytes("<p>x</p>"));
            using var cts = new CancellationTokenSource();
            cts.Cancel();
            long pos = stream.Position;
            await Assert.ThrowsAsync<OperationCanceledException>(() => stream.LoadFromHtmlAsync(cancellationToken: cts.Token));
            Assert.Equal(pos, stream.Position);
        }
    }
}
