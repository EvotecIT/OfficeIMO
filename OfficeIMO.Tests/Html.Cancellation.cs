using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public async Task LoadFromHtmlAsync_Cancelled_Throws() {
            using var cts = new CancellationTokenSource();
            cts.Cancel();
            await Assert.ThrowsAsync<OperationCanceledException>(() => "<p>a</p>".LoadFromHtmlAsync(cancellationToken: cts.Token));
        }

        [Fact]
        public async Task ToHtmlAsync_Cancelled_Throws() {
            using var doc = WordDocument.Create();
            using var cts = new CancellationTokenSource();
            cts.Cancel();
            await Assert.ThrowsAsync<OperationCanceledException>(() => doc.ToHtmlAsync(cancellationToken: cts.Token));
        }

        [Fact]
        public async Task AddHtmlToHeaderAsync_Cancelled_Throws() {
            using var doc = WordDocument.Create();
            using var cts = new CancellationTokenSource();
            cts.Cancel();
            await Assert.ThrowsAsync<OperationCanceledException>(() => doc.AddHtmlToHeaderAsync("<p>h</p>", cancellationToken: cts.Token));
        }

        [Fact]
        public async Task AddHtmlToFooterAsync_Cancelled_Throws() {
            using var doc = WordDocument.Create();
            using var cts = new CancellationTokenSource();
            cts.Cancel();
            await Assert.ThrowsAsync<OperationCanceledException>(() => doc.AddHtmlToFooterAsync("<p>f</p>", cancellationToken: cts.Token));
        }

        [Fact]
        public async Task AddHtmlToBodyAsync_Cancelled_Throws() {
            using var doc = WordDocument.Create();
            using var cts = new CancellationTokenSource();
            cts.Cancel();
            await Assert.ThrowsAsync<OperationCanceledException>(() => doc.AddHtmlToBodyAsync("<p>b</p>", cancellationToken: cts.Token));
        }
    }
}
