using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public async Task ToWordDocumentAsync_Cancelled_Throws() {
            using var cts = new CancellationTokenSource();
            cts.Cancel();
            await Assert.ThrowsAsync<OperationCanceledException>(() => OfficeIMO.Html.HtmlConversionDocument.Parse("<p>a</p>").ToWordDocumentAsync(cancellationToken: cts.Token));
        }

        [Fact]
        public async Task AddHtmlToHeaderAsync_Cancelled_Throws() {
            using var doc = WordDocument.Create();
            using var cts = new CancellationTokenSource();
            cts.Cancel();
            await Assert.ThrowsAsync<OperationCanceledException>(() => doc.AddHtmlToHeaderAsync(OfficeIMO.Html.HtmlConversionDocument.Parse("<p>h</p>"), cancellationToken: cts.Token));
        }

        [Fact]
        public async Task AddHtmlToFooterAsync_Cancelled_Throws() {
            using var doc = WordDocument.Create();
            using var cts = new CancellationTokenSource();
            cts.Cancel();
            await Assert.ThrowsAsync<OperationCanceledException>(() => doc.AddHtmlToFooterAsync(OfficeIMO.Html.HtmlConversionDocument.Parse("<p>f</p>"), cancellationToken: cts.Token));
        }

        [Fact]
        public async Task AddHtmlToBodyAsync_Cancelled_Throws() {
            using var doc = WordDocument.Create();
            using var cts = new CancellationTokenSource();
            cts.Cancel();
            await Assert.ThrowsAsync<OperationCanceledException>(() => doc.AddHtmlToBodyAsync(OfficeIMO.Html.HtmlConversionDocument.Parse("<p>b</p>"), cancellationToken: cts.Token));
        }
    }
}
