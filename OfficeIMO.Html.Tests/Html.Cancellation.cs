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
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => OfficeIMO.Html.HtmlConversionDocument.Parse("<p>a</p>").ToWordAsync(cancellationToken: cts.Token));
        }

        [Fact]
        public async Task Header_AddHtmlAsync_Cancelled_Throws() {
            using var doc = WordDocument.Create();
            using var cts = new CancellationTokenSource();
            cts.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => doc.Sections.Last().GetOrCreateHeader(WordHeaderFooterType.Default).AddHtmlAsync(OfficeIMO.Html.HtmlConversionDocument.Parse("<p>h</p>"), cancellationToken: cts.Token));
        }

        [Fact]
        public async Task Footer_AddHtmlAsync_Cancelled_Throws() {
            using var doc = WordDocument.Create();
            using var cts = new CancellationTokenSource();
            cts.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => doc.Sections.Last().GetOrCreateFooter(WordHeaderFooterType.Default).AddHtmlAsync(OfficeIMO.Html.HtmlConversionDocument.Parse("<p>f</p>"), cancellationToken: cts.Token));
        }

        [Fact]
        public async Task AddHtmlToBodyAsync_Cancelled_Throws() {
            using var doc = WordDocument.Create();
            using var cts = new CancellationTokenSource();
            cts.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => doc.AddHtmlAsync(OfficeIMO.Html.HtmlConversionDocument.Parse("<p>b</p>"), cancellationToken: cts.Token));
        }
    }
}
