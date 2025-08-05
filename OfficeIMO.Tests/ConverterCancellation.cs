using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Html;
using OfficeIMO.Markdown;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public class ConverterCancellation {
        [Fact]
        public async Task MarkdownToWordConverter_CancelledToken_Throws() {
            using MemoryStream input = new MemoryStream(Encoding.UTF8.GetBytes("# Title\nContent"));
            using MemoryStream output = new MemoryStream();
            IWordConverter converter = new MarkdownToWordConverter();
            using CancellationTokenSource cts = new CancellationTokenSource();
            cts.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(async () =>
                await converter.ConvertAsync(input, output, new MarkdownToWordOptions(), cts.Token));
        }

        [Fact]
        public async Task HtmlToWordConverter_CancelledToken_Throws() {
            using MemoryStream input = new MemoryStream(Encoding.UTF8.GetBytes("<p>test</p>"));
            using MemoryStream output = new MemoryStream();
            IWordConverter converter = new HtmlToWordConverter();
            using CancellationTokenSource cts = new CancellationTokenSource();
            cts.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(async () =>
                await converter.ConvertAsync(input, output, new HtmlToWordOptions(), cts.Token));
        }
    }
}
