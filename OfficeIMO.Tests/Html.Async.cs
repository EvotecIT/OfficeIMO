using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
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
            Assert.Equal(docSync.Header!.Default.Paragraphs[0].Text, docAsync.Header!.Default.Paragraphs[0].Text);

            string footerFrag = "<p>Footer</p>";
            docSync.AddHtmlToFooter(footerFrag);
            await docAsync.AddHtmlToFooterAsync(footerFrag);
            Assert.Equal(docSync.Footer!.Default.Paragraphs[0].Text, docAsync.Footer!.Default.Paragraphs[0].Text);
        }

        [Fact]
        public async Task AddHtmlToFooterAsync_CreatesFirstFooter() {
            await AssertFooterCreatedAsync(HeaderFooterValues.First, "First footer fragment", doc => doc.DifferentFirstPage = true);
        }

        [Fact]
        public async Task AddHtmlToFooterAsync_CreatesEvenFooter() {
            await AssertFooterCreatedAsync(HeaderFooterValues.Even, "Even footer fragment", doc => doc.DifferentOddAndEvenPages = true);
        }

        private static async Task AssertFooterCreatedAsync(HeaderFooterValues footerType, string expectedText, Action<WordDocument> configure) {
            using var doc = WordDocument.Create();
            configure(doc);

            string html = $"<p>{expectedText}</p>";
            await doc.AddHtmlToFooterAsync(html, footerType);

            var section = doc.Sections.Last();
            var footers = section.Footer;
            Assert.NotNull(footers);
            Assert.NotNull(ResolveFooter(footers!, footerType));

            string innerText = GetFooterInnerText(doc, footerType);
            Assert.Contains(expectedText, innerText);
        }

        private static WordFooter? ResolveFooter(WordFooters footers, HeaderFooterValues type) {
            if (type == HeaderFooterValues.First) return footers.First;
            if (type == HeaderFooterValues.Even) return footers.Even;
            return footers.Default;
        }

        private static string GetFooterInnerText(WordDocument doc, HeaderFooterValues footerType) {
            using var ms = new MemoryStream();
            doc.Save(ms);
            ms.Position = 0;

            using var package = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(ms, false);
            var body = package.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("The saved document body is missing.");
            var sectionProperties = body.Descendants<SectionProperties>().LastOrDefault() ?? throw new InvalidOperationException("The document does not define section properties.");

            FooterReference? footerReference = sectionProperties.Elements<FooterReference>().FirstOrDefault(reference =>
                footerType == HeaderFooterValues.Default
                    ? reference.Type == null || reference.Type.Value == HeaderFooterValues.Default
                    : reference.Type?.Value == footerType);

            if (footerReference?.Id == null) {
                throw new InvalidOperationException($"The {footerType} footer reference could not be located in the saved document.");
            }

            var footerPart = package.MainDocumentPart!.GetPartById(footerReference.Id.Value) as DocumentFormat.OpenXml.Packaging.FooterPart
                ?? throw new InvalidOperationException("Unable to resolve the footer part from the document package.");

            return footerPart.Footer?.InnerText ?? string.Empty;
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
    }
}
