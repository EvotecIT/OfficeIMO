using AngleSharp;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using System.Threading;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private void ProcessLinkedStylesheetElement(IElement element) {
            var rel = element.GetAttribute("rel");
            if (!string.Equals(rel, "stylesheet", StringComparison.OrdinalIgnoreCase)) {
                return;
            }

            var hrefAttr = element.GetAttribute("href");
            var href = (element as IHtmlLinkElement)?.Href ?? hrefAttr;
            if (string.IsNullOrEmpty(href)) {
                return;
            }

            if (!string.IsNullOrEmpty(hrefAttr) && File.Exists(hrefAttr)) {
                ParseCss(File.ReadAllText(hrefAttr), hrefAttr);
                return;
            }

            var url = new Url(href);
            if (!url.IsAbsolute && element.BaseUrl != null) {
                url = new Url(new Url(element.BaseUrl), href);
            }

            if (url.Scheme == "http" || url.Scheme == "https") {
                if (_context != null) {
                    LoadAndParseCssAsync(_context, url, CancellationToken.None).GetAwaiter().GetResult();
                }
            } else if (url.Scheme == "file") {
                TryLoadCssFromFileUrl(url);
            }
        }
    }
}
