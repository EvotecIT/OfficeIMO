using AngleSharp.Dom;
using System.Globalization;
using System.Threading;

namespace OfficeIMO.Word.Html {
    internal partial class WordToHtmlConverter {
        private static void AppendHeadMetadata(
            WordDocument document,
            IDocument htmlDoc,
            IElement head,
            WordToHtmlOptions options,
            CancellationToken cancellationToken) {
            ApplyDocumentShellMetadata(document, htmlDoc);

            var charset = htmlDoc.CreateElement("meta");
            charset.SetAttribute("charset", "UTF-8");
            head.AppendChild(charset);

            var props = document.BuiltinDocumentProperties;
            var title = htmlDoc.CreateElement("title");
            var titleText = string.IsNullOrEmpty(props?.Title) ? "Document" : props!.Title!;
            title.TextContent = titleText;
            head.AppendChild(title);

            if (props != null) {
                AddMeta(htmlDoc, head, "author", props.Creator);
                AddMeta(htmlDoc, head, "description", props.Description);
                AddMeta(htmlDoc, head, "keywords", props.Keywords);
                AddMeta(htmlDoc, head, "subject", props.Subject);
            }

            if (options.IncludeCustomProperties) {
                foreach (var property in document.CustomDocumentProperties) {
                    cancellationToken.ThrowIfCancellationRequested();
                    AddCustomPropertyMeta(htmlDoc, head, property.Key, property.Value);
                }
            }

            foreach (var (name, content) in options.AdditionalMetaTags) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!string.IsNullOrEmpty(name)) {
                    var meta = htmlDoc.CreateElement("meta");
                    meta.SetAttribute("name", name);
                    if (!string.IsNullOrEmpty(content)) {
                        meta.SetAttribute("content", content);
                    }
                    head.AppendChild(meta);
                }
            }

            foreach (var (rel, href) in options.AdditionalLinkTags) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!string.IsNullOrEmpty(rel) && !string.IsNullOrEmpty(href)) {
                    var link = htmlDoc.CreateElement("link");
                    link.SetAttribute("rel", rel);
                    link.SetAttribute("href", href);
                    head.AppendChild(link);
                }
            }

            if (options.IncludeDefaultCss) {
                var style = htmlDoc.CreateElement("style");
                style.TextContent = WordHtmlResources.DefaultCss;
                head.AppendChild(style);
            }
        }

        private static void ApplyDocumentShellMetadata(WordDocument document, IDocument htmlDoc) {
            var language = document.Settings.Language;
            if (!string.IsNullOrWhiteSpace(language)) {
                htmlDoc.DocumentElement.SetAttribute("lang", language!.Trim());
            }
        }

        private static void AddMeta(IDocument htmlDoc, IElement head, string name, string? value) {
            if (!string.IsNullOrEmpty(value)) {
                var meta = htmlDoc.CreateElement("meta");
                meta.SetAttribute("name", name);
                meta.SetAttribute("content", value);
                head.AppendChild(meta);
            }
        }

        private static void AddCustomPropertyMeta(IDocument htmlDoc, IElement head, string name, WordCustomProperty property) {
            if (string.IsNullOrWhiteSpace(name)) {
                return;
            }

            var value = FormatCustomPropertyValue(property);
            if (value == null) {
                return;
            }

            var meta = htmlDoc.CreateElement("meta");
            meta.SetAttribute("name", "word:custom:" + name);
            meta.SetAttribute("content", value);
            meta.SetAttribute("data-word-custom-property", name);
            meta.SetAttribute("data-property-type", property.PropertyType.ToString());
            head.AppendChild(meta);
        }

        private static string? FormatCustomPropertyValue(WordCustomProperty property) {
            if (property.Value == null) {
                return null;
            }

            return property.PropertyType switch {
                PropertyTypes.YesNo => property.Value is bool value ? value.ToString().ToLowerInvariant() : property.Value.ToString(),
                PropertyTypes.DateTime => property.Value is DateTime value ? value.ToString("O", CultureInfo.InvariantCulture) : property.Value.ToString(),
                PropertyTypes.NumberInteger => System.Convert.ToString(property.Value, CultureInfo.InvariantCulture),
                PropertyTypes.NumberDouble => System.Convert.ToString(property.Value, CultureInfo.InvariantCulture),
                _ => property.Value.ToString()
            };
        }
    }
}
