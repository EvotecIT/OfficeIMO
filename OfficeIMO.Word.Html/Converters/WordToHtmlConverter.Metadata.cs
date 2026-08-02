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

            var charset = CreateOutputElement(htmlDoc, "meta");
            SetMetadataAttribute(htmlDoc, charset, "charset", "UTF-8", "DocumentMetadata:charset");
            head.AppendChild(charset);

            var props = document.BuiltinDocumentProperties;
            var title = CreateOutputElement(htmlDoc, "title");
            var titleText = string.IsNullOrEmpty(props?.Title) ? "Document" : props!.Title!;
            SetMetadataText(htmlDoc, title, titleText, "DocumentMetadata:title");
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
                    var meta = CreateOutputElement(htmlDoc, "meta");
                    SetMetadataAttribute(htmlDoc, meta, "name", name, "AdditionalMeta:name");
                    if (!string.IsNullOrEmpty(content)) {
                        SetMetadataAttribute(htmlDoc, meta, "content", content, "AdditionalMeta:content");
                    }
                    head.AppendChild(meta);
                }
            }

            foreach (var (rel, href) in options.AdditionalLinkTags) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!string.IsNullOrEmpty(rel) && !string.IsNullOrEmpty(href)) {
                    var link = CreateOutputElement(htmlDoc, "link");
                    SetMetadataAttribute(htmlDoc, link, "rel", rel, "AdditionalLink:rel");
                    SetMetadataAttribute(htmlDoc, link, "href", href, "AdditionalLink:href");
                    head.AppendChild(link);
                }
            }

            if (options.IncludeDefaultCss) {
                var style = CreateOutputElement(htmlDoc, "style");
                SetMetadataText(htmlDoc, style, WordHtmlResources.DefaultCss, "DocumentMetadata:default-css");
                head.AppendChild(style);
            }
        }

        private static void ApplyDocumentShellMetadata(WordDocument document, IDocument htmlDoc) {
            var language = document.Settings.Language;
            if (!string.IsNullOrWhiteSpace(language)) {
                SetMetadataAttribute(
                    htmlDoc,
                    htmlDoc.DocumentElement,
                    "lang",
                    language!.Trim(),
                    "DocumentMetadata:language");
            }
        }

        private static void AddMeta(IDocument htmlDoc, IElement head, string name, string? value) {
            if (!string.IsNullOrEmpty(value)) {
                var meta = CreateOutputElement(htmlDoc, "meta");
                SetMetadataAttribute(htmlDoc, meta, "name", name, "DocumentMetadata:" + name + ":name");
                SetMetadataAttribute(htmlDoc, meta, "content", value!, "DocumentMetadata:" + name);
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

            var meta = CreateOutputElement(htmlDoc, "meta");
            SetMetadataAttribute(htmlDoc, meta, "name", "word:custom:" + name, "CustomDocumentMetadata:name");
            SetMetadataAttribute(htmlDoc, meta, "content", value, "CustomDocumentMetadata:content");
            SetMetadataAttribute(htmlDoc, meta, "data-word-custom-property", name, "CustomDocumentMetadata:property-name");
            SetMetadataAttribute(htmlDoc, meta, "data-property-type", property.PropertyType.ToString(), "CustomDocumentMetadata:property-type");
            head.AppendChild(meta);
        }

        private static void SetMetadataAttribute(
            IDocument htmlDoc,
            IElement element,
            string name,
            string value,
            string source) {
            ReserveOutputCharacters(
                htmlDoc,
                GetHtmlEncodedLength(value, attributeValue: true),
                "Generated HTML metadata exceeds the configured output-character limit before DOM construction.",
                source);
            element.SetAttribute(name, value);
        }

        private static void SetMetadataText(
            IDocument htmlDoc,
            IElement element,
            string value,
            string source) {
            ReserveOutputCharacters(
                htmlDoc,
                GetHtmlEncodedLength(value, attributeValue: false),
                "Generated HTML metadata exceeds the configured output-character limit before DOM construction.",
                source);
            element.TextContent = value;
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
