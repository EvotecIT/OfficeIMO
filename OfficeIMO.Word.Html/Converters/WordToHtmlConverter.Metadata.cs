using AngleSharp.Dom;
using OfficeIMO.Html;
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
            if (!options.EmitDocumentShell) {
                if (options.IncludeDefaultCss) {
                    var fragmentStyle = CreateOutputElement(htmlDoc, "style");
                    SetMetadataText(htmlDoc, fragmentStyle, WordHtmlResources.GetDefaultCss(options.Theme, options.UseSharedDocumentShell), "DocumentMetadata:default-css");
                    head.AppendChild(fragmentStyle);
                }
                return;
            }

            ApplyDocumentShellMetadata(document, htmlDoc, options);

            var charset = CreateOutputElement(htmlDoc, "meta");
            SetOutputAttribute(htmlDoc, charset, "charset", "UTF-8", "DocumentMetadata:charset");
            head.AppendChild(charset);

            if (options.IncludeDefaultCss && options.UseSharedDocumentShell) {
                var viewport = CreateOutputElement(htmlDoc, "meta");
                SetOutputAttribute(htmlDoc, viewport, "name", "viewport", "DocumentMetadata:viewport-name");
                SetOutputAttribute(htmlDoc, viewport, "content", "width=device-width, initial-scale=1", "DocumentMetadata:viewport-content");
                head.AppendChild(viewport);
            }

            var props = document.BuiltinDocumentProperties;
            var title = CreateOutputElement(htmlDoc, "title");
            var titleText = !string.IsNullOrWhiteSpace(options.Title)
                ? options.Title!
                : string.IsNullOrEmpty(props?.Title) ? "Document" : props!.Title!;
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
                    SetOutputAttribute(htmlDoc, meta, "name", name, "AdditionalMeta:name");
                    if (!string.IsNullOrEmpty(content)) {
                        SetOutputAttribute(htmlDoc, meta, "content", content, "AdditionalMeta:content");
                    }
                    head.AppendChild(meta);
                }
            }

            foreach (var (rel, href) in options.AdditionalLinkTags) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!string.IsNullOrEmpty(rel) && !string.IsNullOrEmpty(href)) {
                    var link = CreateOutputElement(htmlDoc, "link");
                    SetOutputAttribute(htmlDoc, link, "rel", rel, "AdditionalLink:rel");
                    SetOutputAttribute(htmlDoc, link, "href", href, "AdditionalLink:href");
                    head.AppendChild(link);
                }
            }

            if (options.IncludeDefaultCss) {
                var style = CreateOutputElement(htmlDoc, "style");
                SetMetadataText(htmlDoc, style, WordHtmlResources.GetDefaultCss(options.Theme, options.UseSharedDocumentShell), "DocumentMetadata:default-css");
                head.AppendChild(style);
            }
        }

        private static void ApplyDocumentShellMetadata(WordDocument document, IDocument htmlDoc, WordToHtmlOptions options) {
            var language = options.Language ?? document.Settings.Language;
            if (!string.IsNullOrWhiteSpace(language)) {
                SetOutputAttribute(
                    htmlDoc,
                    htmlDoc.DocumentElement,
                    "lang",
                    language!.Trim(),
                    "DocumentMetadata:language");
            }

            if (options.UseSharedDocumentShell) {
                SetOutputAttribute(
                    htmlDoc,
                    htmlDoc.DocumentElement,
                    "data-officeimo-profile",
                    options.Profile.ToString(),
                    "DocumentMetadata:profile");
            }
            if (htmlDoc.Body != null &&
                (options.UseSharedDocumentShell ||
                 !string.Equals(options.DocumentOutput.BodyClass, "officeimo-html officeimo-word-html", StringComparison.Ordinal))) {
                string bodyClass = OfficeHtmlDocumentShell.MergeBodyClasses(
                    "officeimo-html officeimo-word-html",
                    options.DocumentOutput.BodyClass);
                if (bodyClass.Length > 0) {
                    SetOutputAttribute(
                        htmlDoc,
                        htmlDoc.Body,
                        "class",
                        bodyClass,
                        "DocumentMetadata:body-class");
                }
            }
        }

        private static void AddMeta(IDocument htmlDoc, IElement head, string name, string? value) {
            if (!string.IsNullOrEmpty(value)) {
                var meta = CreateOutputElement(htmlDoc, "meta");
                SetOutputAttribute(htmlDoc, meta, "name", name, "DocumentMetadata:" + name + ":name");
                SetOutputAttribute(htmlDoc, meta, "content", value!, "DocumentMetadata:" + name);
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
            SetOutputAttribute(htmlDoc, meta, "name", "word:custom:" + name, "CustomDocumentMetadata:name");
            SetOutputAttribute(htmlDoc, meta, "content", value, "CustomDocumentMetadata:content");
            SetOutputAttribute(htmlDoc, meta, "data-word-custom-property", name, "CustomDocumentMetadata:property-name");
            SetOutputAttribute(htmlDoc, meta, "data-property-type", property.PropertyType.ToString(), "CustomDocumentMetadata:property-type");
            head.AppendChild(meta);
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
                WordCustomPropertyType.YesNo => property.Value is bool value ? value.ToString().ToLowerInvariant() : property.Value.ToString(),
                WordCustomPropertyType.DateTime => property.Value is DateTime value ? value.ToString("O", CultureInfo.InvariantCulture) : property.Value.ToString(),
                WordCustomPropertyType.NumberInteger => System.Convert.ToString(property.Value, CultureInfo.InvariantCulture),
                WordCustomPropertyType.NumberDouble => System.Convert.ToString(property.Value, CultureInfo.InvariantCulture),
                _ => property.Value.ToString()
            };
        }
    }
}
