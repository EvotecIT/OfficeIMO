using OfficeIMO.Html.Dom;
using System.Globalization;

namespace OfficeIMO.Chm;

public sealed partial class ChmDocument {
    // Keep each topic's inheritance on its own section. The first selected topic
    // supplies the book's primary language; the original containers can override it.
    private void PreserveTopicLanguage(HtmlDocument output, HtmlElement section, HtmlDocument topic, bool first) {
        string? language = ContainerLanguage(topic.Body) ?? ContainerLanguage(topic.DocumentElement);
        if (language == null) {
            try { language = CultureInfo.GetCultureInfo(checked((int)LocaleId)).Name; }
            catch (ArgumentException) { }
            catch (OverflowException) { }
        }
        if (!string.IsNullOrWhiteSpace(language)) {
            section.SetAttribute("lang", language!);
            if (first) output.DocumentElement!.SetAttribute("lang", language!);
        }
        string? direction = topic.Body?.GetAttribute("dir") ?? topic.DocumentElement?.GetAttribute("dir");
        if (direction != null) section.SetAttribute("dir", direction);
    }

    private static string? ContainerLanguage(HtmlElement? element) {
        string? value = element?.GetAttribute("lang") ?? element?.GetAttribute("xml:lang");
        return string.IsNullOrWhiteSpace(value) ? null : value;
    }

    // HTML and body containers are replaced by inert divs, retaining their global
    // attributes and anchors without introducing nested HTML/body elements.
    private static HtmlElement PreserveTopicContainer(HtmlDocument output, HtmlElement parent, HtmlElement? source,
        string topicPath, List<OfficeConversionFidelityDiagnostic> diagnostics, ref int nodes, ChmConversionOptions options) {
        if (source == null) return parent;
        HtmlElement? container = null;
        foreach (HtmlAttribute attribute in source.Attributes) {
            string name = attribute.LocalName;
            bool retained = attribute.NamespaceUri.Length == 0 &&
                (name == "id" || name == "class" || name == "style" || name == "lang" || name == "dir" ||
                 name == "title" || name == "role" || name == "hidden" || name == "tabindex" ||
                 name.StartsWith("aria-", StringComparison.Ordinal) || name.StartsWith("data-", StringComparison.Ordinal));
            bool xmlLanguage = attribute.Name == "xml:lang";
            if (retained || xmlLanguage) {
                if (container == null) {
                    if (nodes >= options.MaxHtmlNodes) throw ChmBinary.Error("CONVERSION_LIMIT", "The combined book exceeds MaxHtmlNodes.");
                    nodes++; container = output.CreateElement("div"); parent.AppendChild(container);
                }
                if (xmlLanguage) {
                    if (!container.HasAttribute("lang")) container.SetAttribute("lang", attribute.Value);
                } else container.SetAttribute(attribute.Name, attribute.Value);
            } else if (name != "xmlns") {
                diagnostics.Add(new OfficeConversionFidelityDiagnostic("CHM_CONTAINER_ATTRIBUTE_OMITTED",
                    "A topic container attribute cannot be applied to its reflowable wrapper: " + attribute.Name + ".",
                    OfficeConversionLossKind.Omission, "OfficeIMO.Chm", topicPath));
            }
        }
        return container ?? parent;
    }
}
