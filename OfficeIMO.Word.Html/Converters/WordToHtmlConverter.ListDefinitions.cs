using AngleSharp.Dom;
using OfficeIMO.Word;
using System.Globalization;
using System.Text;
using System.Threading;

namespace OfficeIMO.Word.Html {
    internal partial class WordToHtmlConverter {
        private readonly struct HtmlListDefinition {
            internal HtmlListDefinition(string className, string listStyleType, int level, int? leftIndentTwips, int? hangingIndentTwips) {
                ClassName = className;
                ListStyleType = listStyleType;
                Level = level;
                LeftIndentTwips = leftIndentTwips;
                HangingIndentTwips = hangingIndentTwips;
            }

            internal string ClassName { get; }
            internal string ListStyleType { get; }
            internal int Level { get; }
            internal int? LeftIndentTwips { get; }
            internal int? HangingIndentTwips { get; }
        }

        private static void ApplyListDefinition(
            IElement listElement,
            DocumentTraversal.ListInfo listInfo,
            string? listStyleType,
            HashSet<HtmlListDefinition> listDefinitions) {
            var effectiveStyle = string.IsNullOrWhiteSpace(listStyleType)
                ? listInfo.Ordered ? "decimal" : "disc"
                : listStyleType!;
            var className = BuildListDefinitionClass(listInfo, effectiveStyle);

            AppendClass(listElement, className);
            listElement.SetAttribute("data-word-list-level", listInfo.Level.ToString(CultureInfo.InvariantCulture));
            if (listInfo.LeftIndentTwips.HasValue) {
                listElement.SetAttribute("data-left-indent-twips", listInfo.LeftIndentTwips.Value.ToString(CultureInfo.InvariantCulture));
            }
            if (listInfo.HangingIndentTwips.HasValue) {
                listElement.SetAttribute("data-hanging-indent-twips", listInfo.HangingIndentTwips.Value.ToString(CultureInfo.InvariantCulture));
            }

            listDefinitions.Add(new HtmlListDefinition(
                className,
                effectiveStyle,
                listInfo.Level,
                listInfo.LeftIndentTwips,
                listInfo.HangingIndentTwips));
        }

        private static void AppendListDefinitions(
            IDocument htmlDoc,
            IElement head,
            HashSet<HtmlListDefinition> listDefinitions,
            CancellationToken cancellationToken) {
            if (listDefinitions.Count == 0) {
                return;
            }

            var styleElement = htmlDoc.CreateElement("style");
            styleElement.SetAttribute("data-word-list-definitions", "true");
            var sb = new StringBuilder();
            foreach (var definition in listDefinitions.OrderBy(definition => definition.ClassName, StringComparer.Ordinal)) {
                cancellationToken.ThrowIfCancellationRequested();
                sb.Append('.').Append(definition.ClassName).Append(" { ");
                sb.Append("list-style-type:").Append(definition.ListStyleType).Append(';');
                sb.Append("list-style-position:outside;");
                if (definition.LeftIndentTwips.HasValue && definition.LeftIndentTwips.Value > 0) {
                    sb.Append("padding-left:").Append(FormatListTwips(definition.LeftIndentTwips.Value)).Append(';');
                }
                if (definition.HangingIndentTwips.HasValue && definition.HangingIndentTwips.Value > 0) {
                    sb.Append("--word-list-hanging:").Append(FormatListTwips(definition.HangingIndentTwips.Value)).Append(';');
                }
                sb.Append(" }\n");
            }
            styleElement.TextContent = sb.ToString();
            head.AppendChild(styleElement);
        }

        private static string BuildListDefinitionClass(DocumentTraversal.ListInfo listInfo, string listStyleType) {
            var kind = listInfo.Ordered ? "ol" : "ul";
            var style = SlugListDefinitionPart(listStyleType);
            var indent = listInfo.LeftIndentTwips.HasValue && listInfo.LeftIndentTwips.Value > 0
                ? "-i" + listInfo.LeftIndentTwips.Value.ToString(CultureInfo.InvariantCulture)
                : string.Empty;
            var hanging = listInfo.HangingIndentTwips.HasValue && listInfo.HangingIndentTwips.Value > 0
                ? "-h" + listInfo.HangingIndentTwips.Value.ToString(CultureInfo.InvariantCulture)
                : string.Empty;
            return $"word-list-l{listInfo.Level.ToString(CultureInfo.InvariantCulture)}-{kind}-{style}{indent}{hanging}";
        }

        private static string FormatListTwips(int twips) {
            return (twips / 20d).ToString("0.##", CultureInfo.InvariantCulture) + "pt";
        }

        private static string SlugListDefinitionPart(string value) {
            var sb = new StringBuilder(value.Length);
            foreach (var ch in value.Trim().Trim('\'', '"').ToLowerInvariant()) {
                if ((ch >= 'a' && ch <= 'z') || (ch >= '0' && ch <= '9')) {
                    sb.Append(ch);
                } else if (ch == '-' || ch == '_') {
                    sb.Append(ch);
                } else {
                    sb.Append('-');
                }
            }

            var slug = sb.ToString().Trim('-');
            return string.IsNullOrEmpty(slug) ? "marker" : slug;
        }

        private static void AppendClass(IElement element, string className) {
            var existing = element.GetAttribute("class");
            if (string.IsNullOrWhiteSpace(existing)) {
                element.SetAttribute("class", className);
            } else if (!existing!.Split(new[] { ' ', '\t', '\r', '\n', '\f' }, StringSplitOptions.RemoveEmptyEntries).Contains(className, StringComparer.Ordinal)) {
                element.SetAttribute("class", existing + " " + className);
            }
        }
    }
}
