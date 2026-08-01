using AngleSharp.Dom;
using System.Runtime.CompilerServices;

namespace OfficeIMO.Word.Html {
    internal partial class WordToHtmlConverter {
        private static readonly HashSet<string> HtmlVoidElements = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
            "area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "param", "source", "track", "wbr"
        };

        private static readonly ConditionalWeakTable<IDocument, OutputElementBudget> OutputElementBudgets =
            new ConditionalWeakTable<IDocument, OutputElementBudget>();

        private static void RegisterOutputElementBudget(
            IDocument owner,
            WordToHtmlOptions options,
            long initialOutputCharacters) {
            OutputElementBudgets.Add(owner, new OutputElementBudget(options, initialOutputCharacters));
        }

        private static IElement CreateOutputElement(IDocument owner, string tagName) {
            if (!OutputElementBudgets.TryGetValue(owner, out OutputElementBudget? budget)) {
                throw new InvalidOperationException("The HTML output element budget was not initialized.");
            }
            return budget.CreateElement(owner, tagName);
        }

        private sealed class OutputElementBudget {
            private readonly WordToHtmlOptions _options;
            private long _minimumOutputCharacters;

            internal OutputElementBudget(WordToHtmlOptions options, long initialOutputCharacters) {
                _options = options;
                _minimumOutputCharacters = initialOutputCharacters;
            }

            internal IElement CreateElement(IDocument owner, string tagName) {
                long minimumSerializedCharacters = HtmlVoidElements.Contains(tagName)
                    ? tagName.Length + 2L
                    : (tagName.Length * 2L) + 5L;
                long projectedCharacters = SaturatingAdd(_minimumOutputCharacters, minimumSerializedCharacters);
                if (projectedCharacters > _options.MaxOutputCharacters) {
                    ThrowExportLimitExceeded(
                        _options,
                        "WordHtmlOutputLimitExceeded",
                        "Generated HTML elements exceed the configured output-character limit before DOM construction.",
                        "GeneratedElement:" + tagName,
                        projectedCharacters,
                        _options.MaxOutputCharacters);
                }
                _minimumOutputCharacters = projectedCharacters;
                return owner.CreateElement(tagName);
            }
        }
    }
}
