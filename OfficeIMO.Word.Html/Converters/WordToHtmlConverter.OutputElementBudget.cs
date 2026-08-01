using AngleSharp.Dom;
using System.Runtime.CompilerServices;

namespace OfficeIMO.Word.Html {
    internal partial class WordToHtmlConverter {
        private static readonly HashSet<string> HtmlVoidElements = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
            "area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "param", "source", "track", "wbr"
        };

        private static readonly ConditionalWeakTable<IDocument, OutputConstructionBudget> OutputConstructionBudgets =
            new ConditionalWeakTable<IDocument, OutputConstructionBudget>();

        private static void RegisterOutputConstructionBudget(
            IDocument owner,
            WordToHtmlOptions options,
            long initialOutputCharacters) {
            OutputConstructionBudgets.Add(owner, new OutputConstructionBudget(options, initialOutputCharacters));
        }

        private static IElement CreateOutputElement(IDocument owner, string tagName) {
            if (!OutputConstructionBudgets.TryGetValue(owner, out OutputConstructionBudget? budget)) {
                throw new InvalidOperationException("The HTML output construction budget was not initialized.");
            }
            return budget.CreateElement(owner, tagName);
        }

        private static void ReserveOutputCharacters(
            IDocument owner,
            long characters,
            string message,
            string source) {
            if (!OutputConstructionBudgets.TryGetValue(owner, out OutputConstructionBudget? budget)) {
                throw new InvalidOperationException("The HTML output construction budget was not initialized.");
            }
            budget.ReserveCharacters(characters, message, source);
        }

        private static long GetRemainingOutputCharacters(IDocument owner) {
            if (!OutputConstructionBudgets.TryGetValue(owner, out OutputConstructionBudget? budget)) {
                throw new InvalidOperationException("The HTML output construction budget was not initialized.");
            }
            return budget.RemainingCharacters;
        }

        private sealed class OutputConstructionBudget {
            private readonly WordToHtmlOptions _options;
            private long _minimumOutputCharacters;

            internal OutputConstructionBudget(WordToHtmlOptions options, long initialOutputCharacters) {
                _options = options;
                _minimumOutputCharacters = initialOutputCharacters;
            }

            internal long RemainingCharacters => Math.Max(0, _options.MaxOutputCharacters - _minimumOutputCharacters);

            internal IElement CreateElement(IDocument owner, string tagName) {
                long minimumSerializedCharacters = HtmlVoidElements.Contains(tagName)
                    ? tagName.Length + 2L
                    : (tagName.Length * 2L) + 5L;
                ReserveCharacters(
                    minimumSerializedCharacters,
                    "Generated HTML elements exceed the configured output-character limit before DOM construction.",
                    "GeneratedElement:" + tagName);
                return owner.CreateElement(tagName);
            }

            internal void ReserveCharacters(long characters, string message, string source) {
                long projectedCharacters = SaturatingAdd(_minimumOutputCharacters, characters);
                if (projectedCharacters > _options.MaxOutputCharacters) {
                    ThrowExportLimitExceeded(
                        _options,
                        "WordHtmlOutputLimitExceeded",
                        message,
                        source,
                        projectedCharacters,
                        _options.MaxOutputCharacters);
                }
                _minimumOutputCharacters = projectedCharacters;
            }
        }
    }
}
