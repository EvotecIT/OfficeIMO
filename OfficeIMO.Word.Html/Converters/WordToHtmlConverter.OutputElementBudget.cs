using AngleSharp.Dom;

namespace OfficeIMO.Word.Html {
    internal partial class WordToHtmlConverter {
        private sealed class OutputElementBudget {
            private readonly WordToHtmlOptions _options;
            private long _minimumOutputCharacters;

            internal OutputElementBudget(WordToHtmlOptions options, long initialOutputCharacters) {
                _options = options;
                _minimumOutputCharacters = initialOutputCharacters;
            }

            internal IElement CreateElement(IDocument owner, string tagName) {
                // Even an empty element serializes with an opening and closing tag.
                long minimumSerializedCharacters = (tagName.Length * 2L) + 5L;
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
