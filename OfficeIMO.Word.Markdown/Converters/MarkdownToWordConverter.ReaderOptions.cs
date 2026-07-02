using Omd = OfficeIMO.Markdown;

namespace OfficeIMO.Word.Markdown {
    internal partial class MarkdownToWordConverter {
        private static Omd.MarkdownReaderOptions CreateEffectiveReaderOptions(MarkdownToWordOptions options) {
            var source = options.ReaderOptions;
            if (source == null) {
                var defaults = new Omd.MarkdownReaderOptions {
                    BaseUri = options.BaseUri,
                    PreferNarrativeSingleLineDefinitions = options.PreferNarrativeSingleLineDefinitions
                };
                WordMarkdownSemanticBlocks.ConfigureReaderOptions(defaults);
                return defaults;
            }

            var effective = source.Clone();

            if (string.IsNullOrWhiteSpace(effective.BaseUri) && !string.IsNullOrWhiteSpace(options.BaseUri)) {
                effective.BaseUri = options.BaseUri;
            }

            if (!effective.PreferNarrativeSingleLineDefinitions && options.PreferNarrativeSingleLineDefinitions) {
                effective.PreferNarrativeSingleLineDefinitions = true;
            }

            WordMarkdownSemanticBlocks.ConfigureReaderOptions(effective);
            return effective;
        }

    }
}
