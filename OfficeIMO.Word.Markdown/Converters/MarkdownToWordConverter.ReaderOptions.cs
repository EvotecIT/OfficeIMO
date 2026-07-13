using Omd = OfficeIMO.Markdown;

namespace OfficeIMO.Word.Markdown {
    internal partial class MarkdownToWordConverter {
        private static Omd.MarkdownReaderOptions CreateEffectiveReaderOptions(MarkdownToWordOptions options) =>
            options.CreateReaderOptions();

    }
}
