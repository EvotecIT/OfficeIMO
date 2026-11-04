using Omd = OfficeIMO.Markdown;

namespace OfficeIMO.Word.Markdown {
    internal partial class MarkdownToWordConverter {
        private static Omd.MarkdownLoadOptions CreateEffectiveReaderOptions(MarkdownWordExportOptions options) =>
            options.CreateReaderOptions();

    }
}
