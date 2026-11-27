using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    // The SDK side of the paragraph's element handles in WordParagraph.Elements.cs: the SDK elements, passed as object so
    // the engine names no SDK type.
    public partial class WordParagraph {
        private object SdkParagraphElement => _paragraph;

        private bool SdkHasRun => _run != null;

        private object? SdkOwnRunPropertiesElement => _runProperties;

        private object? SdkParagraphPropertiesElement => _paragraphProperties;

        private object SdkEditParagraphPropertiesElement() => _paragraph.ParagraphProperties ??= new ParagraphProperties();

        private object? SdkScopedRunPropertiesElement => ScopedRunProperties;

        private object? SdkRunPropertiesIgnoringContentControlElement => IsHyperLink ? Hyperlink?._runProperties : _runProperties;

        private object SdkEditRunPropertiesElement(bool hyperlinkInContentControl) {
            if (IsHyperLink && (hyperlinkInContentControl || _stdRun == null)) {
                WordHyperLink hyperlink = Hyperlink!;
                return VerifyRunProperties(hyperlink._hyperlink!, hyperlink._run!, hyperlink._runProperties);
            }
            return VerifyRunProperties();
        }
    }
}
