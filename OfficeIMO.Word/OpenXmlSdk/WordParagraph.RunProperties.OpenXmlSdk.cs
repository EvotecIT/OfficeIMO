using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    // The SDK side of run formatting that still needs SDK types: the run properties a paragraph scopes to (its run, its
    // hyperlink's run or its content control's first run), and theme colours, whose cached value is resolved against
    // the SDK's theme part.
    public partial class WordParagraph {
        internal RunProperties? ScopedRunProperties => _stdRun != null
            ? _stdRun.SdtContentRun?.Descendants<Run>().FirstOrDefault()?.GetFirstChild<RunProperties>()
            : IsHyperLink ? Hyperlink?._runProperties : _runProperties;

        private WordThemeColor? SdkThemeColor {
            set {
                RunProperties runProperties;
                if (IsHyperLink && _stdRun == null) {
                    var hyperlink = this.Hyperlink!;
                    runProperties = VerifyRunProperties(hyperlink._hyperlink!, hyperlink._run!, hyperlink._runProperties);
                } else {
                    runProperties = VerifyRunProperties();
                }
                if (value != null) {
                    var color = new DocumentFormat.OpenXml.Wordprocessing.Color {
                        ThemeColor = new EnumValue<ThemeColorValues> {
                            Value = value.Value.ToOpenXml()
                        }
                    };
                    // w:val is required; it caches the theme colour for applications that ignore the theme.
                    color.Val = WordRunFormatting.ResolveColorHex(color, _document) ?? "auto";
                    runProperties.Color = color;
                } else {
                    if (runProperties.Color != null) runProperties.Color.Remove();
                }
            }
        }
    }
}
