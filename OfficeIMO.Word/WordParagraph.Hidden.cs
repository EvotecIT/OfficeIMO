using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordParagraph {
    /// <summary>
    /// Gets or sets directly authored hidden formatting on this run. True hides the
    /// run, false overrides an inherited hidden setting, and null restores style
    /// inheritance. The getter reports direct formatting, not the effective style.
    /// </summary>
    public bool? Hidden {
        get {
            Vanish? hidden = ScopedRunProperties?.Vanish;
            return hidden == null ? null : hidden.Val?.Value ?? true;
        }
        set {
            if (value == null) {
                ScopedRunProperties?.Vanish?.Remove();
                return;
            }

            RunProperties properties = IsHyperLink && _stdRun == null
                ? VerifyRunProperties(Hyperlink!._hyperlink!, Hyperlink._run!, Hyperlink._runProperties)
                : VerifyRunProperties();
            properties.Vanish = new Vanish { Val = value.Value };
        }
    }
}
