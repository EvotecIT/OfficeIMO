using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Word.OpenXml;
using static OfficeIMO.Word.OpenXml.WordXmlEdits;

namespace OfficeIMO.Word;

public partial class WordParagraph {
    /// <summary>
    /// Gets or sets directly authored hidden formatting on this run. True hides the
    /// run, false overrides an inherited hidden setting, and null restores style
    /// inheritance. The getter reports direct formatting, not the effective style.
    /// </summary>
    public bool? Hidden {
        get {
            OfficeXmlNode hidden = RunPropertiesNode.Child(W, "vanish");
            return hidden.IsNone ? null : ReadOnOff(hidden);
        }
        set {
            if (value == null) {
                if (!RunPropertiesNode.Child(W, "vanish").IsNone) RemoveChild(EditRunProperties(), "vanish");
                return;
            }
            ReplaceChild(EditRunProperties(), "vanish", WRPrView.Order, value.Value ? "true" : "false");
        }
    }
}
