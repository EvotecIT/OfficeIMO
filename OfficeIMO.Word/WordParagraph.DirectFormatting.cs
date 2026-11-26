using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Word.OpenXml;
using static OfficeIMO.Word.OpenXml.WordXmlEdits;

namespace OfficeIMO.Word;

public partial class WordParagraph {
    /// <summary>Maps authored emphasis overrides from document-format adapters, including explicit off values beneath styles.</summary>
    internal void ApplyDirectEmphasis(bool? bold, bool? italic, WordUnderlineStyle? underline) {
        OfficeXmlNode properties = EditRunProperties();
        SetOptionalOnOff(properties, "b", bold);
        SetOptionalOnOff(properties, "bCs", bold);
        SetOptionalOnOff(properties, "i", italic);
        SetOptionalOnOff(properties, "iCs", italic);
        if (underline.HasValue) ReplaceChild(properties, "u", WRPrView.Order, WordUnderlineStyleTokens.ToToken(underline.Value));
        else RemoveChild(properties, "u");
    }

    private static void SetOptionalOnOff(OfficeXmlNode properties, string name, bool? value) {
        if (value.HasValue) ReplaceChild(properties, name, WRPrView.Order, value.Value ? "true" : "false");
        else RemoveChild(properties, name);
    }
}
