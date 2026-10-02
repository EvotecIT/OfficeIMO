using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.LegacyDoc.Model;

namespace OfficeIMO.Word;

public partial class WordDocument {
    /// <summary>Projects DOC LSPD's signed twips or 240ths-of-a-line into the corresponding OOXML rule.</summary>
    private static void ApplyLegacyDocLineSpacing(SpacingBetweenLines spacing, LegacyDocParagraphFormat format) {
        if (format.LineSpacingTwips is not int value) return;
        spacing.Line = Math.Abs(value).ToString(System.Globalization.CultureInfo.InvariantCulture);
        spacing.LineRule = format.LineSpacingIsMultiple ? LineSpacingRuleValues.Auto :
            value < 0 ? LineSpacingRuleValues.Exact : LineSpacingRuleValues.AtLeast;
    }
}
