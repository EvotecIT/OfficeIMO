using System.Globalization;
using OfficeIMO.Word;

namespace OfficeIMO.Word.Rtf;

public static partial class WordRtfConverterExtensions {
    private static void CopyDefaultLanguage(WordDocument source, RtfDocument destination) {
        string? language = source.HasStylesPart()
            ? source.DocumentDefaultsNode().ModelChild("rPrDefault").ModelChild("rPr").ModelChild("lang").WVal()
            : null;

        destination.Settings.DefaultLanguageId = ToRtfLanguageId(language);
    }

    private static void ApplyDefaultLanguage(RtfDocument source, WordDocument destination) {
        string? language = ToWordLanguageTag(source.Settings.DefaultLanguageId);
        if (language == null) {
            return;
        }

        if (!destination.HasStylesPart()) {
            return;
        }

        OfficeXmlNode styles = destination.GetOrAddStylesNode();
        OfficeXmlNode defaults = styles.GetOrPrependChild(WordXml.W, "docDefaults");
        OfficeXmlNode runDefaults = defaults.GetOrPrependChild(WordXml.W, "rPrDefault");
        OfficeXmlNode runProperties = runDefaults.GetOrPrependChild(WordXml.W, "rPr");
        runProperties.GetOrAddChild(WordXml.W, "lang", WRPrView.Order).SetAttribute(WordXml.W, "val", language);
    }

    private static int? ToRtfLanguageId(string? languageTag) {
        if (string.IsNullOrWhiteSpace(languageTag) ||
            string.Equals(languageTag, "auto", StringComparison.OrdinalIgnoreCase)) {
            return null;
        }

        try {
            return CultureInfo.GetCultureInfo(languageTag!).LCID;
        } catch (CultureNotFoundException) {
            return null;
        }
    }

    private static string? ToWordLanguageTag(int? languageId) {
        if (!languageId.HasValue || languageId.Value <= 0) {
            return null;
        }

        try {
            return CultureInfo.GetCultureInfo(languageId.Value).Name;
        } catch (CultureNotFoundException) {
            return null;
        }
    }

    private static void SetRunLanguage(WordParagraph wordRun, int languageId) {
        string? language = ToWordLanguageTag(languageId);
        if (language == null || wordRun.RunNode().IsNone) {
            return;
        }

        GetOrAddRunProperties(wordRun).GetOrAddChild(WordXml.W, "lang", WRPrView.Order).SetAttribute(WordXml.W, "val", language);
    }
}
