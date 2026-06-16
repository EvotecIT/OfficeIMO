using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Word.Rtf;

public static partial class WordRtfConverterExtensions {
    private static void CopyDefaultLanguage(WordDocument source, RtfDocument destination) {
        string? language = source._wordprocessingDocument.MainDocumentPart?
            .StyleDefinitionsPart?
            .Styles?
            .DocDefaults?
            .RunPropertiesDefault?
            .RunPropertiesBaseStyle?
            .Languages?
            .Val?
            .Value;

        destination.Settings.DefaultLanguageId = ToRtfLanguageId(language);
    }

    private static void ApplyDefaultLanguage(RtfDocument source, WordDocument destination) {
        string? language = ToWordLanguageTag(source.Settings.DefaultLanguageId);
        if (language == null) {
            return;
        }

        MainDocumentPart? mainPart = destination._wordprocessingDocument.MainDocumentPart;
        if (mainPart == null) {
            return;
        }

        StyleDefinitionsPart? stylesPart = mainPart.StyleDefinitionsPart;
        if (stylesPart == null) {
            return;
        }

        Styles styles = stylesPart.Styles ??= new Styles();
        DocDefaults defaults = styles.DocDefaults ??= new DocDefaults();
        RunPropertiesDefault runDefaults = defaults.RunPropertiesDefault ??= new RunPropertiesDefault();
        RunPropertiesBaseStyle runProperties = runDefaults.RunPropertiesBaseStyle ??= new RunPropertiesBaseStyle();
        runProperties.Languages ??= new Languages();
        runProperties.Languages.Val = language;
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
        if (language == null || wordRun._run == null) {
            return;
        }

        wordRun._run.RunProperties ??= new RunProperties();
        wordRun._run.RunProperties.Languages ??= new Languages();
        wordRun._run.RunProperties.Languages.Val = language;
    }
}
