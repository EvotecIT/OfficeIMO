using System.Globalization;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    private static bool AreNativeOpenTypeFeaturesEnabled(WordDocument? document) {
        W.Settings? settings = document?._wordprocessingDocument?.MainDocumentPart?.DocumentSettingsPart?.Settings;
        string? value = settings?.GetFirstChild<W.Compatibility>()?.Elements<W.CompatibilitySetting>()
            .LastOrDefault(setting => setting.Name?.InnerText == "enableOpenTypeFeatures" &&
                setting.Uri?.Value == "http://schemas.microsoft.com/office/word")?.Val?.Value;
        if (value == "0" || value == "false" || value == "off") return false;
        if (value == "1" || value == "true" || value == "on") return true;
        return document != null && UsesModernNativeWordLayout(document);
    }

    /// <summary>Resolves the effective Word layout mode while retaining stored compatibility flags.</summary>
    private static bool UsesModernNativeWordLayout(WordDocument document) {
        if (document.SourceFormat == WordFileFormat.Doc) return false;
        W.Settings? settings = document._wordprocessingDocument?.MainDocumentPart?.DocumentSettingsPart?.Settings;
        string? modeText = settings?.GetFirstChild<W.Compatibility>()?.Elements<W.CompatibilitySetting>()
            .LastOrDefault(setting => setting.Name?.Value == W.CompatSettingNameValues.CompatibilityMode)?.Val?.Value;
        return int.TryParse(modeText, NumberStyles.Integer, CultureInfo.InvariantCulture, out int mode) && mode >= 15;
    }
}
