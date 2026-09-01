namespace OfficeIMO.Reader.Ocr;

/// <summary>Languages available through the easy OCR facade and its checksum-pinned data catalog.</summary>
[Flags]
public enum OfficeOcrLanguage {
    /// <summary>English.</summary>
    English = 1,

    /// <summary>Polish.</summary>
    Polish = 2
}

/// <summary>Converts discoverable OCR language selections for advanced Tesseract APIs.</summary>
public static class OfficeOcrLanguageExtensions {
    private const OfficeOcrLanguage SupportedLanguages = OfficeOcrLanguage.English | OfficeOcrLanguage.Polish;

    /// <summary>Returns the Tesseract expression represented by one or more selected languages.</summary>
    public static string ToTesseractExpression(this OfficeOcrLanguage languages) {
        if (languages == 0) {
            throw new ArgumentOutOfRangeException(nameof(languages), "Select at least one OCR language.");
        }
        if ((languages & ~SupportedLanguages) != 0) {
            throw new ArgumentOutOfRangeException(nameof(languages), languages, "The OCR language selection contains an unsupported value.");
        }

        var values = new List<string>(2);
        if ((languages & OfficeOcrLanguage.English) != 0) values.Add("eng");
        if ((languages & OfficeOcrLanguage.Polish) != 0) values.Add("pol");
        return string.Join("+", values);
    }
}
