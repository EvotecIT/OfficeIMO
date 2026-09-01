namespace OfficeIMO.Reader.Ocr;

/// <summary>Languages available through the easy OCR facade and its checksum-pinned data catalog.</summary>
[Flags]
public enum OfficeOcrLanguage : ulong {
    /// <summary>English.</summary>
    English = 1UL << 0,
    /// <summary>Polish.</summary>
    Polish = 1UL << 1,
    /// <summary>Arabic.</summary>
    Arabic = 1UL << 2,
    /// <summary>Chinese, simplified.</summary>
    ChineseSimplified = 1UL << 3,
    /// <summary>Chinese, traditional.</summary>
    ChineseTraditional = 1UL << 4,
    /// <summary>Czech.</summary>
    Czech = 1UL << 5,
    /// <summary>Danish.</summary>
    Danish = 1UL << 6,
    /// <summary>Dutch.</summary>
    Dutch = 1UL << 7,
    /// <summary>Finnish.</summary>
    Finnish = 1UL << 8,
    /// <summary>French.</summary>
    French = 1UL << 9,
    /// <summary>German.</summary>
    German = 1UL << 10,
    /// <summary>Modern Greek.</summary>
    Greek = 1UL << 11,
    /// <summary>Hebrew.</summary>
    Hebrew = 1UL << 12,
    /// <summary>Hindi.</summary>
    Hindi = 1UL << 13,
    /// <summary>Hungarian.</summary>
    Hungarian = 1UL << 14,
    /// <summary>Italian.</summary>
    Italian = 1UL << 15,
    /// <summary>Japanese.</summary>
    Japanese = 1UL << 16,
    /// <summary>Korean.</summary>
    Korean = 1UL << 17,
    /// <summary>Norwegian.</summary>
    Norwegian = 1UL << 18,
    /// <summary>Portuguese.</summary>
    Portuguese = 1UL << 19,
    /// <summary>Romanian.</summary>
    Romanian = 1UL << 20,
    /// <summary>Russian.</summary>
    Russian = 1UL << 21,
    /// <summary>Slovak.</summary>
    Slovak = 1UL << 22,
    /// <summary>Spanish.</summary>
    Spanish = 1UL << 23,
    /// <summary>Swedish.</summary>
    Swedish = 1UL << 24,
    /// <summary>Turkish.</summary>
    Turkish = 1UL << 25,
    /// <summary>Ukrainian.</summary>
    Ukrainian = 1UL << 26,
    /// <summary>Vietnamese.</summary>
    Vietnamese = 1UL << 27
}

/// <summary>Discoverable language catalog for dropdowns, command completion, and validation.</summary>
public static class OfficeOcrLanguages {
    private static readonly LanguageMapping[] Mappings = {
        new(OfficeOcrLanguage.English, "eng"),
        new(OfficeOcrLanguage.Polish, "pol"),
        new(OfficeOcrLanguage.Arabic, "ara"),
        new(OfficeOcrLanguage.ChineseSimplified, "chi_sim"),
        new(OfficeOcrLanguage.ChineseTraditional, "chi_tra"),
        new(OfficeOcrLanguage.Czech, "ces"),
        new(OfficeOcrLanguage.Danish, "dan"),
        new(OfficeOcrLanguage.Dutch, "nld"),
        new(OfficeOcrLanguage.Finnish, "fin"),
        new(OfficeOcrLanguage.French, "fra"),
        new(OfficeOcrLanguage.German, "deu"),
        new(OfficeOcrLanguage.Greek, "ell"),
        new(OfficeOcrLanguage.Hebrew, "heb"),
        new(OfficeOcrLanguage.Hindi, "hin"),
        new(OfficeOcrLanguage.Hungarian, "hun"),
        new(OfficeOcrLanguage.Italian, "ita"),
        new(OfficeOcrLanguage.Japanese, "jpn"),
        new(OfficeOcrLanguage.Korean, "kor"),
        new(OfficeOcrLanguage.Norwegian, "nor"),
        new(OfficeOcrLanguage.Portuguese, "por"),
        new(OfficeOcrLanguage.Romanian, "ron"),
        new(OfficeOcrLanguage.Russian, "rus"),
        new(OfficeOcrLanguage.Slovak, "slk"),
        new(OfficeOcrLanguage.Spanish, "spa"),
        new(OfficeOcrLanguage.Swedish, "swe"),
        new(OfficeOcrLanguage.Turkish, "tur"),
        new(OfficeOcrLanguage.Ukrainian, "ukr"),
        new(OfficeOcrLanguage.Vietnamese, "vie")
    };

    private static readonly OfficeOcrLanguage SupportedMask = Mappings.Aggregate(
        (OfficeOcrLanguage)0,
        static (current, mapping) => current | mapping.Language);

    /// <summary>All typed languages supported by the zero-setup facade.</summary>
    public static IReadOnlyList<OfficeOcrLanguage> Supported { get; } =
        Array.AsReadOnly(Mappings.Select(static mapping => mapping.Language).ToArray());

    internal static OfficeOcrLanguage Mask => SupportedMask;
    internal static IReadOnlyList<LanguageMapping> Entries => Mappings;

    internal sealed class LanguageMapping {
        internal LanguageMapping(OfficeOcrLanguage language, string code) {
            Language = language;
            Code = code;
        }

        internal OfficeOcrLanguage Language { get; }
        internal string Code { get; }
    }
}

/// <summary>Converts discoverable OCR language selections for advanced Tesseract APIs.</summary>
public static class OfficeOcrLanguageExtensions {
    /// <summary>Returns the Tesseract expression represented by one or more selected languages.</summary>
    public static string ToTesseractExpression(this OfficeOcrLanguage languages) {
        if (languages == 0) {
            throw new ArgumentOutOfRangeException(nameof(languages), "Select at least one OCR language.");
        }
        if ((languages & ~OfficeOcrLanguages.Mask) != 0) {
            throw new ArgumentOutOfRangeException(nameof(languages), languages, "The OCR language selection contains an unsupported value.");
        }

        return string.Join(
            "+",
            OfficeOcrLanguages.Entries
                .Where(mapping => (languages & mapping.Language) != 0)
                .Select(static mapping => mapping.Code));
    }
}
