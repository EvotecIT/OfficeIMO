using System.Globalization;
using System.Text;

#if OFFICEIMO_READER_CORE
namespace OfficeIMO.Reader.Internal.TextAnalysis;
#else
namespace OfficeIMO.TextAnalysis;
#endif

/// <summary>Unicode normalization applied to each token before folding.</summary>
#if OFFICEIMO_READER_CORE
internal
#else
public
#endif
enum OfficeTextNormalization {
    /// <summary>Tokens keep their source code points.</summary>
    None,
    /// <summary>Canonical composition (NFC).</summary>
    Nfc,
    /// <summary>Compatibility composition (NFKC), which also unifies ligatures, full-width forms and similar variants.</summary>
    Nfkc
}

/// <summary>Controls how <see cref="OfficeTextAnalyzer"/> turns text into index terms.</summary>
#if OFFICEIMO_READER_CORE
internal
#else
public
#endif
sealed class OfficeTextAnalyzerOptions {
    /// <summary>Unicode normalization of each token. Default: <see cref="OfficeTextNormalization.Nfkc"/>.</summary>
    public OfficeTextNormalization Normalization { get; set; } = OfficeTextNormalization.Nfkc;

    /// <summary>When true (the default), terms are case folded so that "Résumé", "RÉSUMÉ" and "résumé" match.</summary>
    public bool CaseFolding { get; set; } = true;

    /// <summary>When true, combining marks are removed so that "résumé" also matches "resume". Default: false.</summary>
    public bool AccentFolding { get; set; }

    /// <summary>Tokens longer than this many UTF-16 code units are skipped. Default: 256.</summary>
    public int MaximumTokenLength { get; set; } = 256;

    internal OfficeTextAnalyzerOptions Clone() => new() {
        Normalization = Normalization,
        CaseFolding = CaseFolding,
        AccentFolding = AccentFolding,
        MaximumTokenLength = MaximumTokenLength
    };
}

/// <summary>One analysed token with its term and its location in the source text.</summary>
#if OFFICEIMO_READER_CORE
internal
#else
public
#endif
readonly struct OfficeTextToken {
    /// <summary>Creates a token.</summary>
    public OfficeTextToken(string term, int start, int length, int position) {
        Term = term ?? throw new ArgumentNullException(nameof(term));
        Start = start;
        Length = length;
        Position = position;
    }

    /// <summary>Normalized, folded index term.</summary>
    public string Term { get; }

    /// <summary>Zero-based UTF-16 offset of the source word.</summary>
    public int Start { get; }

    /// <summary>Source word length in UTF-16 code units.</summary>
    public int Length { get; }

    /// <summary>Zero-based token position, used for phrase matching.</summary>
    public int Position { get; }

    /// <inheritdoc />
    public override string ToString() => Term + "@" + Position + "[" + Start + "+" + Length + "]";
}

/// <summary>
/// Splits text into UAX #29 word tokens and normalizes them into index terms: Unicode normalization, case folding
/// and optional accent folding. Punctuation, whitespace, symbols and emoji are not tokens.
/// </summary>
#if OFFICEIMO_READER_CORE
internal
#else
public
#endif
sealed class OfficeTextAnalyzer {
    private readonly OfficeTextAnalyzerOptions _options;

    /// <summary>Creates an analyser. The options are copied.</summary>
    public OfficeTextAnalyzer(OfficeTextAnalyzerOptions? options = null) {
        _options = (options ?? new OfficeTextAnalyzerOptions()).Clone();
        if (_options.MaximumTokenLength < 1) {
            throw new ArgumentOutOfRangeException(nameof(options), "MaximumTokenLength must be positive.");
        }
    }

    /// <summary>Returns the tokens of <paramref name="text"/> in source order.</summary>
    public IReadOnlyList<OfficeTextToken> Analyze(string text) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        var tokens = new List<OfficeTextToken>();
        foreach (UnicodeTextSegment segment in UnicodeTextSegmenter.GetWordSegments(text)) {
            if (!segment.IsWordLike || segment.Length > _options.MaximumTokenLength) continue;
            string term = NormalizeTerm(text.Substring(segment.Start, segment.Length));
            if (term.Length == 0) continue;
            tokens.Add(new OfficeTextToken(term, segment.Start, segment.Length, tokens.Count));
        }
        return tokens.AsReadOnly();
    }

    /// <summary>
    /// Applies normalization and folding to one word without segmenting it. Used for prefixes, which may end inside
    /// a word.
    /// </summary>
    public string NormalizeTerm(string word) {
        if (word == null) throw new ArgumentNullException(nameof(word));
        if (IsAscii(word)) {
            return _options.CaseFolding ? word.ToLowerInvariant() : word;
        }

        string value = Normalize(word, _options.Normalization);
        if (_options.CaseFolding) value = FoldCase(value);
        if (_options.AccentFolding) value = FoldAccents(value);
        return value;
    }

    private static bool IsAscii(string value) {
        foreach (char character in value) {
            if (character > 0x7F) return false;
        }
        return true;
    }

    private static string Normalize(string value, OfficeTextNormalization normalization) {
        NormalizationForm form;
        switch (normalization) {
            case OfficeTextNormalization.Nfc: form = NormalizationForm.FormC; break;
            case OfficeTextNormalization.Nfkc: form = NormalizationForm.FormKC; break;
            default: return value;
        }
        try {
            return value.IsNormalized(form) ? value : value.Normalize(form);
        } catch (Exception exception) when (exception is PlatformNotSupportedException or ArgumentException) {
            // Globalization-invariant hosts and ill-formed UTF-16 keep the source code points.
            return value;
        }
    }

    // Simple case folding through the invariant upper and lower mappings, so final sigma, long s and similar
    // variants fold together; the sharp s pair folds to "ss" as in full case folding.
    private static string FoldCase(string value) {
        string folded = value.ToUpperInvariant().ToLowerInvariant();
        return folded.IndexOf('ß') >= 0 ? folded.Replace("ß", "ss") : folded;
    }

    private static string FoldAccents(string value) {
        string decomposed;
        try {
            decomposed = value.Normalize(NormalizationForm.FormD);
        } catch (Exception exception) when (exception is PlatformNotSupportedException or ArgumentException) {
            decomposed = value;
        }
        var builder = new StringBuilder(decomposed.Length);
        foreach (char character in decomposed) {
            UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(character);
            if (category == UnicodeCategory.NonSpacingMark || category == UnicodeCategory.EnclosingMark) continue;
            builder.Append(FoldLetterWithoutDecomposition(character));
        }
        string result = builder.ToString();
        try {
            return result.Normalize(NormalizationForm.FormC);
        } catch (Exception exception) when (exception is PlatformNotSupportedException or ArgumentException) {
            return result;
        }
    }

    // Latin letters whose marks are not separable by canonical decomposition.
    private static char FoldLetterWithoutDecomposition(char value) => value switch {
        'ł' => 'l', 'Ł' => 'L',
        'ø' => 'o', 'Ø' => 'O',
        'đ' => 'd', 'Đ' => 'D',
        'ħ' => 'h', 'Ħ' => 'H',
        'ŧ' => 't', 'Ŧ' => 'T',
        'ı' => 'i',
        _ => value
    };
}
