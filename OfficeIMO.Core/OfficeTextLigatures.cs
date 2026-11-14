namespace OfficeIMO.Drawing;

/// <summary>Dependency-free Unicode presentation-ligature helpers for text renderers.</summary>
internal static class OfficeTextLigatures {
    /// <summary>
    /// Resolves the standard Latin <c>ff</c>, <c>fi</c>, <c>fl</c>, <c>ffi</c>, or <c>ffl</c>
    /// presentation form at a UTF-16 index, preferring the longest match.
    /// </summary>
    public static bool TryGetLatinPresentationForm(
        string text,
        int index,
        out int ligatureScalar,
        out int utf16Length) {
        ArgumentNullException.ThrowIfNull(text);
        // Every supported ligature starts with 'f'.
        if (index < 0 || index >= text.Length || text[index] != 'f') {
            ligatureScalar = 0;
            utf16Length = 0;
            return false;
        }

        if (StartsWithOrdinal(text, index, "ffi")) {
            ligatureScalar = 0xFB03;
            utf16Length = 3;
            return true;
        }
        if (StartsWithOrdinal(text, index, "ffl")) {
            ligatureScalar = 0xFB04;
            utf16Length = 3;
            return true;
        }
        if (StartsWithOrdinal(text, index, "ff")) {
            ligatureScalar = 0xFB00;
            utf16Length = 2;
            return true;
        }
        if (StartsWithOrdinal(text, index, "fi")) {
            ligatureScalar = 0xFB01;
            utf16Length = 2;
            return true;
        }
        if (StartsWithOrdinal(text, index, "fl")) {
            ligatureScalar = 0xFB02;
            utf16Length = 2;
            return true;
        }

        ligatureScalar = 0;
        utf16Length = 0;
        return false;
    }

    private static bool StartsWithOrdinal(string text, int index, string value) {
        if (index > text.Length - value.Length) return false;
        // An ordinal comparison of a few characters; string.Compare is slow for this on .NET Framework.
        for (int offset = 0; offset < value.Length; offset++) {
            if (text[index + offset] != value[offset]) return false;
        }
        return true;
    }
}
