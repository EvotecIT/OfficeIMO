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
        if (index < 0 || index >= text.Length) {
            ligatureScalar = 0;
            utf16Length = 0;
            return false;
        }

        // Every form starts with 'f'. Text renderers ask at each index of every run, so reject
        // the common case with one character test before comparing the longer sequences.
        if (text[index] == 'f' && index + 1 < text.Length) {
            char second = text[index + 1];
            if (second == 'f') {
                char third = index + 2 < text.Length ? text[index + 2] : default;
                if (third == 'i') return Found(0xFB03, 3, out ligatureScalar, out utf16Length);
                if (third == 'l') return Found(0xFB04, 3, out ligatureScalar, out utf16Length);
                return Found(0xFB00, 2, out ligatureScalar, out utf16Length);
            }
            if (second == 'i') return Found(0xFB01, 2, out ligatureScalar, out utf16Length);
            if (second == 'l') return Found(0xFB02, 2, out ligatureScalar, out utf16Length);
        }

        ligatureScalar = 0;
        utf16Length = 0;
        return false;
    }

    private static bool Found(int scalar, int length, out int ligatureScalar, out int utf16Length) {
        ligatureScalar = scalar;
        utf16Length = length;
        return true;
    }
}
