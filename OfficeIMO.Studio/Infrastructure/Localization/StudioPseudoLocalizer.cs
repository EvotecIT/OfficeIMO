using System.Text;

namespace OfficeIMO.Studio.Infrastructure.Localization;

internal static class StudioPseudoLocalizer {
    private static readonly IReadOnlyDictionary<char, char> Accents = new Dictionary<char, char> {
        ['a'] = 'à', ['A'] = 'À', ['b'] = 'ƀ', ['B'] = 'Ɓ', ['c'] = 'ç', ['C'] = 'Ç',
        ['d'] = 'đ', ['D'] = 'Đ', ['e'] = 'ë', ['E'] = 'Ë', ['f'] = 'ƒ', ['F'] = 'Ƒ',
        ['g'] = 'ĝ', ['G'] = 'Ĝ', ['h'] = 'ĥ', ['H'] = 'Ĥ', ['i'] = 'ï', ['I'] = 'Ï',
        ['j'] = 'ĵ', ['J'] = 'Ĵ', ['k'] = 'ķ', ['K'] = 'Ķ', ['l'] = 'ł', ['L'] = 'Ł',
        ['m'] = 'ḿ', ['M'] = 'Ḿ', ['n'] = 'ñ', ['N'] = 'Ñ', ['o'] = 'ö', ['O'] = 'Ö',
        ['p'] = 'þ', ['P'] = 'Þ', ['q'] = 'ɋ', ['Q'] = 'Ɋ', ['r'] = 'ŕ', ['R'] = 'Ŕ',
        ['s'] = 'š', ['S'] = 'Š', ['t'] = 'ŧ', ['T'] = 'Ŧ', ['u'] = 'ü', ['U'] = 'Ü',
        ['v'] = 'ṽ', ['V'] = 'Ṽ', ['w'] = 'ŵ', ['W'] = 'Ŵ', ['x'] = 'ẋ', ['X'] = 'Ẋ',
        ['y'] = 'ÿ', ['Y'] = 'Ÿ', ['z'] = 'ž', ['Z'] = 'Ž'
    };

    internal static string Transform(string text) {
        if (string.IsNullOrEmpty(text)) return text;
        var output = new StringBuilder(text.Length + Math.Max(8, text.Length / 3));
        output.Append('⟦');
        bool inFormatItem = false;
        foreach (char character in text) {
            if (character == '{') inFormatItem = true;
            if (!inFormatItem && Accents.TryGetValue(character, out char accented)) output.Append(accented);
            else output.Append(character);
            if (character == '}') inFormatItem = false;
        }
        output.Append(" ···⟧");
        return output.ToString();
    }
}
