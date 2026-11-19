using System.Text;

namespace OfficeIMO.Html;

/// <summary>
/// The <c>full-width</c> and <c>full-size-kana</c> parts of CSS <c>text-transform</c>. Both map one
/// UTF-16 code unit to one code unit, so logical text offsets stay aligned with the painted text.
/// </summary>
internal static class HtmlTextWidthTransform {
    internal static string Apply(string text, bool fullWidth, bool fullSizeKana, bool fullWidthSpaces) {
        if (text.Length == 0 || !fullWidth && !fullSizeKana) return text;
        StringBuilder? builder = null;
        for (int index = 0; index < text.Length; index++) {
            char source = text[index];
            char mapped = source;
            if (fullSizeKana) mapped = ToFullSizeKana(mapped);
            if (fullWidth && (fullWidthSpaces || mapped != ' ')) mapped = ToFullWidth(mapped);
            if (mapped == source) {
                builder?.Append(source);
                continue;
            }
            builder ??= new StringBuilder(text, 0, index, text.Length);
            builder.Append(mapped);
        }
        return builder?.ToString() ?? text;
    }

    // Unicode's <wide> forms for ASCII and the halfwidth signs, and the inverse of the
    // <narrow> decompositions for halfwidth katakana, hangul and symbols.
    private static char ToFullWidth(char value) {
        if (value == ' ') return '　';
        if (value >= '!' && value <= '~') return (char)(value + 0xFEE0);
        switch (value) {
            case '¢': return '￠';
            case '£': return '￡';
            case '¬': return '￢';
            case '¯': return '￣';
            case '¦': return '￤';
            case '¥': return '￥';
            case '₩': return '￦';
            case '⦅': return '｟';
            case '⦆': return '｠';
        }
        // The <narrow> mappings themselves: compatibility normalization would continue past them,
        // turning compatibility jamo such as U+3131 into conjoining jamo.
        if (value >= '\uFF61' && value <= '\uFF9F') return HalfwidthKatakana[value - 0xFF61];
        if (value == '\uFFA0') return '\u3164';
        if (value >= '\uFFA1' && value <= '\uFFBE') return (char)(value - 0xFFA1 + 0x3131);
        if (value >= '\uFFC2' && value <= '\uFFC7') return (char)(value - 0xFFC2 + 0x314F);
        if (value >= '\uFFCA' && value <= '\uFFCF') return (char)(value - 0xFFCA + 0x3155);
        if (value >= '\uFFD2' && value <= '\uFFD7') return (char)(value - 0xFFD2 + 0x315B);
        if (value >= '\uFFDA' && value <= '\uFFDC') return (char)(value - 0xFFDA + 0x3161);
        if (value >= '\uFFE8' && value <= '\uFFEE') return HalfwidthSymbols[value - 0xFFE8];
        return value;
    }

    // U+FF61 to U+FF9F: halfwidth CJK punctuation, katakana and sound marks.
    private const string HalfwidthKatakana =
        "\u3002\u300C\u300D\u3001\u30FB\u30F2\u30A1\u30A3\u30A5\u30A7\u30A9\u30E3\u30E5\u30E7\u30C3\u30FC"
        + "\u30A2\u30A4\u30A6\u30A8\u30AA\u30AB\u30AD\u30AF\u30B1\u30B3\u30B5\u30B7\u30B9\u30BB\u30BD\u30BF"
        + "\u30C1\u30C4\u30C6\u30C8\u30CA\u30CB\u30CC\u30CD\u30CE\u30CF\u30D2\u30D5\u30D8\u30DB\u30DE\u30DF"
        + "\u30E0\u30E1\u30E2\u30E4\u30E6\u30E8\u30E9\u30EA\u30EB\u30EC\u30ED\u30EF\u30F3\u3099\u309A";

    // U+FFE8 to U+FFEE: halfwidth forms light vertical, arrows, black square and white circle.
    private const string HalfwidthSymbols = "\u2502\u2190\u2191\u2192\u2193\u25A0\u25CB";

    // Small kana and their full-size counterparts in the Basic Multilingual Plane.
    private static char ToFullSizeKana(char value) => value switch {
        'ぁ' => 'あ', 'ぃ' => 'い', 'ぅ' => 'う', 'ぇ' => 'え', 'ぉ' => 'お',
        'っ' => 'つ', 'ゃ' => 'や', 'ゅ' => 'ゆ', 'ょ' => 'よ', 'ゎ' => 'わ',
        'ゕ' => 'か', 'ゖ' => 'け',
        'ァ' => 'ア', 'ィ' => 'イ', 'ゥ' => 'ウ', 'ェ' => 'エ', 'ォ' => 'オ',
        'ッ' => 'ツ', 'ャ' => 'ヤ', 'ュ' => 'ユ', 'ョ' => 'ヨ', 'ヮ' => 'ワ',
        'ヵ' => 'カ', 'ヶ' => 'ケ',
        'ㇰ' => 'ク', 'ㇱ' => 'シ', 'ㇲ' => 'ス', 'ㇳ' => 'ト', 'ㇴ' => 'ヌ',
        'ㇵ' => 'ハ', 'ㇶ' => 'ヒ', 'ㇷ' => 'フ', 'ㇸ' => 'ヘ', 'ㇹ' => 'ホ',
        'ㇺ' => 'ム', 'ㇻ' => 'ラ', 'ㇼ' => 'リ', 'ㇽ' => 'ル', 'ㇾ' => 'レ',
        'ㇿ' => 'ロ',
        'ｧ' => 'ｱ', 'ｨ' => 'ｲ', 'ｩ' => 'ｳ', 'ｪ' => 'ｴ', 'ｫ' => 'ｵ',
        'ｬ' => 'ﾔ', 'ｭ' => 'ﾕ', 'ｮ' => 'ﾖ', 'ｯ' => 'ﾂ',
        _ => value
    };
}
