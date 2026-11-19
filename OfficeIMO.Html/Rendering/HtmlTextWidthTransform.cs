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
        if (value >= '｡' && value <= 'ￜ' || value >= '￨' && value <= '￮') {
            string normalized = value.ToString().Normalize(NormalizationForm.FormKC);
            if (normalized.Length == 1) return normalized[0];
        }
        return value;
    }

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
