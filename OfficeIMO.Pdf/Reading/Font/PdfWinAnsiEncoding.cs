namespace OfficeIMO.Pdf;

internal static class PdfWinAnsiEncoding {
    private const string UnsupportedTextMessageSuffix = "cannot be encoded with PDF WinAnsiEncoding. OfficeIMO.Pdf currently writes standard PDF fonts for generated text; embedded Unicode fonts are required for this text.";

    // Windows-1252 mapping (aka WinAnsiEncoding in PDF). Index is byte value.
    private static readonly char[] Map = new char[256] {
        '\u0000','\u0001','\u0002','\u0003','\u0004','\u0005','\u0006','\u0007','\b','\t','\n','\u000B','\f','\r','\u000E','\u000F',
        '\u0010','\u0011','\u0012','\u0013','\u0014','\u0015','\u0016','\u0017','\u0018','\u0019','\u001A','\u001B','\u001C','\u001D','\u001E','\u001F',
        ' ','!','"','#','$','%','&','\'','(',')','*','+',',','-','.','/',
        '0','1','2','3','4','5','6','7','8','9',':',';','<','=','>','?',
        '@','A','B','C','D','E','F','G','H','I','J','K','L','M','N','O',
        'P','Q','R','S','T','U','V','W','X','Y','Z','[','\\',']','^','_',
        '`','a','b','c','d','e','f','g','h','i','j','k','l','m','n','o',
        'p','q','r','s','t','u','v','w','x','y','z','{','|','}','~','\u007F',
        '€','\u0081','‚','ƒ','„','…','†','‡','ˆ','‰','Š','‹','Œ','\u008D','Ž','\u008F',
        '\u0090','‘','’','“','”','•','–','—','˜','™','š','›','œ','\u009D','ž','Ÿ',
        '\u00A0','¡','¢','£','¤','¥','¦','§','¨','©','ª','«','¬','\u00AD','®','¯',
        '°','±','²','³','´','µ','¶','·','¸','¹','º','»','¼','½','¾','¿',
        'À','Á','Â','Ã','Ä','Å','Æ','Ç','È','É','Ê','Ë','Ì','Í','Î','Ï',
        'Ð','Ñ','Ò','Ó','Ô','Õ','Ö','×','Ø','Ù','Ú','Û','Ü','Ý','Þ','ß',
        'à','á','â','ã','ä','å','æ','ç','è','é','ê','ë','ì','í','î','ï',
        'ð','ñ','ò','ó','ô','õ','ö','÷','ø','ù','ú','û','ü','ý','þ','ÿ'
    };

    private static readonly System.Collections.Generic.Dictionary<char, byte> ReverseMap = BuildReverse();

    public static string Decode(byte[] bytes) {
        return Decode(bytes, int.MaxValue);
    }

    public static string Decode(byte[] bytes, int maxOutputCharacters) {
        if (bytes.LongLength > maxOutputCharacters) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.DecodedTextCharacters, maxOutputCharacters, bytes.LongLength);
        }
        var chars = new char[bytes.Length];
        for (int i = 0; i < bytes.Length; i++) chars[i] = Map[bytes[i]];
        return new string(chars);
    }

    public static char Decode(byte value) => Map[value];

    public static byte[] Encode(string s) {
        var bytes = new byte[s.Length];
        for (int i = 0; i < s.Length; i++) {
            var ch = s[i];
            if (!TryGetByte(ch, out var b)) {
                throw CreateUnsupportedCharacterException(s, i);
            }

            bytes[i] = b;
        }
        return bytes;
    }

    public static bool CanEncode(string s, out int unsupportedIndex) {
        for (int i = 0; i < s.Length; i++) {
            if (!TryGetByte(s[i], out _)) {
                unsupportedIndex = i;
                return false;
            }
        }

        unsupportedIndex = -1;
        return true;
    }

    private static bool TryGetByte(char ch, out byte value) {
        if (IsUnsupportedControlCharacter(ch)) {
            value = 0;
            return false;
        }

        if (!ReverseMap.TryGetValue(ch, out value)) {
            return false;
        }

        return value != 0x81 && value != 0x8D && value != 0x8F && value != 0x90 && value != 0x9D;
    }

    private static bool IsUnsupportedControlCharacter(char ch) =>
        ch < ' ' || ch == '\u007F';

    private static ArgumentException CreateUnsupportedCharacterException(string text, int index) {
        string codePoint;
        string display;
        char ch = text[index];
        if (char.IsHighSurrogate(ch) && index + 1 < text.Length && char.IsLowSurrogate(text[index + 1])) {
            int value = char.ConvertToUtf32(ch, text[index + 1]);
            codePoint = "U+" + value.ToString("X", System.Globalization.CultureInfo.InvariantCulture);
            display = new string(new[] { ch, text[index + 1] });
        } else {
            codePoint = "U+" + ((int)ch).ToString("X4", System.Globalization.CultureInfo.InvariantCulture);
            display = char.IsControl(ch) ? string.Empty : ch.ToString();
        }

        if (IsUnsupportedControlCharacter(ch)) {
            return new ArgumentException("Text contains control character " + codePoint + " at index " + index.ToString(System.Globalization.CultureInfo.InvariantCulture) + ". PDF text output cannot render control characters directly; use paragraphs, line breaks, tables, or spacing primitives for layout.", nameof(text));
        }

        string rendered = display.Length == 0 ? string.Empty : " '" + display + "'";
        return new ArgumentException("Text contains character " + codePoint + rendered + " at index " + index.ToString(System.Globalization.CultureInfo.InvariantCulture) + " that " + UnsupportedTextMessageSuffix, nameof(text));
    }

    private static System.Collections.Generic.Dictionary<char, byte> BuildReverse() {
        var dict = new System.Collections.Generic.Dictionary<char, byte>();
        for (int i = 0; i < Map.Length; i++) {
            char c = Map[i];
            if (!dict.ContainsKey(c)) dict[c] = (byte)i;
        }
        return dict;
    }
}
