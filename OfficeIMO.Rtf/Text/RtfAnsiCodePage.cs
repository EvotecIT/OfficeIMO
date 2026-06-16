namespace OfficeIMO.Rtf;

internal static partial class RtfAnsiCodePage {
    public const int MacRomanCodePage = 10000;
    public const int IbmPcCodePage = 437;
    public const int IbmPcaCodePage = 850;
    public const int DefaultWindowsCodePage = 1252;

    public static bool IsSupported(int codePage) =>
        codePage == IbmPcCodePage ||
        codePage == IbmPcaCodePage ||
        codePage == 874 ||
        codePage == MacRomanCodePage ||
        (codePage >= 1250 && codePage <= 1258);

    public static int GetDefaultCodePage(RtfDocumentCharacterSet? characterSet) {
        return characterSet switch {
            RtfDocumentCharacterSet.Mac => MacRomanCodePage,
            RtfDocumentCharacterSet.Pc => IbmPcCodePage,
            RtfDocumentCharacterSet.Pca => IbmPcaCodePage,
            _ => DefaultWindowsCodePage
        };
    }

    public static string DecodeText(int codePage, string text) {
        if (string.IsNullOrEmpty(text)) return text;

        StringBuilder? builder = null;
        for (int i = 0; i < text.Length; i++) {
            char ch = text[i];
            string decoded = ch <= byte.MaxValue ? DecodeByte(codePage, ch) : ch.ToString();
            if (builder == null) {
                if (decoded.Length == 1 && decoded[0] == ch) {
                    continue;
                }

                builder = new StringBuilder(text.Length);
                builder.Append(text, 0, i);
            }

            builder.Append(decoded);
        }

        return builder?.ToString() ?? text;
    }

    public static string DecodeByte(int codePage, int value) {
        int b = value & 0xFF;
        int mapped = codePage switch {
            IbmPcCodePage => DecodeIbmPc437(b),
            IbmPcaCodePage => DecodeIbmPc850(b),
            874 => DecodeWindows874(b),
            1250 => DecodeWindows1250(b),
            1251 => DecodeWindows1251(b),
            1252 => DecodeWindows1252(b),
            1253 => DecodeWindows1253(b),
            1254 => DecodeWindows1254(b),
            1255 => DecodeWindows1255(b),
            1256 => DecodeWindows1256(b),
            1257 => DecodeWindows1257(b),
            1258 => DecodeWindows1258(b),
            MacRomanCodePage => DecodeMacRoman(b),
            _ => DecodeWindows1252(b)
        };

        return char.ConvertFromUtf32(mapped);
    }
}
