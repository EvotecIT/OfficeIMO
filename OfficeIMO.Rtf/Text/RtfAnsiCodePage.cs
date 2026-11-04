using System.Collections.Concurrent;

namespace OfficeIMO.Rtf;

internal static partial class RtfAnsiCodePage {
    public const int MacRomanCodePage = 10000;
    public const int IbmPcCodePage = 437;
    public const int IbmPcaCodePage = 850;
    public const int DefaultWindowsCodePage = 1252;

    static RtfAnsiCodePage() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    public static bool IsSupported(int codePage) =>
        codePage == IbmPcCodePage ||
        codePage == IbmPcaCodePage ||
        codePage == 874 ||
        IsDoubleByteCodePage(codePage) ||
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

    public static int? GetCodePageForCharset(int? charset) {
        return charset switch {
            77 => MacRomanCodePage,
            128 => 932,
            129 => 949,
            130 => 1361,
            134 => 936,
            136 => 950,
            161 => 1253,
            162 => 1254,
            163 => 1258,
            177 => 1255,
            178 => 1256,
            186 => 1257,
            204 => 1251,
            222 => 874,
            238 => 1250,
            254 => IbmPcCodePage,
            255 => IbmPcCodePage,
            _ => null
        };
    }

    private static readonly ConcurrentDictionary<int, Encoding> Encodings = new();

    /// <summary>Appends the decoded form of a run of source characters without materializing the run.</summary>
    public static void AppendDecodedText(StringBuilder builder, int codePage, string text, int start, int length) {
        if (length <= 0) return;

        if (IsDoubleByteCodePage(codePage)) {
            var bytes = new byte[length];
            for (int index = 0; index < length; index++) {
                bytes[index] = (byte)(text[start + index] & 0xFF);
            }

            builder.Append(DecodeBytes(codePage, bytes));
            return;
        }

        int end = start + length;
        for (int index = start; index < end; index++) {
            char ch = text[index];
            if (ch > byte.MaxValue) {
                builder.Append(ch);
                continue;
            }

            int mapped = MapSingleByte(codePage, ch);
            if (mapped <= char.MaxValue && !char.IsSurrogate((char)mapped)) {
                builder.Append((char)mapped);
            } else {
                builder.Append(char.ConvertFromUtf32(mapped));
            }
        }
    }

    /// <summary>Returns true when every character of a single-byte run decodes to itself.</summary>
    public static bool DecodesToItself(int codePage, string text) {
        if (IsDoubleByteCodePage(codePage)) return false;

        foreach (char ch in text) {
            if (ch <= byte.MaxValue && MapSingleByte(codePage, ch) != ch) {
                return false;
            }
        }

        return true;
    }

    public static string DecodeByte(int codePage, int value) {
        if (IsDoubleByteCodePage(codePage)) {
            return DecodeBytes(codePage, new[] { (byte)(value & 0xFF) });
        }

        return char.ConvertFromUtf32(MapSingleByte(codePage, value));
    }

    private static int MapSingleByte(int codePage, int value) {
        int b = value & 0xFF;
        return codePage switch {
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
    }

    public static bool IsDoubleByteCodePage(int codePage) =>
        codePage == 932 || codePage == 936 || codePage == 949 || codePage == 950 || codePage == 1361;

    public static bool IsLeadByte(int codePage, int value) {
        int b = value & 0xFF;
        if (codePage == 932) {
            return (b >= 0x81 && b <= 0x9F) || (b >= 0xE0 && b <= 0xFC);
        }

        if (codePage == 1361) {
            return b >= 0x84 && b <= 0xD3;
        }

        return (codePage == 936 || codePage == 949 || codePage == 950) && b >= 0x81 && b <= 0xFE;
    }

    public static string DecodeBytes(int codePage, byte[] bytes) {
        ArgumentNullException.ThrowIfNull(bytes);
        // GetEncoding with explicit fallbacks clones the encoding on every call; decoding is thread-safe, so share one.
        return Encodings.GetOrAdd(codePage, static _ => Encoding.GetEncoding(
            _,
            EncoderFallback.ReplacementFallback,
            DecoderFallback.ReplacementFallback)).GetString(bytes);
    }
}
