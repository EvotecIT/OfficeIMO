namespace OfficeIMO.Email;

/// <summary>
/// Resolves and applies the ordered code-page fallbacks used by MAPI PT_STRING8 properties.
/// </summary>
internal sealed class MapiStringEncodingContext {
    private readonly int[] _candidateCodePages;

    static MapiStringEncodingContext() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    private MapiStringEncodingContext(IEnumerable<int> candidateCodePages) {
        _candidateCodePages = candidateCodePages
            .Where(IsUsableString8CodePage)
            .Distinct()
            .DefaultIfEmpty(1252)
            .ToArray();
    }

    internal int PrimaryCodePage => _candidateCodePages[0];

    internal static MapiStringEncodingContext FromCodePage(int? codePage) {
        return new MapiStringEncodingContext(new[] { codePage.GetValueOrDefault(1252), 1252 });
    }

    internal static MapiStringEncodingContext Resolve(
        byte[] propertyStream,
        int headerLength,
        MapiStringEncodingContext? inherited) {
        var candidates = new List<int>();
        AddPropertyCodePage(candidates, propertyStream, headerLength, MapiKnownProperties.PidTag.InternetCodepage);
        AddPropertyCodePage(candidates, propertyStream, headerLength, MapiKnownProperties.PidTag.CodePageId);
        AddPropertyCodePage(candidates, propertyStream, headerLength, MapiKnownProperties.PidTag.MessageCodepage);

        int? localeId = ReadInt32Property(propertyStream, headerLength, MapiKnownProperties.PidTag.MessageLocaleId);
        if (localeId > 0) {
            try {
                candidates.Add(CultureInfo.GetCultureInfo(localeId.Value).TextInfo.ANSICodePage);
            } catch (CultureNotFoundException) {
                // Invalid producer locale: continue with inherited and deterministic defaults.
            }
        }

        if (inherited != null) candidates.AddRange(inherited._candidateCodePages);
        candidates.Add(1252);
        return new MapiStringEncodingContext(candidates);
    }

    internal string Decode(byte[] bytes, IList<EmailDiagnostic> diagnostics, string location) {
        string? best = null;
        int bestReplacementCount = int.MaxValue;

        for (int index = 0; index < _candidateCodePages.Length; index++) {
            int codePage = _candidateCodePages[index];
            try {
                Encoding strict = Encoding.GetEncoding(
                    codePage,
                    EncoderFallback.ExceptionFallback,
                    DecoderFallback.ExceptionFallback);
                string decoded = strict.GetString(bytes);
                if (index > 0) {
                    diagnostics.Add(new EmailDiagnostic(
                        "EMAIL_MSG_STRING8_CODEPAGE_FALLBACK",
                        string.Concat("PT_STRING8 content used code page ", codePage.ToString(CultureInfo.InvariantCulture),
                            " after the preferred code page could not decode it."),
                        EmailDiagnosticSeverity.Warning,
                        location));
                }
                return decoded;
            } catch (DecoderFallbackException) {
                string decoded = DecodeWithReplacement(codePage, bytes);
                int replacements = CountReplacementCharacters(decoded);
                if (replacements < bestReplacementCount) {
                    best = decoded;
                    bestReplacementCount = replacements;
                }
            } catch (ArgumentException) {
                // Unsupported producer code page: try the next deterministic candidate.
            } catch (NotSupportedException) {
                // Unsupported producer code page: try the next deterministic candidate.
            }
        }

        diagnostics.Add(new EmailDiagnostic(
            "EMAIL_MSG_STRING8_CODEPAGE_RECOVERED",
            "PT_STRING8 content could not be decoded strictly; replacement decoding was used.",
            EmailDiagnosticSeverity.Warning,
            location));
        return best ?? DecodeWithReplacement(1252, bytes);
    }

    private static void AddPropertyCodePage(List<int> candidates, byte[] stream, int headerLength,
        MapiPropertyKey<int> key) {
        int? value = ReadInt32Property(stream, headerLength, key);
        if (value.HasValue) candidates.Add(value.Value);
    }

    private static int? ReadInt32Property(byte[] stream, int headerLength, MapiPropertyKey<int> key) {
        int count = Math.Max(0, stream.Length - headerLength) / 16;
        for (int index = 0; index < count; index++) {
            int offset = headerLength + index * 16;
            uint tag = MsgBinary.ReadUInt32(stream, offset);
            if (key.MatchesIdentity((ushort)(tag >> 16)) && key.Accepts((MapiPropertyType)(ushort)tag)) {
                return MsgBinary.ReadInt32(stream, offset + 8);
            }
        }
        return null;
    }

    private static bool IsUsableString8CodePage(int codePage) {
        if (codePage <= 0 || codePage == 1200 || codePage == 1201) return false;
        try {
            _ = Encoding.GetEncoding(codePage, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);
            return true;
        } catch (ArgumentException) {
            return false;
        } catch (NotSupportedException) {
            return false;
        }
    }

    private static string DecodeWithReplacement(int codePage, byte[] bytes) {
        try {
            return Encoding.GetEncoding(
                codePage,
                EncoderFallback.ReplacementFallback,
                DecoderFallback.ReplacementFallback).GetString(bytes);
        } catch (ArgumentException) {
            return Encoding.GetEncoding(1252).GetString(bytes);
        } catch (NotSupportedException) {
            return Encoding.GetEncoding(1252).GetString(bytes);
        }
    }

    private static int CountReplacementCharacters(string value) {
        int count = 0;
        for (int index = 0; index < value.Length; index++) {
            if (value[index] == '\uFFFD') count++;
        }
        return count;
    }
}
