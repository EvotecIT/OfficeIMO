namespace OfficeIMO.Rtf;

/// <summary>Scoped decoding state shared by body text, metadata, and encapsulated content.</summary>
internal class RtfTextDecodingState {
    public int AnsiCodePage { get; set; } = RtfAnsiCodePage.DefaultWindowsCodePage;
    public int UnicodeSkipCount { get; set; } = 1;
    public int SkipCharacters { get; set; }
    public char? PendingHighSurrogate { get; set; }
    public byte? PendingAnsiLeadByte { get; set; }
}

/// <summary>Decodes RTF source bytes and Unicode fallback without treating source newlines as content.</summary>
internal static class RtfTextDecoder {
    public static bool ConsumeFallback(RtfTextDecodingState state, int count = 1) {
        if (state.SkipCharacters <= 0) return false;
        state.SkipCharacters = Math.Max(0, state.SkipCharacters - count);
        return true;
    }

    public static string DecodeAnsiText(string text, RtfTextDecodingState state) {
        if (string.IsNullOrEmpty(text)) return string.Empty;
        var result = new StringBuilder(text.Length);
        int start = 0;
        while (start < text.Length) {
            char character = text[start];
            if (character == '\r' || character == '\n') {
                start++;
                continue;
            }
            if (state.SkipCharacters > 0) {
                ConsumeFallback(state);
                start++;
                continue;
            }
            // Decode contiguous segments together; retain a final DBCS lead byte for the next token.
            int end = start + 1;
            while (end < text.Length && text[end] != '\r' && text[end] != '\n') end++;
            int length = end - start;
            if (state.PendingAnsiLeadByte.HasValue) {
                if (character <= byte.MaxValue) {
                    result.Append(DecodeAnsiByte(character, state));
                    start++;
                    length--;
                } else {
                    result.Append(FlushAnsi(state));
                }
            }
            if (length > 0) {
                int last = start + length - 1;
                bool retainLead = text[last] <= byte.MaxValue &&
                    RtfAnsiCodePage.IsLeadByte(state.AnsiCodePage, (byte)text[last]) &&
                    EndsWithUnpairedLead(text, start, length, state.AnsiCodePage);
                if (retainLead) length--;
                if (length > 0) {
                    result.Append(FlushUnicode(state));
                    result.Append(RtfAnsiCodePage.DecodeText(state.AnsiCodePage, text.Substring(start, length)));
                }
                if (retainLead) state.PendingAnsiLeadByte = (byte)text[last];
            }
            start = end;
        }
        return result.ToString();
    }

    private static bool EndsWithUnpairedLead(string text, int start, int length, int codePage) {
        int end = start + length;
        for (int index = start; index < end; index++) {
            if (text[index] <= byte.MaxValue && RtfAnsiCodePage.IsLeadByte(codePage, (byte)text[index])) {
                if (index == end - 1) return true;
                index++;
            }
        }
        return false;
    }

    public static string DecodeAnsiByte(int value, RtfTextDecodingState state) {
        if (ConsumeFallback(state)) return string.Empty;
        byte current = (byte)(value & 0xFF);
        if (state.PendingAnsiLeadByte.HasValue) {
            byte lead = state.PendingAnsiLeadByte.Value;
            state.PendingAnsiLeadByte = null;
            return FlushUnicode(state) + RtfAnsiCodePage.DecodeBytes(state.AnsiCodePage, new[] { lead, current });
        }
        if (RtfAnsiCodePage.IsLeadByte(state.AnsiCodePage, current)) {
            state.PendingAnsiLeadByte = current;
            return string.Empty;
        }
        return FlushUnicode(state) + RtfAnsiCodePage.DecodeByte(state.AnsiCodePage, current);
    }

    public static string DecodeUnicode(int value, RtfTextDecodingState state) {
        string prefix = FlushAnsi(state);
        char codeUnit = unchecked((char)value);
        string result;
        if (char.IsHighSurrogate(codeUnit)) {
            result = prefix + FlushUnicode(state);
            state.PendingHighSurrogate = codeUnit;
        } else if (char.IsLowSurrogate(codeUnit)) {
            result = prefix + (state.PendingHighSurrogate.HasValue
                ? new string(new[] { state.PendingHighSurrogate.Value, codeUnit })
                : "\uFFFD");
            state.PendingHighSurrogate = null;
        } else {
            result = prefix + FlushUnicode(state) + codeUnit;
        }
        state.SkipCharacters = state.UnicodeSkipCount;
        return result;
    }

    public static string DecodeLiteral(string text, RtfTextDecodingState state) =>
        ConsumeFallback(state) ? string.Empty : Flush(state) + text;

    public static string Flush(RtfTextDecodingState state) => FlushAnsi(state) + FlushUnicode(state);

    private static string FlushAnsi(RtfTextDecodingState state) {
        if (!state.PendingAnsiLeadByte.HasValue) return string.Empty;
        state.PendingAnsiLeadByte = null;
        return "\uFFFD";
    }

    private static string FlushUnicode(RtfTextDecodingState state) {
        if (!state.PendingHighSurrogate.HasValue) return string.Empty;
        state.PendingHighSurrogate = null;
        return "\uFFFD";
    }
}
