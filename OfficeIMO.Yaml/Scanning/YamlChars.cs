using System.Runtime.CompilerServices;

namespace OfficeIMO.Yaml;

/// <summary>Character classes from the YAML 1.2.2 productions and the vectorized searches the scanner's hot loops use.</summary>
internal static class YamlChars {
    // Flags for the ASCII class table.
    private const byte Control = 1;      // c-printable violation (C0 except tab/LF/CR, and DEL)
    private const byte Break = 2;        // b-char
    private const byte FlowIndicator = 4;// , [ ] { }
    private const byte Word = 8;         // ns-word-char
    private const byte Uri = 16;         // ns-uri-char (without %)
    private const byte PlainStop = 32;   // ':' and '#'
    private const byte DoubleStop = 64;  // '"' and '\\'
    private const byte SingleStop = 128; // '\''

    private static readonly byte[] s_table = CreateTable();

#if NET8_0_OR_GREATER
    private const string Controls = "\u0000\u0001\u0002\u0003\u0004\u0005\u0006\u0007\u0008\u000B\u000C\u000E\u000F\u0010\u0011\u0012\u0013\u0014\u0015\u0016\u0017\u0018\u0019\u001A\u001B\u001C\u001D\u001E\u001F\u007F";
    private static readonly SearchValues<char> s_plainBlock = SearchValues.Create(":#\r\n" + Controls);
    private static readonly SearchValues<char> s_plainFlow = SearchValues.Create(":#\r\n,[]{}" + Controls);
    private static readonly SearchValues<char> s_double = SearchValues.Create("\"\\\r\n" + Controls);
    private static readonly SearchValues<char> s_single = SearchValues.Create("'\r\n" + Controls);
    private static readonly SearchValues<char> s_lineEnd = SearchValues.Create("\r\n" + Controls);
#endif

    private static byte[] CreateTable() {
        var table = new byte[128];
        for (int i = 0; i < 0x20; i++) table[i] = Control;
        table['\t'] = 0;
        table['\n'] = Break;
        table['\r'] = Break;
        table[0x7F] = Control;
        foreach (char c in ",[]{}") table[c] |= FlowIndicator;
        for (char c = '0'; c <= '9'; c++) table[c] |= Word | Uri;
        for (char c = 'a'; c <= 'z'; c++) table[c] |= Word | Uri;
        for (char c = 'A'; c <= 'Z'; c++) table[c] |= Word | Uri;
        table['-'] |= Word | Uri;
        foreach (char c in "#;/?:@&=+$,_.!~*'()[]") table[c] |= Uri;
        table[':'] |= PlainStop;
        table['#'] |= PlainStop;
        table['"'] |= DoubleStop;
        table['\\'] |= DoubleStop;
        table['\''] |= SingleStop;
        return table;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool IsBreak(char c) => c == '\n' || c == '\r';

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool IsBlank(char c) => c == ' ' || c == '\t';

    /// <summary>Space, tab, line break, or the end-of-input sentinel.</summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool IsBlankOrBreakOrNul(char c) => c == ' ' || c == '\t' || c == '\n' || c == '\r' || c == '\0';

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool IsFlowIndicator(char c) => c < 128 && (s_table[c] & FlowIndicator) != 0;

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool IsWord(char c) => c < 128 && (s_table[c] & Word) != 0;

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool IsUri(char c) => c < 128 && (s_table[c] & Uri) != 0;

    /// <summary>ns-tag-char without the percent escape: URI characters except '!' and flow indicators.</summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool IsTagChar(char c) => c < 128 && (s_table[c] & (Uri | FlowIndicator)) == Uri && c != '!';

    /// <summary>Characters outside c-printable that the scanner rejects (C0 controls other than tab and breaks, DEL, U+FFFE/U+FFFF).</summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool IsNonPrintable(char c) => c < 128 ? (s_table[c] & Control) != 0 : c is '￾' or '￿';

    /// <summary>c-indicator: characters that cannot start a plain scalar (with the '-', '?', ':' exceptions handled by the caller).</summary>
    public static bool IsIndicator(char c) => c switch {
        '-' or '?' or ':' or ',' or '[' or ']' or '{' or '}' or '#' or '&' or '*' or '!' or '|' or '>' or '\'' or '"' or '%' or '@' or '`' => true,
        _ => false
    };

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static int HexValue(char c) {
        if (c >= '0' && c <= '9') return c - '0';
        if (c >= 'a' && c <= 'f') return c - 'a' + 10;
        if (c >= 'A' && c <= 'F') return c - 'A' + 10;
        return -1;
    }

    /// <summary>Finds the next character that may end or interrupt a plain scalar line: ':', '#', a line break, a control character, or (in flow context) a flow indicator.</summary>
    public static int IndexOfPlainStop(ReadOnlySpan<char> s, bool flow) {
#if NET8_0_OR_GREATER
        // Most plain scalars are short keys and values; a table lookup over the first characters finds their end
        // without the setup of a vectorized search, which takes over for longer text.
        int mask = Control | Break | PlainStop | (flow ? FlowIndicator : 0);
        byte[] table = s_table;
        int head = Math.Min(s.Length, 16);
        for (int i = 0; i < head; i++) {
            char c = s[i];
            if (c < 128 && (table[c] & mask) != 0) return i;
        }
        if (head == s.Length) return -1;
        int index = s.Slice(head).IndexOfAny(flow ? s_plainFlow : s_plainBlock);
        return index < 0 ? -1 : head + index;
#else
        int mask = Control | Break | PlainStop | (flow ? FlowIndicator : 0);
        return IndexOfClass(s, (byte)mask);
#endif
    }

    /// <summary>Finds the next '"', '\\', line break or control character.</summary>
    public static int IndexOfDoubleQuotedStop(ReadOnlySpan<char> s) {
#if NET8_0_OR_GREATER
        return s.IndexOfAny(s_double);
#else
        return IndexOfClass(s, Control | Break | DoubleStop);
#endif
    }

    /// <summary>Finds the next '\'', line break or control character.</summary>
    public static int IndexOfSingleQuotedStop(ReadOnlySpan<char> s) {
#if NET8_0_OR_GREATER
        return s.IndexOfAny(s_single);
#else
        return IndexOfClass(s, Control | Break | SingleStop);
#endif
    }

    /// <summary>Finds the next line break or control character.</summary>
    public static int IndexOfLineEnd(ReadOnlySpan<char> s) {
#if NET8_0_OR_GREATER
        return s.IndexOfAny(s_lineEnd);
#else
        return IndexOfClass(s, Control | Break);
#endif
    }

#if !NET8_0_OR_GREATER
    private static int IndexOfClass(ReadOnlySpan<char> s, byte mask) {
        byte[] table = s_table;
        for (int i = 0; i < s.Length; i++) {
            char c = s[i];
            if (c < 128 && (table[c] & mask) != 0) return i;
        }
        return -1;
    }
#endif
}
