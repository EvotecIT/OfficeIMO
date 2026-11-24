using System.Globalization;
using System.Text;

namespace OfficeIMO.Notebook.Import;

public static partial class NotebookPowerShellScript {
    private const string Language = "powershell";
    private const string HeaderPrefix = "# OfficeIMO PowerShell Notebook ";
    private const string Header = HeaderPrefix + "v1 ";
    private const string CellPrefix = "# %% [";
    private const string MarkdownPrefix = "# ";
    private static readonly Encoding Utf16LittleEndian = new UnicodeEncoding(false, false, true);
    private static readonly Encoding Utf16BigEndian = new UnicodeEncoding(true, false, true);

    internal static void Read(ReadOnlyMemory<byte> input, NotebookImportContext context) {
        string text = DecodeScript(input);
        NotebookScriptImporter.SetLanguageInfo(context.Document, Language);
        int cursor = 0;
        ReadOnlySpan<char> line = ReadLine(text.AsSpan(), ref cursor);
        if (!line.StartsWith(HeaderPrefix.AsSpan(), StringComparison.Ordinal)) {
            context.AddCell(NotebookCellKind.Code, text, Language);
            return;
        }
        if (!line.StartsWith(Header.AsSpan(), StringComparison.Ordinal) ||
            !TryReadNumber(line.Slice(Header.Length), out int salt)) {
            throw new FormatException("The PowerShell notebook header has an unsupported version or invalid boundary number.");
        }

        string marker = CellMarker(salt);
        string lineEnding = text.Substring(line.Length, cursor - line.Length);
        int contentStart = cursor;
        NotebookCellKind kind = default;
        string? cellId = null;
        while (cursor < text.Length) {
            context.CancellationToken.ThrowIfCancellationRequested();
            int lineStart = cursor;
            line = ReadLine(text.AsSpan(), ref cursor);
            if (!line.StartsWith(marker.AsSpan(), StringComparison.Ordinal)) {
                if (cellId == null) throw new FormatException("The PowerShell notebook header must be followed by a cell boundary.");
                continue;
            }

            if (cellId != null) AddCell(context, text.AsSpan(contentStart, lineStart - contentStart), lineEnding, kind, cellId);
            (kind, cellId) = CellDescriptor(line.Slice(marker.Length));
            contentStart = cursor;
        }

        if (cellId != null) AddCell(context, text.AsSpan(contentStart), lineEnding, kind, cellId);
    }

    private static void AddCell(NotebookImportContext context, ReadOnlySpan<char> text, string lineEnding,
        NotebookCellKind kind, string cellId) {
        if (!text.EndsWith(lineEnding.AsSpan(), StringComparison.Ordinal)) {
            throw new FormatException("A PowerShell notebook cell is missing its framing line ending.");
        }
        text = text.Slice(0, text.Length - lineEnding.Length);
        bool isMarkdown = kind == NotebookCellKind.Markdown;
        string source = isMarkdown ? ReadCommentedMarkdown(text, context.CancellationToken) : text.ToString();
        context.AddCell(kind, source, isMarkdown ? null : Language, cellId);
    }

    private static (NotebookCellKind Kind, string Id) CellDescriptor(ReadOnlySpan<char> descriptor) {
        int separator = descriptor.IndexOf(' ');
        if (separator < 0) throw new FormatException("A PowerShell notebook boundary must specify a cell kind and ID.");
        ReadOnlySpan<char> kind = descriptor.Slice(0, separator);
        NotebookCellKind cellKind;
        if (kind.SequenceEqual("code".AsSpan())) cellKind = NotebookCellKind.Code;
        else if (kind.SequenceEqual("markdown".AsSpan())) cellKind = NotebookCellKind.Markdown;
        else throw new FormatException("A PowerShell notebook cell kind is unsupported.");
        string id = descriptor.Slice(separator + 1).ToString();
        if (!NotebookParser.IsValidId(id)) throw new FormatException("A PowerShell notebook cell ID is invalid.");
        return (cellKind, id);
    }

    private static string ReadCommentedMarkdown(ReadOnlySpan<char> text, CancellationToken cancellationToken) {
        var markdown = new StringBuilder(text.Length);
        int cursor = 0;
        while (cursor < text.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            int start = cursor;
            if (!ReadLine(text, ref cursor).StartsWith(MarkdownPrefix.AsSpan(), StringComparison.Ordinal)) {
                throw new FormatException("A PowerShell notebook Markdown line must begin with '# '.");
            }
            int contentStart = start + MarkdownPrefix.Length;
            int length = cursor - contentStart;
#if NET8_0_OR_GREATER
            markdown.Append(text.Slice(contentStart, length));
#else
            markdown.Append(text.Slice(contentStart, length).ToString());
#endif
        }
        return markdown.ToString();
    }

    private static void WriteCommentedMarkdown(StringBuilder script, string source, CancellationToken cancellationToken) {
        int cursor = 0;
        while (cursor < source.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            int start = cursor;
            ReadLine(source.AsSpan(), ref cursor);
            script.Append(MarkdownPrefix).Append(source, start, cursor - start);
            CheckOutputLength(script.Length);
        }
    }

    private static void ReserveBoundaries(string source, HashSet<int> reserved, CancellationToken cancellationToken) {
        int cursor = 0;
        while (cursor < source.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            ReadOnlySpan<char> line = ReadLine(source.AsSpan(), ref cursor);
            if (!line.StartsWith(CellPrefix.AsSpan(), StringComparison.Ordinal)) continue;
            ReadOnlySpan<char> suffix = line.Slice(CellPrefix.Length);
            int closing = suffix.IndexOf(']');
            if (closing > 0 && TryReadNumber(suffix.Slice(0, closing), out int number)) reserved.Add(number);
        }
    }

    private static bool TryReadNumber(ReadOnlySpan<char> value, out int number) {
        number = 0;
        if (value.Length is 0 or > 10) return false;
        foreach (char character in value) {
            int digit = character - '0';
            if (digit is < 0 or > 9 || number > (int.MaxValue - digit) / 10) return false;
            number = (number * 10) + digit;
        }
        return true;
    }

    private static string CellMarker(int salt) => CellPrefix + salt.ToString(CultureInfo.InvariantCulture) + "] ";

    private static ReadOnlySpan<char> ReadLine(ReadOnlySpan<char> text, ref int cursor) {
        ReadOnlySpan<char> rest = text.Slice(cursor);
        int ending = rest.IndexOfAny('\r', '\n');
        if (ending < 0) {
            cursor = text.Length;
            return rest;
        }

        cursor += ending + 1;
        if (rest[ending] == '\r' && cursor < text.Length && text[cursor] == '\n') cursor++;
        return rest.Slice(0, ending);
    }

    private static string DecodeScript(ReadOnlyMemory<byte> input) {
        ReadOnlySpan<byte> prefix = input.Span;
        if (prefix.Length < 2) return NotebookImportContext.DecodeUtf8(input);
        Encoding? encoding = prefix[0] == 0xFF && prefix[1] == 0xFE ? Utf16LittleEndian
            : prefix[0] == 0xFE && prefix[1] == 0xFF ? Utf16BigEndian : null;
        if (encoding == null) return NotebookImportContext.DecodeUtf8(input);
        if (prefix.Length >= 4 && (prefix[0] == 0xFF && prefix[2] == 0 && prefix[3] == 0)) {
            throw new FormatException("UTF-32 PowerShell scripts are unsupported. Use UTF-8 or BOM-marked UTF-16.");
        }

        try {
#if NET8_0_OR_GREATER
            return encoding.GetString(prefix.Slice(2));
#else
            if (System.Runtime.InteropServices.MemoryMarshal.TryGetArray(input, out ArraySegment<byte> segment)) {
                return encoding.GetString(segment.Array!, segment.Offset + 2, segment.Count - 2);
            }
            return encoding.GetString(input.Slice(2).ToArray());
#endif
        } catch (DecoderFallbackException error) {
            throw new FormatException("The PowerShell script contains invalid UTF-16.", error);
        }
    }
}
