#if NET8_0_OR_GREATER
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.Intrinsics;
using System.Text;

namespace OfficeIMO.CSV;

internal static partial class CsvWriter
{
    /// <summary>Rows that may need more room than this in the staging buffer are appended field by field.</summary>
    private const int MaxStagedTextRowLength = 1 << 20;

    /// <summary>
    /// Appends a row of text under the default quoting rule as <see cref="AppendRecordDefault(StringBuilder, string?[], char, string)"/>
    /// does, but escapes fields into one staging buffer and appends it with as few builder calls as it can: short fields
    /// with the one-pass and vector escapers, longer ones holding quotes with a vector search and a copy that doubles
    /// quotes eight units at a time, and longer ones without quotes straight from the string.
    /// </summary>
    internal static void AppendRecordDefault(StringBuilder buffer, string?[] values, char delimiter, string newLine, ref char[]? staging)
    {
        long room = newLine.Length + values.Length;
        foreach (string? value in values) room += (value?.Length ?? 0) * 2L + 2;
        if (room > MaxStagedTextRowLength)
        {
            AppendRecordDefault(buffer, values, delimiter, newLine);
            return;
        }

        if (staging == null || staging.Length < room)
        {
            staging = new char[Math.Max((int)room, Math.Max(1024, (staging?.Length ?? 0) * 2))];
        }

        Span<char> destination = staging;
        int position = 0;
        for (int index = 0; index < values.Length; index++)
        {
            if (index > 0) destination[position++] = delimiter;
            string? text = values[index];
            if (text == null) continue;
            if (text.Length > EscapedShortTextLength)
            {
                // Long text without quotes to double goes to the builder as it is, so it is copied once, not twice.
                int special = IndexOfCsvSpecial(text, delimiter);
                if (special < 0 || text.IndexOf('"', special) < 0)
                {
                    if (special >= 0) destination[position++] = '"';
                    buffer.Append(destination.Slice(0, position));
                    buffer.Append(text);
                    position = 0;
                    if (special >= 0) destination[position++] = '"';
                    continue;
                }
            }

            position += WriteEscapedTextField(text, destination.Slice(position), delimiter);
        }

        newLine.AsSpan().CopyTo(destination.Slice(position));
        buffer.Append(destination.Slice(0, position + newLine.Length));
    }

    /// <summary>Writes one field under the default quoting rule into room for twice its length and two quotes.</summary>
    private static int WriteEscapedTextField(string text, Span<char> destination, char delimiter)
    {
        if (delimiter == ',' && text.Length <= EscapedShortTextLength)
        {
            if (text.Length <= ShortTextLength && TryCopyPlainText(text, destination, out _)) return text.Length;
            if (TryWriteEscapedShortText(text, destination, out int written)) return written;
        }

        if (IndexOfCsvSpecial(text, delimiter) < 0)
        {
            text.AsSpan().CopyTo(destination);
            return text.Length;
        }

        destination[0] = '"';
        int count = 1 + DoubleQuotes(text.AsSpan(), destination.Slice(1));
        destination[count] = '"';
        return count + 1;
    }

    /// <summary>
    /// Copies text with every quote doubled: eight units at a time while they hold no quote, one at a time where they do.
    /// <paramref name="destination"/> holds at least twice the text length.
    /// </summary>
    private static int DoubleQuotes(ReadOnlySpan<char> text, Span<char> destination)
    {
        int read = 0;
        int written = 0;
        if (Vector128.IsHardwareAccelerated)
        {
            ref ushort source = ref Unsafe.As<char, ushort>(ref MemoryMarshal.GetReference(text));
            Vector128<ushort> quote = Vector128.Create((ushort)'"');
            while (read <= text.Length - Vector128<ushort>.Count)
            {
                Vector128<ushort> units = Vector128.LoadUnsafe(ref source, (nuint)read);
                if (!Vector128.EqualsAny(units, quote))
                {
                    units.CopyTo(MemoryMarshal.Cast<char, ushort>(destination.Slice(written, Vector128<ushort>.Count)));
                    read += Vector128<ushort>.Count;
                    written += Vector128<ushort>.Count;
                    continue;
                }

                for (int end = read + Vector128<ushort>.Count; read < end; read++)
                {
                    char character = text[read];
                    destination[written++] = character;
                    if (character == '"') destination[written++] = '"';
                }
            }
        }

        for (; read < text.Length; read++)
        {
            char character = text[read];
            destination[written++] = character;
            if (character == '"') destination[written++] = '"';
        }

        return written;
    }
}
#endif
