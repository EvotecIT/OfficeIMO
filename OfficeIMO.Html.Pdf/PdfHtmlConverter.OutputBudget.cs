using System.Collections.Generic;
using System.Text;

namespace OfficeIMO.Html.Pdf;

public static partial class PdfHtmlConverterExtensions {
    private const int HtmlEncodingChunkCharacters = 4096;

    private static StringBuilder CreateOutputBuilder(PdfToHtmlOptions options) {
        if (!options.MaximumOutputCharacters.HasValue) {
            return new StringBuilder();
        }

        int maximum = options.MaximumOutputCharacters.Value;
        // The rendered buffer can temporarily be larger than the requested output when CRLF
        // sequences are normalized to a one-character line ending and when the final line ending
        // is removed. Keep that representation bounded without rejecting a valid final result.
        long rawMaximum = Math.Min(int.MaxValue, checked((long)maximum * 2L + 2L));
        return new StringBuilder(Math.Min(256, maximum), checked((int)rawMaximum));
    }

    private static string NormalizeOutputNewLinesWithinBudget(StringBuilder value, PdfToHtmlOptions options) {
        string newLine = options.NewLine;
        int sourceLength = value.Length;
        while (sourceLength > 0 && value[sourceLength - 1] is '\r' or '\n') sourceLength--;
        long normalizedLength = 0L;
        bool changed = sourceLength != value.Length;
        for (int index = 0; index < sourceLength; index++) {
            char current = value[index];
            if (current is not ('\r' or '\n')) {
                normalizedLength++;
                continue;
            }

            int sourceNewLineLength = current == '\r' && index + 1 < sourceLength && value[index + 1] == '\n' ? 2 : 1;
            changed |= !MatchesNewLine(value, index, sourceNewLineLength, newLine);
            normalizedLength = checked(normalizedLength + newLine.Length);
            index += sourceNewLineLength - 1;
        }
        if (options.MaximumOutputCharacters.HasValue &&
            normalizedLength > options.MaximumOutputCharacters.Value) {
            throw new InvalidOperationException(
                $"Generated HTML exceeded the configured {options.MaximumOutputCharacters.Value:N0}-character output limit while requested newlines were being rendered.");
        }
        if (sourceLength != value.Length) value.Length = sourceLength;
        if (changed) {
            value.Replace("\r\n", "\n");
            value.Replace("\r", "\n");
            if (!string.Equals(newLine, "\n", StringComparison.Ordinal)) value.Replace("\n", newLine);
        }
        return value.ToString();

        static bool MatchesNewLine(StringBuilder source, int index, int length, string expected) {
            if (length != expected.Length) return false;
            for (int offset = 0; offset < length; offset++) {
                if (source[index + offset] != expected[offset]) return false;
            }
            return true;
        }
    }

    private static void AddHtmlItem(
        List<HtmlItem> items,
        HtmlItem item,
        PdfToHtmlOptions options,
        ref long retainedHtmlCharacters) {
        retainedHtmlCharacters = checked(retainedHtmlCharacters + item.Html.Length);
        if (options.MaximumOutputCharacters.HasValue &&
            retainedHtmlCharacters > options.MaximumOutputCharacters.Value) {
            throw new InvalidOperationException(
                $"Generated HTML exceeded the configured {options.MaximumOutputCharacters.Value:N0}-character output limit while page items were being rendered.");
        }

        items.Add(item);
    }

    private static string RenderPageItemWithinBudget(
        PdfToHtmlOptions options,
        long retainedHtmlCharacters,
        Action<StringBuilder> render) {
        StringBuilder builder;
        if (options.MaximumOutputCharacters.HasValue) {
            long remaining = options.MaximumOutputCharacters.Value - retainedHtmlCharacters;
            if (remaining < 1L) ThrowPageItemLimit(options);
            int maximum = (int)Math.Min(int.MaxValue, remaining);
            builder = new StringBuilder(Math.Min(256, maximum), maximum);
        } else {
            builder = new StringBuilder();
        }

        try {
            render(builder);
            return builder.ToString();
        } catch (Exception exception) when (
            options.MaximumOutputCharacters.HasValue &&
            IsOutputBuilderCapacityException(exception)) {
            ThrowPageItemLimit(options);
            return string.Empty;
        }
    }

    private static void AppendHtmlText(StringBuilder builder, string? value) {
        string source = value ?? string.Empty;
        for (int offset = 0; offset < source.Length;) {
            int count = Math.Min(HtmlEncodingChunkCharacters, source.Length - offset);
            if (offset + count < source.Length &&
                char.IsHighSurrogate(source[offset + count - 1]) &&
                char.IsLowSurrogate(source[offset + count])) {
                count--;
            }
            if (count == 0) count = Math.Min(2, source.Length - offset);

            string encoded = System.Net.WebUtility.HtmlEncode(source.Substring(offset, count));
            if (encoded.Length > builder.MaxCapacity - builder.Length) {
                throw new PdfHtmlOutputCapacityException();
            }
            builder.Append(encoded);
            offset += count;
        }
    }

    private static void ThrowPageItemLimit(PdfToHtmlOptions options) {
        throw new InvalidOperationException(
            $"Generated HTML exceeded the configured {options.MaximumOutputCharacters!.Value:N0}-character output limit while page items were being rendered.");
    }

    private sealed class PdfHtmlOutputCapacityException : Exception {
    }
}
