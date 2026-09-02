using System.Collections.Generic;
using System.Text;

namespace OfficeIMO.Html.Pdf;

public static partial class PdfHtmlConverterExtensions {
    private const int HtmlEncodingChunkCharacters = 4096;

    private static StringBuilder CreateOutputBuilder(PdfHtmlSaveOptions options) {
        if (!options.MaximumOutputCharacters.HasValue) {
            return new StringBuilder();
        }

        int maximum = options.MaximumOutputCharacters.Value;
        return new StringBuilder(Math.Min(256, maximum), maximum);
    }

    private static string NormalizeOutputNewLinesWithinBudget(string value, PdfHtmlSaveOptions options) {
        string normalized = NormalizeOutputNewLines(value, options.NewLine);
        if (options.MaximumOutputCharacters.HasValue &&
            normalized.Length > options.MaximumOutputCharacters.Value) {
            throw new InvalidOperationException(
                $"Generated HTML exceeded the configured {options.MaximumOutputCharacters.Value:N0}-character output limit while requested newlines were being rendered.");
        }
        return normalized;
    }

    private static void AddHtmlItem(
        List<HtmlItem> items,
        HtmlItem item,
        PdfHtmlSaveOptions options,
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
        PdfHtmlSaveOptions options,
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

    private static void ThrowPageItemLimit(PdfHtmlSaveOptions options) {
        throw new InvalidOperationException(
            $"Generated HTML exceeded the configured {options.MaximumOutputCharacters!.Value:N0}-character output limit while page items were being rendered.");
    }

    private sealed class PdfHtmlOutputCapacityException : Exception {
    }
}
