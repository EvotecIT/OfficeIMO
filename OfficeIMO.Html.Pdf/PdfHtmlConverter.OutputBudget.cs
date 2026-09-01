using System.Collections.Generic;
using System.Text;

namespace OfficeIMO.Html.Pdf;

public static partial class PdfHtmlConverterExtensions {
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
}
