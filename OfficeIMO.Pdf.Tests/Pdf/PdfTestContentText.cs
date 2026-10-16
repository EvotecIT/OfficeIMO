using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Pdf;

namespace OfficeIMO.Tests.Pdf;

/// <summary>Locates encoded text in content streams independently of TJ kerning and run grouping.</summary>
internal static class PdfTestContentText {
    private static readonly Regex TextArrayAdjustment = new(@">\s*-?(?:\d+\.?\d*|\.\d+)\s*<", RegexOptions.CultureInvariant);

    /// <summary>Displays object dictionaries and decoded streams for assertions; the result is not PDF serialization.</summary>
    internal static string ReadDecoded(byte[] pdf) {
        string raw = PdfEncoding.Latin1GetString(pdf);
        var (objects, _) = PdfSyntax.ParseObjects(pdf);
        var result = new StringBuilder();
        int cursor = 0;
        foreach (PdfIndirectObject item in objects.Values.OrderBy(value => value.SourceOffset)) {
            if (item.SourceOffset < cursor || item.Value is not PdfStream stream ||
                !stream.Dictionary.Items.ContainsKey("Filter") ||
                stream.Dictionary.Get<PdfName>("Subtype")?.Name == "Image") continue;
            int marker = raw.IndexOf("\nstream", item.SourceOffset, StringComparison.Ordinal);
            if (marker < 0) throw new InvalidOperationException("The test stream envelope is missing.");
            int start = marker + "\nstream".Length;
            if (raw[start] == '\r') start++;
            if (raw[start] == '\n') start++;
            byte[] decoded = OfficeIMO.Pdf.Filters.StreamDecoder.Decode(stream.Dictionary, stream.Data, objects);
            result.Append(raw, cursor, start - cursor).Append(PdfEncoding.Latin1GetString(decoded));
            cursor = start + stream.DataLength;
        }
        result.Append(raw, cursor, raw.Length - cursor);
        return result.ToString();
    }

    /// <summary>Matches a BaseFont/FontName entry with or without an embedded-subset tag (ABCDEF+).</summary>
    internal static string BaseFontPattern(string name) => "/BaseFont /(?:[A-Z]{6}\\+)?" + Regex.Escape(name);

    /// <summary>Removes an embedded-subset tag (ABCDEF+) from a font name.</summary>
    internal static string? WithoutSubsetTag(string? name) =>
        name != null && name.Length > 7 && name[6] == '+' && name.Take(6).All(static c => c >= 'A' && c <= 'Z') ? name.Substring(7) : name;

    /// <summary>Joins hex strings that a text-show array separates with numeric adjustments.</summary>
    internal static string JoinTextArrays(string content) => TextArrayAdjustment.Replace(content, string.Empty);
}
