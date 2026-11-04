using System.Globalization;

namespace OfficeIMO.Pdf;

public sealed partial class PdfCreateOptions {
    /// <summary>Records actual font-program payload cost once per emitted stream, excluding PDF dictionaries and maps.</summary>
    internal void ReportFontEmbedding(string source, string fontName, string format, int objectId,
        long originalBytes, long fontFileBytes, long streamBytes, bool compressed, string embeddingMode) {
        if (_diagnosticsReport == null) return;
        _diagnosticsReport.Add(new PdfConversionWarning(
            _diagnosticsConverter, "font-embedding-cost", source,
            "Font '" + fontName + "' embeds " + streamBytes.ToString(CultureInfo.InvariantCulture) +
            " font-stream bytes (" + fontFileBytes.ToString(CultureInfo.InvariantCulture) +
            " before compression; " + originalBytes.ToString(CultureInfo.InvariantCulture) + " source bytes).",
            PdfConversionWarningSeverity.Information,
            details: new Dictionary<string, string> {
                ["fontName"] = fontName,
                ["format"] = format,
                ["fontFileObjectId"] = objectId.ToString(CultureInfo.InvariantCulture),
                ["originalFontBytes"] = originalBytes.ToString(CultureInfo.InvariantCulture),
                ["embeddedFontFileBytes"] = fontFileBytes.ToString(CultureInfo.InvariantCulture),
                ["embeddedStreamBytes"] = streamBytes.ToString(CultureInfo.InvariantCulture),
                ["compressed"] = compressed ? "true" : "false",
                ["embeddingMode"] = embeddingMode,
                ["byteScope"] = "font-program-stream-only"
            }));
    }
}
