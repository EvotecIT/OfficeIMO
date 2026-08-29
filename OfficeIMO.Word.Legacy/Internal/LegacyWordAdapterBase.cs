using System;
using System.IO;
using System.Linq;
using System.Threading;

namespace OfficeIMO.Word.Legacy;

internal abstract class LegacyWordAdapterBase : ILegacyWordAdapter {
    public abstract LegacyWordFormat Format { get; }
    public abstract string ProfileId { get; }
    public abstract int Probe(byte[] data, string? sourceName, out string reason);
    public abstract LegacyWordModel Parse(byte[] data, OfficeLegacyImportLimits limits, CancellationToken cancellationToken);

    protected static LegacyWordModel Salvage(byte[] data, OfficeLegacyImportLimits limits, int offset, bool stripHighBit, string profileId, string limitation, CancellationToken cancellationToken) {
        string text = OfficeLegacyImportBuffer.ExtractPrintableText(data, Math.Min(offset, data.Length), Math.Max(0, data.Length - Math.Min(offset, data.Length)), limits.MaxTextCharacters, stripHighBit, cancellationToken: cancellationToken);
        if (string.IsNullOrWhiteSpace(text)) throw new InvalidDataException($"{profileId} did not contain recoverable bounded text.");
        var model = new LegacyWordModel { Quality = OfficeLegacyImportQuality.Salvage };
        AddParagraphs(model, text, limits, cancellationToken);
        model.Findings.Add(Loss("LEGACY_WORD_SALVAGE", "Structure", limitation));
        return model;
    }

    protected static void AddParagraphs(LegacyWordModel model, string text, OfficeLegacyImportLimits limits, CancellationToken cancellationToken) {
        foreach (string raw in text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n')) {
            cancellationToken.ThrowIfCancellationRequested();
            if (model.Paragraphs.Count >= limits.MaxItems) throw new InvalidDataException("Legacy word content exceeds the configured item limit.");
            string line = raw.TrimEnd();
            if (line.Length == 0 && (model.Paragraphs.Count == 0 || model.Paragraphs[model.Paragraphs.Count - 1].Text.Length == 0)) continue;
            bool list = line.StartsWith("- ", StringComparison.Ordinal) || line.StartsWith("* ", StringComparison.Ordinal);
            model.Paragraphs.Add(new LegacyWordParagraph(list ? line.Substring(2) : line, list));
        }
        while (model.Paragraphs.Count > 0 && model.Paragraphs[model.Paragraphs.Count - 1].Text.Length == 0) {
            model.Paragraphs.RemoveAt(model.Paragraphs.Count - 1);
        }
    }

    protected static bool ExtensionIs(string? sourceName, params string[] extensions) {
        string extension = Path.GetExtension(sourceName ?? string.Empty);
        return extensions.Any(candidate => string.Equals(extension, candidate, StringComparison.OrdinalIgnoreCase));
    }

    protected static OfficeCompatibilityFinding Loss(string code, string category, string message) =>
        new(code, category, message, OfficeCompatibilityState.Approximated, OfficeCompatibilitySeverity.Warning,
            OfficeCompatibilityImpact.Semantic | OfficeCompatibilityImpact.Visual | OfficeCompatibilityImpact.Carrier, true);

    protected static OfficeCompatibilityFinding Inert(string code, string category, string message) =>
        new(code, category, message, OfficeCompatibilityState.Dropped, OfficeCompatibilitySeverity.Warning,
            OfficeCompatibilityImpact.Behavioral | OfficeCompatibilityImpact.Security | OfficeCompatibilityImpact.Carrier, true);
}
