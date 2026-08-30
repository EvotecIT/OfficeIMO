using System;
using System.IO;
using System.Linq;
using System.Threading;

namespace OfficeIMO.Word.Legacy;

internal abstract class LegacyWordAdapterBase : ILegacyWordAdapter {
    public abstract LegacyWordFormat Format { get; }
    public abstract string ProfileId { get; }
    public virtual string GetProfileId(byte[] data, OfficeLegacyImportLimits limits, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        return ProfileId;
    }
    public abstract int Probe(byte[] data, string? sourceName, OfficeLegacyImportLimits limits, CancellationToken cancellationToken, out string reason);
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
        int inspectedRecords = 0;
        string normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
        for (int lineStart = 0; lineStart <= normalized.Length;) {
            cancellationToken.ThrowIfCancellationRequested();
            if (++inspectedRecords > limits.MaxRecords) throw new InvalidDataException("Legacy word content exceeds the configured record limit.");
            if (model.Paragraphs.Count >= limits.MaxItems) throw new InvalidDataException("Legacy word content exceeds the configured item limit.");
            int lineEnd = normalized.IndexOf('\n', lineStart);
            if (lineEnd < 0) lineEnd = normalized.Length;
            int trimmedEnd = lineEnd;
            while (trimmedEnd > lineStart && char.IsWhiteSpace(normalized[trimmedEnd - 1])) trimmedEnd--;
            int lineLength = trimmedEnd - lineStart;
            if (lineLength > 0 || (model.Paragraphs.Count > 0 && model.Paragraphs[model.Paragraphs.Count - 1].Text.Length > 0)) {
                bool list = lineLength >= 2 &&
                    (normalized[lineStart] == '-' || normalized[lineStart] == '*') &&
                    normalized[lineStart + 1] == ' ';
                int contentStart = list ? lineStart + 2 : lineStart;
                string line = normalized.Substring(contentStart, trimmedEnd - contentStart);
                model.Paragraphs.Add(new LegacyWordParagraph(line, list));
            }
            if (lineEnd == normalized.Length) break;
            lineStart = lineEnd + 1;
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

    internal static OfficeCompatibilityFinding LossFinding(string code, string category, string message) => Loss(code, category, message);

    internal static OfficeCompatibilityFinding InertFinding(string code, string category, string message) => Inert(code, category, message);
}
