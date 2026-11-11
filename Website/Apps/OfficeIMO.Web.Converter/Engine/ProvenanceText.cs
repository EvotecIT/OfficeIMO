using System.Globalization;
using OfficeIMO.Provenance;
using OfficeIMO.Web.Converter.Models;

namespace OfficeIMO.Web.Converter.Engine;

/// <summary>
/// Plain-language text for Content Credentials and AI source labels in <see cref="ShareTool"/>'s report: what each record claims,
/// where it sits and what removing it leaves behind.
/// </summary>
internal static class ProvenanceText {
    internal const string Group = "Content Credentials and AI labels";
    internal const string Manifests = "manifests";
    internal const string References = "references";
    internal const string Declarations = "declarations";

    internal static OfficeC2paManifestSummary[] ManifestSummaries(OfficeProvenanceReport report) =>
        report.Evidence.Select(static evidence => evidence.Manifest).OfType<OfficeC2paManifestSummary>().Take(3).ToArray();

    /// <summary>Spells out what one Content Credentials record says, as written in the file.</summary>
    internal static IEnumerable<ToolItem> Explain(OfficeC2paManifestSummary manifest, int index, int count, OfficeProvenanceEvidence evidence,
        Func<OfficeProvenanceEvidence, string?> assetOf) {
        string group = count > 1 ? $"What Content Credentials record {index + 1} says" : "What the Content Credentials say";
        if (assetOf(evidence) is { } asset) group += $" · picture {asset}";
        string prefix = "m" + index.ToString(CultureInfo.InvariantCulture) + "-";
        if (manifest.DeclaresGenerativeAi) yield return new ToolItem(prefix + "ai", "Generative AI declared",
            "This credential declares AI-generated content, including actions beyond the displayed timeline. The declaration is unverified.", ToolState.Detail, group);
        // The first record's app is already a fact at the top of the result.
        if (index > 0 && Generator(manifest) is { } generator) {
            yield return new ToolItem(prefix + "app", "Made with " + generator, "The app that wrote this record.", ToolState.Detail, group);
        }
        int shown = 0;
        foreach (OfficeC2paAction action in manifest.Actions) {
            if (shown++ == 8) {
                yield return new ToolItem(prefix + "more", ToolFormat.Count(manifest.Actions.Count - 8, "more step"), "The record lists more steps than are shown here.", ToolState.Detail, group);
                break;
            }
            string title = ActionName(action.Action) + (Clip(action.SoftwareAgent) is { } agent ? " · by " + agent : string.Empty);
            string detail = (Watermarked(action.Action) ? "Usually invisible. It stays in the file even if this record is removed. " : string.Empty)
                + SourceText(action.DigitalSourceKind, action.DigitalSourceType)
                + (Clip(action.When) is { } when ? $" Recorded {when}." : string.Empty);
            yield return new ToolItem(prefix + "a" + shown.ToString(CultureInfo.InvariantCulture), title, detail.Trim(), ToolState.Detail, group);
        }
        int ingredientIndex = 0;
        foreach (string ingredient in manifest.Ingredients.Take(4)) {
            if (Clip(ingredient) is { } name) {
                yield return new ToolItem(prefix + "i" + (ingredientIndex++).ToString(CultureInfo.InvariantCulture), "Based on " + name,
                    "Another file this one was made from (an ingredient).", ToolState.Detail, group);
            }
        }
        string signature = Clip(manifest.SignedBy) is { } signer
            ? "Certificate subject: " + signer + " (unverified)"
            : "Certificate subject not named";
        string issuer = Clip(manifest.CertificateIssuer) is { } ca && !string.Equals(ca, manifest.SignedBy, StringComparison.Ordinal)
            ? $"The certificate was issued by {ca}. "
            : string.Empty;
        yield return new ToolItem(prefix + "signer", signature,
            issuer + "This shows what the record claims. The tool reads the certificate's name but does not check the signature, so it cannot tell whether the record was changed afterwards.",
            ToolState.Detail, group);
    }

    private static string? Generator(OfficeC2paManifestSummary manifest) {
        string? generator = Clip(manifest.ClaimGenerator);
        string? model = manifest.Actions.Select(static action => Clip(action.SoftwareAgent)).FirstOrDefault(static agent => agent != null);
        if (generator == null) return model;
        return model != null && !generator.Contains(model, StringComparison.OrdinalIgnoreCase) ? $"{generator} ({model})" : generator;
    }

    private static string? AiGenerator(OfficeC2paManifestSummary manifest) {
        string? generator = Clip(manifest.ClaimGenerator);
        string? model = manifest.Actions.Where(static action => action.DigitalSourceKind is
                OfficeProvenanceDigitalSourceKind.TrainedAlgorithmicMedia or OfficeProvenanceDigitalSourceKind.CompositeWithTrainedAlgorithmicMedia)
            .Select(static action => Clip(action.SoftwareAgent)).FirstOrDefault(static agent => agent != null);
        return generator == null ? model : model != null && !generator.Contains(model, StringComparison.OrdinalIgnoreCase)
            ? $"{generator} ({model})" : generator;
    }

    private static string? Clip(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return null;
        string text = new string(value.Trim().Where(static ch => !char.IsControl(ch)).ToArray());
        return text.Length <= 80 ? text : text.Substring(0, 79) + "…";
    }

    private static string ActionName(string action) => action switch {
        "c2pa.created" => "Created",
        "c2pa.opened" => "Opened an existing file",
        "c2pa.placed" => "Placed other content into it",
        "c2pa.edited" => "Edited",
        "c2pa.cropped" => "Cropped",
        "c2pa.resized" => "Resized",
        "c2pa.filtered" => "Applied a filter",
        "c2pa.color_adjustments" => "Adjusted colours",
        "c2pa.drawing" => "Drew on it",
        "c2pa.orientation" => "Rotated or flipped",
        "c2pa.converted" or "c2pa.transcoded" => "Converted to another format",
        "c2pa.repackaged" => "Repackaged",
        "c2pa.published" => "Published",
        "c2pa.removed" => "Removed part of it",
        "c2pa.redacted" => "Redacted part of it",
        "c2pa.translated" => "Translated",
        _ when Watermarked(action) => "Added a watermark to the content itself",
        _ when action.StartsWith("c2pa.", StringComparison.Ordinal) => Capitalize(action.Substring(5).Replace('_', ' ').Replace('.', ' ')),
        _ => Clip(action) ?? "Unnamed step"
    };

    /// <summary>c2pa.watermarked, c2pa.watermarked.bound and c2pa.watermarked.unbound: a mark woven into the pixels or text, not into the metadata.</summary>
    internal static bool Watermarked(string action) =>
        action == "c2pa.watermarked" || action.StartsWith("c2pa.watermarked.", StringComparison.Ordinal);

    internal const string WatermarkNote = "The record says a watermark was also added to the content itself. Removing the record does not remove that watermark, and detection tools may still recognise the file.";

    private static string Capitalize(string text) => text.Length == 0 ? text : char.ToUpperInvariant(text[0]) + text.Substring(1);

    private static string SourceText(OfficeProvenanceDigitalSourceKind kind, string? raw) => kind switch {
        OfficeProvenanceDigitalSourceKind.TrainedAlgorithmicMedia => "Made by a generative AI model.",
        OfficeProvenanceDigitalSourceKind.CompositeWithTrainedAlgorithmicMedia => "Mixes AI-generated and other content.",
        OfficeProvenanceDigitalSourceKind.AlgorithmicMedia => "Made by software, not captured by a camera.",
        OfficeProvenanceDigitalSourceKind.DigitalCapture => "Captured by a camera or scanner.",
        OfficeProvenanceDigitalSourceKind.CompositeCapture => "Assembled from captured media.",
        _ => Clip(raw) is { } value ? $"Source type: {value}." : string.Empty
    };

    internal static string SourceTypeText(OfficeProvenanceEvidence[] evidence) {
        if (evidence.Select(static item => item.DigitalSourceKind).Distinct().Skip(1).Any())
            return "Contains source labels with different recorded source types; see the JSON report for each value.";
        OfficeProvenanceDigitalSourceKind kind = evidence[0].DigitalSourceKind;
        return kind switch {
            OfficeProvenanceDigitalSourceKind.TrainedAlgorithmicMedia => "Says the file was made by a generative AI tool.",
            OfficeProvenanceDigitalSourceKind.CompositeWithTrainedAlgorithmicMedia => "Says the file combines AI-generated and other content.",
            OfficeProvenanceDigitalSourceKind.AlgorithmicMedia => "Says the file was made by software, not captured by a camera.",
            OfficeProvenanceDigitalSourceKind.DigitalCapture => "Says the file was captured by a camera or scanner.",
            OfficeProvenanceDigitalSourceKind.CompositeCapture => "Says the file was assembled from captured media.",
            _ => "An IPTC label describing where the content came from" + (evidence[0].Value is { Length: > 0 } value ? $" ({value})" : string.Empty) + "."
        };
    }

    /// <param name="assetOf">Names the picture a record belongs to, or null for the file itself.</param>
    internal static string Summary(OfficeProvenanceReport report, OfficeC2paManifestSummary[] manifests, Func<OfficeProvenanceEvidence, string?> assetOf) {
        var parts = new List<string>();
        OfficeC2paManifestSummary? first = manifests.FirstOrDefault();
        string? generator = first == null ? null : Generator(first);
        if (report.HasGenerativeAiDeclaration) {
            OfficeProvenanceEvidence? aiRecord = report.Evidence.FirstOrDefault(static evidence => evidence.Manifest?.DeclaresGenerativeAi == true);
            string? aiGenerator = aiRecord?.Manifest is { } aiManifest ? AiGenerator(aiManifest) : null;
            string subject = aiRecord != null && assetOf(aiRecord) is { } aiAsset ? $"A picture ({aiAsset})" : "An origin record";
            parts.Add(aiGenerator != null ? $"{subject} says it was made by a generative AI tool: {aiGenerator}." : "An origin record declares generative AI.");
        } else if (generator != null) {
            OfficeProvenanceEvidence firstRecord = report.Evidence.First(static evidence => evidence.Manifest != null);
            string subject = assetOf(firstRecord) is { } asset ? $"A picture ({asset})" : "A Content Credentials record";
            parts.Add($"{subject} says it was made with {generator}.");
        }
        if (report.HasC2paManifest) {
            parts.Add(Clip(first?.SignedBy) is { } signer
                ? $"The first Content Credentials record names {signer} as its certificate subject; this is unverified."
                : "It carries Content Credentials.");
        }
        if (report.HasExternalC2paManifest) parts.Add("It links to Content Credentials stored elsewhere.");
        if (parts.Count == 0) parts.Add("Review what was found before choosing what to remove.");
        return string.Join(" ", parts);
    }
}
