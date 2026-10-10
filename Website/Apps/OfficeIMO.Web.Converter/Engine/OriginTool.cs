using System.Globalization;
using System.Text;
using OfficeIMO.Provenance;
using OfficeIMO.Web.Converter.Models;
using OfficeIMO.Web.Converter.Services;
using OfficeIMO.Workflows;

namespace OfficeIMO.Web.Converter.Engine;

/// <summary>Finds Content Credentials and AI source labels, and writes a copy without the records the visitor selects.</summary>
internal static class OriginTool {
    internal const string Manifests = "manifests";
    internal const string References = "references";
    internal const string Declarations = "declarations";

    internal static ToolResultDocument Run(ToolSession session, string action, ToolOptions options) => action switch {
        "inspect" => Inspect(session),
        "remove" => Remove(session, options),
        _ => throw new NotSupportedException($"The file origin tool has no '{action}' action.")
    };

    private static ToolResultDocument Inspect(ToolSession session) {
        SelectedDocument file = session.Input();
        EnsureSupported(file);
        OfficeProvenanceReport report = OfficeProvenanceBufferWorkflow.Inspect(file.Bytes, file.Name, BrowserProvenancePolicy.Limits());
        session.LastReview = report;
        BrowserConversionArtifact json = Report(file, report, removal: null);
        ToolArtifact[] artifacts = session.ReplaceArtifacts((json, "report"));

        var items = Categories(report).Select(category => new ToolItem(
            category.Id,
            category.Title,
            category.Detail,
            ToolState.Found,
            Selectable: category.Selectable,
            Selected: category.Selectable)).ToList();
        OfficeC2paManifestSummary[] manifests = ManifestSummaries(report);
        var manifestEvidence = report.Evidence.Where(static evidence => evidence.Manifest != null).Take(3).ToArray();
        for (int index = 0; index < manifests.Length; index++) items.AddRange(Explain(manifests[index], index, manifests.Length, manifestEvidence[index]));
        items.AddRange(report.Diagnostics.Select((diagnostic, index) => new ToolItem("d" + index.ToString(CultureInfo.InvariantCulture), "Note", diagnostic, ToolState.Info)));

        ToolVerdict verdict = report.Evidence.Count == 0
            ? new ToolVerdict(ToolTone.Good, "No origin records found",
                "This file has no Content Credentials or AI source labels that the tool can read. That alone does not prove where it came from.")
            : new ToolVerdict(ToolTone.Warn,
                "Found " + ToolFormat.Count(report.Evidence.Count, "origin record"),
                Summary(report, manifests));
        var facts = new List<ToolFact> { new("Format", FormatName(report.Format)) };
        if (manifests.Length > 0) {
            OfficeC2paManifestSummary first = manifests[0];
            string? model = first.Actions.Select(static action => Clip(action.SoftwareAgent)).FirstOrDefault(static agent => agent != null);
            if (EmbeddedAsset(manifestEvidence[0]) is { } asset) facts.Add(new ToolFact("Credential asset", asset));
            if ((model ?? Clip(first.ClaimGenerator)) is { } madeWith) facts.Add(new ToolFact(manifests.Length > 1 ? "First recorded tool" : "Recorded tool", madeWith));
            if (Clip(first.SignedBy) is { } signer) facts.Add(new ToolFact(manifests.Length > 1 ? "First certificate subject" : "Certificate subject", signer + " · unverified", ToolTone.Warn));
        }
        if (report.HasGenerativeAiDeclaration) facts.Add(new ToolFact("Generative AI", "Declared in an origin record", ToolTone.Warn));
        return new ToolResultDocument(true, verdict, facts.ToArray(), items.ToArray(), artifacts, null, 0);
    }

    private static OfficeC2paManifestSummary[] ManifestSummaries(OfficeProvenanceReport report) =>
        report.Evidence.Select(static evidence => evidence.Manifest).OfType<OfficeC2paManifestSummary>().Take(3).ToArray();

    /// <summary>Spells out what one Content Credentials record says, as written in the file.</summary>
    private static IEnumerable<ToolItem> Explain(OfficeC2paManifestSummary manifest, int index, int count, OfficeProvenanceEvidence evidence) {
        string group = count > 1 ? $"What Content Credentials record {index + 1} says" : "What the Content Credentials say";
        if (EmbeddedAsset(evidence) is { } asset) group += $" · embedded asset {asset}";
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
                yield return new ToolItem(prefix + "more", ToolFormat.Count(manifest.Actions.Count - 8, "more step"), "Listed in the JSON report.", ToolState.Detail, group);
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
    private static bool Watermarked(string action) =>
        action == "c2pa.watermarked" || action.StartsWith("c2pa.watermarked.", StringComparison.Ordinal);

    private const string WatermarkNote = "The record says a watermark was also added to the content itself. Removing the record does not remove that watermark, and detection tools may still recognise the file.";

    private static string Capitalize(string text) => text.Length == 0 ? text : char.ToUpperInvariant(text[0]) + text.Substring(1);

    private static string SourceText(OfficeProvenanceDigitalSourceKind kind, string? raw) => kind switch {
        OfficeProvenanceDigitalSourceKind.TrainedAlgorithmicMedia => "Made by a generative AI model.",
        OfficeProvenanceDigitalSourceKind.CompositeWithTrainedAlgorithmicMedia => "Mixes AI-generated and other content.",
        OfficeProvenanceDigitalSourceKind.AlgorithmicMedia => "Made by software, not captured by a camera.",
        OfficeProvenanceDigitalSourceKind.DigitalCapture => "Captured by a camera or scanner.",
        OfficeProvenanceDigitalSourceKind.CompositeCapture => "Assembled from captured media.",
        _ => Clip(raw) is { } value ? $"Source type: {value}." : string.Empty
    };

    private static ToolResultDocument Remove(ToolSession session, ToolOptions options) {
        SelectedDocument file = session.Input();
        EnsureSupported(file);
        IReadOnlySet<string> selected = options.List("remove");
        if (selected.Count == 0) throw new InvalidOperationException("Select at least one kind of record to remove.");
        OfficeProvenanceRemovalResult removal = OfficeProvenanceBufferWorkflow.Remove(
            file.Bytes,
            file.Name,
            BrowserProvenancePolicy.Removal(selected.Contains(Manifests), selected.Contains(References), selected.Contains(Declarations)));
        byte[] output = removal.ToArray();
        string outputName = Path.GetFileNameWithoutExtension(file.Name) + "-cleaned" + file.Extension;
        ToolArtifact[] artifacts = session.ReplaceArtifacts(
            (new BrowserConversionArtifact(output, outputName, ContentType(file.Extension)), "primary"),
            (Report(file, removal.After, removal), "report"));

        var remaining = Categories(removal.After).ToDictionary(static category => category.Id, StringComparer.Ordinal);
        var items = new List<ToolItem>();
        int removed = 0, keptByChoice = 0, stuck = 0;
        foreach (Category category in Categories(removal.Before)) {
            if (!category.Selectable || !selected.Contains(category.Id)) {
                keptByChoice += category.Count;
                items.Add(new ToolItem(category.Id, category.Title, category.Selectable
                    ? "Kept because you chose to keep it."
                    : "Kept: these source labels do not declare generative AI. " + category.Detail, ToolState.Kept));
            } else if (remaining.TryGetValue(category.Id, out Category? left)) {
                stuck += left.Count;
                removed += category.Count - left.Count;
                items.Add(new ToolItem(category.Id, category.Title,
                    "Could not be removed: the record is signed, malformed, or in a place this tool won't rewrite. " + left.Detail,
                    ToolState.Bad));
            } else {
                removed += category.Count;
                items.Add(new ToolItem(category.Id, category.Title, category.Detail, ToolState.Removed));
            }
        }
        items.Add(new ToolItem("content", file.Extension is ".png" or ".jpg" or ".jpeg" or ".webp" ? "Image pixels" : "Document content",
            "Not changed. Only the hidden records were touched.", ToolState.Kept));
        bool watermark = selected.Contains(Manifests) && removal.Before.Evidence.Select(static evidence => evidence.Manifest)
            .OfType<OfficeC2paManifestSummary>().Any(static manifest => manifest.Actions.Any(static action => Watermarked(action.Action)));
        if (watermark) {
            items.Add(new ToolItem("watermark", "Watermark in the content", WatermarkNote, ToolState.Warning));
        }
        items.AddRange(removal.After.Diagnostics.Select((diagnostic, index) => new ToolItem("d" + index.ToString(CultureInfo.InvariantCulture), "Note", diagnostic, ToolState.Info)));

        string title = removed == 0
            ? "Nothing was removed"
            : "Clean copy ready · " + ToolFormat.Count(removed, "record") + " removed";
        string detail = stuck > 0
            ? ToolFormat.Count(stuck, "record") + " could not be removed. See the list below."
            : keptByChoice > 0
                ? "We checked the copy again. The records shown as kept remain."
                : "We checked the copy again and found no origin records.";
        if (watermark) detail += " The watermark the record mentions is part of the content and is still there.";
        var verdict = new ToolVerdict(stuck > 0 || removed == 0 ? ToolTone.Warn : ToolTone.Good, title, detail);

        ToolPreview? preview = file.Extension is ".png" or ".jpg" or ".jpeg" or ".webp"
            ? new ToolPreview("image", 0)
            : file.Extension == ".pdf" ? new ToolPreview("pdf", 0) : null;
        return new ToolResultDocument(true, verdict,
            [new ToolFact("Removed", removed.ToString(CultureInfo.InvariantCulture), removed > 0 ? ToolTone.Good : null),
             new ToolFact("Still present", removal.After.Evidence.Count.ToString(CultureInfo.InvariantCulture), removal.After.Evidence.Count > 0 ? ToolTone.Warn : null),
             new ToolFact("Size", ToolFormat.Bytes(output.LongLength))],
            items.ToArray(), artifacts, preview, 0);
    }

    private sealed record Category(string Id, string Title, string Detail, int Count, bool Selectable);

    private static IEnumerable<Category> Categories(OfficeProvenanceReport report) {
        foreach (var group in report.Evidence.GroupBy(static evidence => (evidence.Carrier,
            Ai: evidence.DigitalSourceKind is OfficeProvenanceDigitalSourceKind.TrainedAlgorithmicMedia or
                OfficeProvenanceDigitalSourceKind.CompositeWithTrainedAlgorithmicMedia)).OrderBy(static group => group.Key.Carrier)) {
            OfficeProvenanceEvidence[] evidence = group.ToArray();
            string[] places = evidence.Select(static item => Where(item)).Distinct(StringComparer.Ordinal).ToArray();
            string where = string.Join("; ", places.Take(3)) + (places.Length > 3 ? $"; and {places.Length - 3} more places" : string.Empty);
            string validity = evidence.All(static item => item.IsStructurallyValid) ? string.Empty : " Some look malformed and will be left in place.";
            (string id, string title, string what) = group.Key.Carrier switch {
                OfficeProvenanceCarrierKind.C2paManifest => (Manifests, "Content Credentials",
                    "A record claiming how the file was made or edited (C2PA). Its signature is not verified here."),
                OfficeProvenanceCarrierKind.C2paExternalManifest => (References, "Link to Content Credentials",
                    "Points to a record stored outside the file" + (evidence[0].Value is { Length: > 0 } uri ? $" ({uri})" : string.Empty) + "."),
                _ when group.Key.Ai => (Declarations, "AI source label", SourceTypeText(evidence)),
                _ => ("source-labels", "Other source label", SourceTypeText(evidence))
            };
            yield return new Category(id, evidence.Length > 1 ? $"{title} ({evidence.Length})" : title, $"{what} Found in {where}.{validity}", evidence.Length, id != "source-labels");
        }
    }

    /// <summary>Describes where a record sits in plain words instead of a format-native path such as JPEG[0]/APP11@20.</summary>
    internal static string Where(OfficeProvenanceEvidence evidence) => EmbeddedAsset(evidence) is { } asset
        ? "embedded asset " + asset
        : evidence.Carrier switch {
        OfficeProvenanceCarrierKind.C2paManifest => "an embedded Content Credentials store",
        _ when evidence.Location.Contains("XMP", StringComparison.OrdinalIgnoreCase) => "the file's XMP metadata",
        _ => evidence.Location
    };

    private static string? EmbeddedAsset(OfficeProvenanceEvidence evidence) {
        // ZIP's package-level credential has no nested format path. Nested evidence retains its entry name.
        if (!evidence.Location.StartsWith("ZIP/", StringComparison.Ordinal)) return null;
        string path = evidence.Location.Substring(4);
        int end = -1;
        foreach (string boundary in new[] { "/PNG/", "/JPEG[", "/RIFF/", "/WebP/" }) {
            end = Math.Max(end, path.LastIndexOf(boundary, StringComparison.Ordinal));
        }
        return end < 0 ? null : path.Substring(0, end);
    }

    private static string SourceTypeText(OfficeProvenanceEvidence[] evidence) {
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

    private static string Summary(OfficeProvenanceReport report, OfficeC2paManifestSummary[] manifests) {
        var parts = new List<string>();
        OfficeC2paManifestSummary? first = manifests.FirstOrDefault();
        string? generator = first == null ? null : Generator(first);
        if (report.HasGenerativeAiDeclaration) {
            OfficeProvenanceEvidence? aiRecord = report.Evidence.FirstOrDefault(static evidence => evidence.Manifest?.DeclaresGenerativeAi == true);
            string? aiGenerator = aiRecord?.Manifest is { } aiManifest ? AiGenerator(aiManifest) : null;
            string subject = aiRecord != null && EmbeddedAsset(aiRecord) is { } aiAsset ? $"An embedded asset ({aiAsset})" : "An origin record";
            parts.Add(aiGenerator != null ? $"{subject} says it was made by a generative AI tool: {aiGenerator}." : "An origin record declares generative AI.");
        } else if (generator != null) {
            OfficeProvenanceEvidence firstRecord = report.Evidence.First(static evidence => evidence.Manifest != null);
            string subject = EmbeddedAsset(firstRecord) is { } asset ? $"An embedded asset ({asset})" : "A Content Credentials record";
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

    private static BrowserConversionArtifact Report(SelectedDocument file, OfficeProvenanceReport report, OfficeProvenanceRemovalResult? removal) =>
        new(Encoding.UTF8.GetBytes(OfficeProvenanceReportSerializer.Serialize(OfficeProvenanceReportSerializer.FromBuffer(file.Name, file.Bytes, report, removal))),
            Path.GetFileNameWithoutExtension(file.Name) + "-origin-report.json",
            "application/json");

    private static void EnsureSupported(SelectedDocument file) {
        if (!OfficeProvenanceWorkflowCatalog.BrowserExtensions.Contains(file.Extension, StringComparer.OrdinalIgnoreCase)) {
            throw new NotSupportedException("Choose a PNG, JPEG or WebP image, a PDF, or a Word, Excel or PowerPoint file.");
        }
    }

    private static string FormatName(OfficeProvenanceAssetFormat format) => format switch {
        OfficeProvenanceAssetFormat.ZipPackage => "Office document",
        OfficeProvenanceAssetFormat.Pdf => "PDF",
        _ => format.ToString().ToUpperInvariant()
    };

    private static string ContentType(string extension) => extension switch {
        ".jpg" or ".jpeg" => "image/jpeg",
        ".png" => "image/png",
        ".webp" => "image/webp",
        ".pdf" => "application/pdf",
        ".docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".pptx" => "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        _ => "application/octet-stream"
    };
}
