using AngleSharp.Html.Dom;
using OfficeIMO.Core.Internal;
using OfficeIMO.Provenance;

namespace OfficeIMO.Html;

/// <summary>Inspects and selectively removes standards-defined provenance from HTML documents.</summary>
public static class HtmlProvenance {
    /// <summary>Inspects embedded and external C2PA carriers plus supported embedded image data URIs.</summary>
    public static OfficeProvenanceReport Inspect(string html, OfficeProvenanceOptions? options = null) {
        if (html == null) throw new ArgumentNullException(nameof(html));
        options ??= new OfficeProvenanceOptions();
        OfficeProvenanceBinary.ValidateLimits(options);
        byte[] encoded = Encoding.UTF8.GetBytes(html);
        if (encoded.LongLength > options.MaxAssetBytes) throw new InvalidDataException("The HTML document exceeds the configured asset limit.");

        IHtmlDocument document = HtmlDocumentParser.ParseDocument(html);
        var evidence = new List<OfficeProvenanceEvidence>();
        var diagnostics = new List<string>();
        int carrierIndex = 0;
        foreach (IElement script in document.QuerySelectorAll("script[type]")) {
            if (!string.Equals(script.GetAttribute("type")?.Trim(), "application/c2pa", StringComparison.OrdinalIgnoreCase)) continue;
            TryDecodeManifest(script.TextContent, options.MaxManifestBytes, out byte[] manifest);
            bool valid = manifest.Length != 0 && OfficeC2paManifestStore.IsValid(manifest, 0, manifest.Length, options.MaxManifestBytes, out _);
            AddEvidence(evidence, options, new OfficeProvenanceEvidence(
                OfficeProvenanceCarrierKind.C2paManifest,
                $"HTML/script[type=application/c2pa][{carrierIndex++}]",
                valid,
                manifest.Length));
        }

        foreach (IElement link in document.QuerySelectorAll("link[rel][href]")) {
            if (!HasRelationship(link.GetAttribute("rel"), "c2pa-manifest")) continue;
            string value = link.GetAttribute("href")?.Trim() ?? string.Empty;
            bool valid = Uri.TryCreate(value, UriKind.Absolute, out Uri? uri) &&
                (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps);
            AddEvidence(evidence, options, new OfficeProvenanceEvidence(
                OfficeProvenanceCarrierKind.C2paExternalManifest,
                $"HTML/link[rel=c2pa-manifest][{carrierIndex++}]",
                valid,
                0,
                valid ? uri!.AbsoluteUri : value));
        }

        if (options.ProcessEmbeddedAssets) InspectEmbeddedImages(document, options, evidence, diagnostics);
        return new OfficeProvenanceReport(OfficeProvenanceAssetFormat.Html, evidence.AsReadOnly(), diagnostics.AsReadOnly());
    }

    /// <summary>Inspects a bounded HTML file without resolving external resources.</summary>
    public static OfficeProvenanceReport InspectFile(string filePath, OfficeProvenanceOptions? options = null) {
        if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("A file path is required.", nameof(filePath));
        options ??= new OfficeProvenanceOptions();
        OfficeProvenanceBinary.ValidateLimits(options);
        byte[] data = ReadBounded(filePath, options.MaxAssetBytes);
        return Inspect(Encoding.UTF8.GetString(data), options);
    }

    /// <summary>Removes selected HTML provenance and provenance in supported embedded image data URIs.</summary>
    public static OfficeProvenanceRemovalResult Remove(string html, OfficeProvenanceRemovalOptions? options = null) {
        if (html == null) throw new ArgumentNullException(nameof(html));
        options ??= new OfficeProvenanceRemovalOptions();
        OfficeProvenanceBinary.ValidateLimits(options.Limits);
        if (options.MaxEmbeddedAssets <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxEmbeddedAssets));
        OfficeProvenanceReport before = Inspect(html, options.Limits);
        IHtmlDocument document = HtmlDocumentParser.ParseDocument(html);
        var changes = new List<OfficeProvenanceChange>();

        int carrierIndex = 0;
        foreach (IElement script in document.QuerySelectorAll("script[type]").ToArray()) {
            if (!string.Equals(script.GetAttribute("type")?.Trim(), "application/c2pa", StringComparison.OrdinalIgnoreCase)) continue;
            TryDecodeManifest(script.TextContent, options.Limits.MaxManifestBytes, out byte[] manifest);
            bool valid = manifest.Length != 0 && OfficeC2paManifestStore.IsValid(manifest, 0, manifest.Length, options.Limits.MaxManifestBytes, out _);
            string location = $"HTML/script[type=application/c2pa][{carrierIndex++}]";
            if (!options.RemoveC2paManifests || (!valid && options.RequireStructurallyValidCarrier)) continue;
            script.Remove();
            changes.Add(new OfficeProvenanceChange(OfficeProvenanceCarrierKind.C2paManifest, location, manifest.Length));
        }

        foreach (IElement link in document.QuerySelectorAll("link[rel][href]").ToArray()) {
            if (!HasRelationship(link.GetAttribute("rel"), "c2pa-manifest")) continue;
            string value = link.GetAttribute("href")?.Trim() ?? string.Empty;
            bool valid = Uri.TryCreate(value, UriKind.Absolute, out Uri? uri) &&
                (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps);
            string location = $"HTML/link[rel=c2pa-manifest][{carrierIndex++}]";
            if (!options.RemoveExternalC2paReferences || (!valid && options.RequireStructurallyValidCarrier)) continue;
            link.Remove();
            changes.Add(new OfficeProvenanceChange(OfficeProvenanceCarrierKind.C2paExternalManifest, location, 0));
        }

        if (options.ProcessEmbeddedAssets) RemoveEmbeddedImages(document, options, changes);
        if (changes.Count == 0) {
            byte[] original = Encoding.UTF8.GetBytes(html);
            return new OfficeProvenanceRemovalResult(original, before, before, changes.AsReadOnly(), false);
        }

        string outputHtml = document.DocumentElement?.OuterHtml ?? string.Empty;
        byte[] output = Encoding.UTF8.GetBytes(outputHtml);
        if (output.LongLength > options.Limits.MaxAssetBytes) throw new InvalidDataException("The rewritten HTML document exceeds the configured asset limit.");
        OfficeProvenanceReport after = Inspect(outputHtml, options.Limits);
        return new OfficeProvenanceRemovalResult(output, before, after, changes.AsReadOnly(), true);
    }

    /// <summary>Removes selected provenance and atomically writes the resulting HTML file.</summary>
    public static OfficeProvenanceRemovalResult RemoveFile(
        string inputPath,
        string outputPath,
        OfficeProvenanceRemovalOptions? options = null) {
        if (string.IsNullOrWhiteSpace(inputPath)) throw new ArgumentException("An input path is required.", nameof(inputPath));
        if (string.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("An output path is required.", nameof(outputPath));
        options ??= new OfficeProvenanceRemovalOptions();
        byte[] input = ReadBounded(inputPath, options.Limits.MaxAssetBytes);
        OfficeProvenanceRemovalResult result = Remove(Encoding.UTF8.GetString(input), options);
        OfficeFileCommit.WriteAllBytes(Path.GetFullPath(outputPath), result.ToArray());
        return result;
    }

    private static void InspectEmbeddedImages(
        IHtmlDocument document,
        OfficeProvenanceOptions options,
        List<OfficeProvenanceEvidence> evidence,
        List<string> diagnostics) {
        int count = 0;
        foreach (IElement element in document.QuerySelectorAll("img[src],source[src]")) {
            if (!HtmlImageDataUri.TryParse(element.GetAttribute("src"), out HtmlImageDataUri dataUri)) continue;
            count++;
            if (count > options.MaxEmbeddedAssets) throw new InvalidDataException("The HTML document exceeds the configured embedded-asset limit.");
            if (dataUri.EstimateDecodedByteCount() > options.MaxAssetBytes) throw new InvalidDataException("An embedded HTML image exceeds the configured asset limit.");
            if (!dataUri.TryDecodeBytes(out byte[] image)) {
                diagnostics.Add($"HTML/{element.LocalName}[src][{count - 1}]: embedded image data URI could not be decoded.");
                continue;
            }
            try {
                OfficeProvenanceReport nested = OfficeProvenanceInspector.Inspect(image, "asset" + dataUri.FileExtension, CreateNestedOptions(options));
                foreach (OfficeProvenanceEvidence item in nested.Evidence) {
                    AddEvidence(evidence, options, Prefix($"HTML/{element.LocalName}[src][{count - 1}]", item));
                }
                foreach (string diagnostic in nested.Diagnostics) diagnostics.Add($"HTML/{element.LocalName}[src][{count - 1}]: {diagnostic}");
            } catch (InvalidDataException exception) {
                diagnostics.Add($"HTML/{element.LocalName}[src][{count - 1}]: embedded image was preserved because inspection failed: {exception.Message}");
            }
        }
    }

    private static void RemoveEmbeddedImages(
        IHtmlDocument document,
        OfficeProvenanceRemovalOptions options,
        List<OfficeProvenanceChange> changes) {
        int count = 0;
        foreach (IElement element in document.QuerySelectorAll("img[src],source[src]")) {
            if (!HtmlImageDataUri.TryParse(element.GetAttribute("src"), out HtmlImageDataUri dataUri)) continue;
            count++;
            if (count > options.MaxEmbeddedAssets) throw new InvalidDataException("The HTML document exceeds the configured embedded-asset limit.");
            if (dataUri.EstimateDecodedByteCount() > options.Limits.MaxAssetBytes) throw new InvalidDataException("An embedded HTML image exceeds the configured asset limit.");
            if (!dataUri.TryDecodeBytes(out byte[] image)) continue;
            try {
                OfficeProvenanceRemovalResult nested = OfficeProvenanceRemover.Remove(
                    image,
                    "asset" + dataUri.FileExtension,
                    CreateNestedRemovalOptions(options));
                if (!nested.WasChanged) continue;
                element.SetAttribute("src", "data:" + dataUri.MediaType + ";base64," + Convert.ToBase64String(nested.ToArray()));
                foreach (OfficeProvenanceChange change in nested.Changes) {
                    changes.Add(new OfficeProvenanceChange(
                        change.Carrier,
                        $"HTML/{element.LocalName}[src][{count - 1}]/{change.Location}",
                        change.RemovedBytes));
                }
            } catch (InvalidDataException) {
                // Preserve malformed embedded data; structural diagnostics are available through Inspect.
            }
        }
    }

    private static bool TryDecodeManifest(string? value, long maximumBytes, out byte[] manifest) {
        manifest = Array.Empty<byte>();
        string encoded = (value ?? string.Empty).Trim();
        const string prefix = "data:application/c2pa;base64,";
        if (encoded.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) encoded = encoded.Substring(prefix.Length);
        if (encoded.Length == 0 || encoded.Length > maximumBytes * 2L || encoded.Length > int.MaxValue) return false;
        try {
            manifest = Convert.FromBase64String(encoded);
            return manifest.LongLength <= maximumBytes;
        } catch (FormatException) {
            manifest = Array.Empty<byte>();
            return false;
        }
    }

    private static bool HasRelationship(string? value, string relationship) =>
        (value ?? string.Empty).Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)
            .Any(item => item.Equals(relationship, StringComparison.OrdinalIgnoreCase));

    private static void AddEvidence(List<OfficeProvenanceEvidence> evidence, OfficeProvenanceOptions options, OfficeProvenanceEvidence item) {
        if (evidence.Count >= options.MaxCarriers) throw new InvalidDataException($"The asset exceeds the configured carrier limit of {options.MaxCarriers}.");
        evidence.Add(item);
    }

    private static OfficeProvenanceEvidence Prefix(string prefix, OfficeProvenanceEvidence item) =>
        new OfficeProvenanceEvidence(item.Carrier, prefix + "/" + item.Location, item.IsStructurallyValid, item.PayloadLength, item.Value, item.DigitalSourceKind);

    private static OfficeProvenanceOptions CreateNestedOptions(OfficeProvenanceOptions source) => new OfficeProvenanceOptions {
        MaxAssetBytes = source.MaxAssetBytes,
        MaxManifestBytes = source.MaxManifestBytes,
        MaxCarriers = source.MaxCarriers,
        MaxContainerEntries = source.MaxContainerEntries,
        MaxExpandedContainerBytes = source.MaxExpandedContainerBytes,
        ProcessEmbeddedAssets = false,
        MaxEmbeddedAssets = source.MaxEmbeddedAssets
    };

    private static OfficeProvenanceRemovalOptions CreateNestedRemovalOptions(OfficeProvenanceRemovalOptions source) {
        var nested = new OfficeProvenanceRemovalOptions {
            RemoveC2paManifests = source.RemoveC2paManifests,
            RemoveExternalC2paReferences = source.RemoveExternalC2paReferences,
            RemoveAiSourceMetadata = source.RemoveAiSourceMetadata,
            RequireStructurallyValidCarrier = source.RequireStructurallyValidCarrier,
            ProcessEmbeddedAssets = false,
            MaxEmbeddedAssets = source.MaxEmbeddedAssets
        };
        nested.Limits.MaxAssetBytes = source.Limits.MaxAssetBytes;
        nested.Limits.MaxManifestBytes = source.Limits.MaxManifestBytes;
        nested.Limits.MaxCarriers = source.Limits.MaxCarriers;
        nested.Limits.MaxContainerEntries = source.Limits.MaxContainerEntries;
        nested.Limits.MaxExpandedContainerBytes = source.Limits.MaxExpandedContainerBytes;
        nested.Limits.ProcessEmbeddedAssets = false;
        nested.Limits.MaxEmbeddedAssets = source.Limits.MaxEmbeddedAssets;
        return nested;
    }

    private static byte[] ReadBounded(string filePath, long maximumBytes) {
        string fullPath = Path.GetFullPath(filePath);
        using var stream = File.OpenRead(fullPath);
        return OfficeProvenanceBinary.ReadBounded(stream, maximumBytes);
    }
}
