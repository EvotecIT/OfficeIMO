using AngleSharp;
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
        IElement? head = document.Head;
        if (head == null) return new OfficeProvenanceReport(OfficeProvenanceAssetFormat.Html, evidence.AsReadOnly(), diagnostics.AsReadOnly());
        IElement[] manifestElements = head.QuerySelectorAll("script[type],link[rel][href]")
            .Where(IsManifestElement)
            .ToArray();
        if (manifestElements.Length > 1) diagnostics.Add("manifest.html.multipleManifests: the HTML head contains multiple C2PA manifest associations.");
        int carrierIndex = 0;
        foreach (IElement script in head.QuerySelectorAll("script[type]")) {
            if (!string.Equals(script.GetAttribute("type")?.Trim(), "application/c2pa", StringComparison.OrdinalIgnoreCase)) continue;
            TryDecodeManifest(script.TextContent, options.MaxManifestBytes, out byte[] manifest);
            bool valid = manifestElements.Length == 1 && manifest.Length != 0 &&
                OfficeC2paManifestStore.IsValid(manifest, 0, manifest.Length, options.MaxManifestBytes, out _);
            AddEvidence(evidence, options, new OfficeProvenanceEvidence(
                OfficeProvenanceCarrierKind.C2paManifest,
                $"HTML/script[type=application/c2pa][{carrierIndex++}]",
                valid,
                manifest.Length));
        }

        foreach (IElement link in head.QuerySelectorAll("link[rel][href]")) {
            if (!HasRelationship(link.GetAttribute("rel"), "c2pa-manifest")) continue;
            string value = link.GetAttribute("href")?.Trim() ?? string.Empty;
            bool safeReference = IsSafeManifestReference(value, out Uri? uri);
            bool valid = manifestElements.Length == 1 && safeReference;
            AddEvidence(evidence, options, new OfficeProvenanceEvidence(
                OfficeProvenanceCarrierKind.C2paExternalManifest,
                $"HTML/link[rel=c2pa-manifest][{carrierIndex++}]",
                valid,
                0,
                valid && uri!.IsAbsoluteUri ? uri.AbsoluteUri : value));
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
        return Inspect(DecodeHtml(data, out _, out _), options);
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
        IElement? head = document.Head;
        IEnumerable<IElement> scripts = head == null
            ? Enumerable.Empty<IElement>()
            : head.QuerySelectorAll("script[type]");
        IEnumerable<IElement> links = head == null
            ? Enumerable.Empty<IElement>()
            : head.QuerySelectorAll("link[rel][href]");
        int manifestElementCount = head?.QuerySelectorAll("script[type],link[rel][href]").Count(IsManifestElement) ?? 0;

        int carrierIndex = 0;
        foreach (IElement script in scripts.ToArray()) {
            if (!string.Equals(script.GetAttribute("type")?.Trim(), "application/c2pa", StringComparison.OrdinalIgnoreCase)) continue;
            TryDecodeManifest(script.TextContent, options.Limits.MaxManifestBytes, out byte[] manifest);
            bool valid = manifestElementCount == 1 && manifest.Length != 0 &&
                OfficeC2paManifestStore.IsValid(manifest, 0, manifest.Length, options.Limits.MaxManifestBytes, out _);
            string location = $"HTML/script[type=application/c2pa][{carrierIndex++}]";
            if (!options.RemoveC2paManifests || (!valid && options.RequireStructurallyValidCarrier)) continue;
            script.Remove();
            changes.Add(new OfficeProvenanceChange(OfficeProvenanceCarrierKind.C2paManifest, location, manifest.Length));
        }

        foreach (IElement link in links.ToArray()) {
            if (!HasRelationship(link.GetAttribute("rel"), "c2pa-manifest")) continue;
            string value = link.GetAttribute("href")?.Trim() ?? string.Empty;
            bool valid = manifestElementCount == 1 && IsSafeManifestReference(value, out _);
            string location = $"HTML/link[rel=c2pa-manifest][{carrierIndex++}]";
            if (!options.RemoveExternalC2paReferences || (!valid && options.RequireStructurallyValidCarrier)) continue;
            RemoveRelationship(link, "c2pa-manifest");
            changes.Add(new OfficeProvenanceChange(OfficeProvenanceCarrierKind.C2paExternalManifest, location, 0));
        }

        if (options.ProcessEmbeddedAssets) RemoveEmbeddedImages(document, options, changes);
        if (changes.Count == 0) {
            byte[] original = Encoding.UTF8.GetBytes(html);
            return new OfficeProvenanceRemovalResult(original, before, before, changes.AsReadOnly(), false);
        }

        string outputHtml = document.ToHtml();
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
        string html = DecodeHtml(input, out Encoding encoding, out bool hadPreamble);
        OfficeProvenanceRemovalResult result = Remove(html, options);
        if (!result.WasChanged) {
            OfficeFileCommit.WriteAllBytes(Path.GetFullPath(outputPath), input);
            return new OfficeProvenanceRemovalResult(
                (byte[])input.Clone(), result.Before, result.After, result.Changes, result.WasReserialized);
        }

        string outputHtml = Encoding.UTF8.GetString(result.ToArray());
        byte[] output = EncodeHtml(outputHtml, encoding, hadPreamble);
        if (output.LongLength > options.Limits.MaxAssetBytes) throw new InvalidDataException("The rewritten HTML document exceeds the configured asset limit.");
        OfficeFileCommit.WriteAllBytes(Path.GetFullPath(outputPath), output);
        return new OfficeProvenanceRemovalResult(
            output, result.Before, result.After, result.Changes, result.WasReserialized);
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

    private static void RemoveRelationship(IElement element, string relationship) {
        string[] retained = (element.GetAttribute("rel") ?? string.Empty)
            .Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)
            .Where(item => !item.Equals(relationship, StringComparison.OrdinalIgnoreCase))
            .ToArray();
        if (retained.Length == 0) element.Remove();
        else element.SetAttribute("rel", string.Join(" ", retained));
    }

    private static bool IsManifestElement(IElement element) =>
        element.LocalName.Equals("script", StringComparison.OrdinalIgnoreCase)
            ? string.Equals(element.GetAttribute("type")?.Trim(), "application/c2pa", StringComparison.OrdinalIgnoreCase)
            : element.LocalName.Equals("link", StringComparison.OrdinalIgnoreCase) &&
                HasRelationship(element.GetAttribute("rel"), "c2pa-manifest");

    private static bool IsSafeManifestReference(string value, out Uri? uri) {
        uri = null;
        if (string.IsNullOrWhiteSpace(value) || !Uri.TryCreate(value, UriKind.RelativeOrAbsolute, out Uri? parsed)) return false;
        if (parsed.IsAbsoluteUri && parsed.Scheme != Uri.UriSchemeHttp && parsed.Scheme != Uri.UriSchemeHttps) return false;
        uri = parsed;
        return true;
    }

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

    private static string DecodeHtml(byte[] data, out Encoding encoding, out bool hadPreamble) {
        using var stream = new MemoryStream(data, writable: false);
        encoding = HtmlTextEncodingResolver.ResolveHtmlEncoding(stream);
        byte[] preamble = encoding.GetPreamble();
        hadPreamble = preamble.Length != 0 && data.Length >= preamble.Length &&
            preamble.SequenceEqual(data.Take(preamble.Length));
        int offset = hadPreamble ? preamble.Length : 0;
        return encoding.GetString(data, offset, data.Length - offset);
    }

    private static byte[] EncodeHtml(string html, Encoding encoding, bool includePreamble) {
        byte[] body = encoding.GetBytes(html);
        byte[] preamble = includePreamble ? encoding.GetPreamble() : Array.Empty<byte>();
        if (preamble.Length == 0) return body;
        byte[] output = new byte[preamble.Length + body.Length];
        Buffer.BlockCopy(preamble, 0, output, 0, preamble.Length);
        Buffer.BlockCopy(body, 0, output, preamble.Length, body.Length);
        return output;
    }
}
