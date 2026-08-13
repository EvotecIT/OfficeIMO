using AngleSharp;
using AngleSharp.Html.Dom;
using OfficeIMO.Core.Internal;
using OfficeIMO.Provenance;

namespace OfficeIMO.Html;

/// <summary>Inspects and selectively removes standards-defined provenance from HTML documents.</summary>
public static class HtmlProvenance {
    private static readonly string[] EmbeddedImageSourceAttributes = {
        "src", "data-src", "data-original", "data-original-src", "data-lazy-src"
    };
    private static readonly string[] EmbeddedImageSourceSetAttributes = {
        "srcset", "data-srcset", "data-original-srcset", "data-lazy-srcset"
    };

    /// <summary>Inspects embedded and external C2PA carriers plus supported embedded image data URIs.</summary>
    public static OfficeProvenanceReport Inspect(string html, OfficeProvenanceOptions? options = null) {
        if (html == null) throw new ArgumentNullException(nameof(html));
        options ??= new OfficeProvenanceOptions();
        return InspectCore(html, options, enforceUtf8Size: true);
    }

    private static OfficeProvenanceReport InspectCore(string html, OfficeProvenanceOptions options, bool enforceUtf8Size) {
        OfficeProvenanceBinary.ValidateLimits(options);
        if (enforceUtf8Size && Encoding.UTF8.GetByteCount(html) > options.MaxAssetBytes) {
            throw new InvalidDataException("The HTML document exceeds the configured asset limit.");
        }

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
                OfficeC2paManifestStore.IsValid(
                    manifest, 0, manifest.Length, options.MaxManifestBytes, options.MaxContainerEntries, out _);
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
        return InspectCore(DecodeHtml(data, out _, out _), options, enforceUtf8Size: false);
    }

    /// <summary>Removes selected HTML provenance and provenance in supported embedded image data URIs.</summary>
    public static OfficeProvenanceRemovalResult Remove(string html, OfficeProvenanceRemovalOptions? options = null) {
        if (html == null) throw new ArgumentNullException(nameof(html));
        options ??= new OfficeProvenanceRemovalOptions();
        return RemoveCore(html, options, enforceUtf8Size: true);
    }

    private static OfficeProvenanceRemovalResult RemoveCore(string html, OfficeProvenanceRemovalOptions options, bool enforceUtf8Size) {
        OfficeProvenanceBinary.ValidateLimits(options.Limits);
        if (options.MaxEmbeddedAssets <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxEmbeddedAssets));
        OfficeProvenanceOptions inspectionOptions = CreateInspectionOptions(options);
        OfficeProvenanceReport before = InspectCore(html, inspectionOptions, enforceUtf8Size);
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
                OfficeC2paManifestStore.IsValid(
                    manifest, 0, manifest.Length, options.Limits.MaxManifestBytes, options.Limits.MaxContainerEntries, out _);
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

        if (inspectionOptions.ProcessEmbeddedAssets) RemoveEmbeddedImages(document, options, changes);
        if (changes.Count == 0) {
            byte[] original = Encoding.UTF8.GetBytes(html);
            return new OfficeProvenanceRemovalResult(original, before, before, changes.AsReadOnly(), false);
        }

        string outputHtml = document.ToHtml();
        byte[] output = Encoding.UTF8.GetBytes(outputHtml);
        if (enforceUtf8Size && output.LongLength > options.Limits.MaxAssetBytes) throw new InvalidDataException("The rewritten HTML document exceeds the configured asset limit.");
        OfficeProvenanceReport after = InspectCore(outputHtml, inspectionOptions, enforceUtf8Size);
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
        OfficeProvenanceRemovalResult result = RemoveCore(html, options, enforceUtf8Size: false);
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
        InspectEmbeddedImages(document, options, evidence, diagnostics, ref count, srcDocDepth: 0);
    }

    private static void InspectEmbeddedImages(
        IHtmlDocument document,
        OfficeProvenanceOptions options,
        List<OfficeProvenanceEvidence> evidence,
        List<string> diagnostics,
        ref int count,
        int srcDocDepth) {
        IElement[] elements = GetEmbeddedImageElements(document).ToArray();
        HashSet<string> usedImageProperties = GetUsedCssImageCustomProperties(elements);
        foreach (IElement element in elements) {
            foreach (EmbeddedImageReference reference in GetEmbeddedImageReferences(element, usedImageProperties)) {
                if (!HtmlImageDataUri.TryParse(reference.Value, out HtmlImageDataUri dataUri)) continue;
                int index = count++;
                string location = $"HTML/{element.LocalName}[{reference.AttributeName}][{index}]";
                if (count > options.MaxEmbeddedAssets) throw new InvalidDataException("The HTML document exceeds the configured embedded-asset limit.");
                if (!dataUri.TryEstimateDecodedByteCount(out long estimatedBytes)) {
                    diagnostics.Add($"{location}: embedded image data URI could not be decoded.");
                    continue;
                }
                if (estimatedBytes > options.MaxAssetBytes) throw new InvalidDataException("An embedded HTML image exceeds the configured asset limit.");
                if (!TryDecodeEmbeddedImage(dataUri, out byte[] image)) {
                    diagnostics.Add($"{location}: embedded image data URI could not be decoded.");
                    continue;
                }
                if (image.LongLength > options.MaxAssetBytes) throw new InvalidDataException("An embedded HTML image exceeds the configured asset limit.");
                try {
                    OfficeProvenanceReport nested = OfficeProvenanceInspector.Inspect(image, "asset" + dataUri.FileExtension, CreateNestedOptions(options));
                    foreach (OfficeProvenanceEvidence item in nested.Evidence) AddEvidence(evidence, options, Prefix(location, item));
                    foreach (string diagnostic in nested.Diagnostics) diagnostics.Add($"{location}: {diagnostic}");
                } catch (Exception exception) when (exception is InvalidDataException || exception is System.Xml.XmlException) {
                    diagnostics.Add($"{location}: embedded image was preserved because inspection failed: {exception.Message}");
                }
            }
        }
        if (srcDocDepth >= HtmlConversionInputGuard.MaxSrcDocDepth) return;
        foreach (IElement iframe in document.QuerySelectorAll("iframe[srcdoc]")) {
            string? srcdoc = iframe.GetAttribute("srcdoc");
            if (srcdoc == null || string.IsNullOrWhiteSpace(srcdoc)) continue;
            IHtmlDocument nested = HtmlDocumentParser.ParseDocument(srcdoc);
            InspectEmbeddedImages(nested, options, evidence, diagnostics, ref count, srcDocDepth + 1);
        }
    }

    private static void RemoveEmbeddedImages(
        IHtmlDocument document,
        OfficeProvenanceRemovalOptions options,
        List<OfficeProvenanceChange> changes) {
        int count = 0;
        RemoveEmbeddedImages(document, options, changes, ref count, srcDocDepth: 0);
    }

    private static void RemoveEmbeddedImages(
        IHtmlDocument document,
        OfficeProvenanceRemovalOptions options,
        List<OfficeProvenanceChange> changes,
        ref int count,
        int srcDocDepth) {
        int maxEmbeddedAssets = Math.Min(options.MaxEmbeddedAssets, options.Limits.MaxEmbeddedAssets);
        IElement[] elements = GetEmbeddedImageElements(document).ToArray();
        HashSet<string> usedImageProperties = GetUsedCssImageCustomProperties(elements);
        foreach (IElement element in elements) {
            EmbeddedImageReference[] references = GetEmbeddedImageReferences(element, usedImageProperties).ToArray();
            var replacements = new List<(EmbeddedImageReference Reference, string Value)>();
            foreach (EmbeddedImageReference reference in references) {
                if (!HtmlImageDataUri.TryParse(reference.Value, out HtmlImageDataUri dataUri)) continue;
                int index = count++;
                if (count > maxEmbeddedAssets) throw new InvalidDataException("The HTML document exceeds the configured embedded-asset limit.");
                if (!dataUri.TryEstimateDecodedByteCount(out long estimatedBytes)) continue;
                if (estimatedBytes > options.Limits.MaxAssetBytes) throw new InvalidDataException("An embedded HTML image exceeds the configured asset limit.");
                if (!TryDecodeEmbeddedImage(dataUri, out byte[] image)) continue;
                if (image.LongLength > options.Limits.MaxAssetBytes) throw new InvalidDataException("An embedded HTML image exceeds the configured asset limit.");
                try {
                    OfficeProvenanceRemovalResult nested = OfficeProvenanceRemover.Remove(
                        image,
                        "asset" + dataUri.FileExtension,
                        CreateNestedRemovalOptions(options));
                    if (!nested.WasChanged) continue;
                    string metadata = CreateRewrittenDataUriMetadata(dataUri);
                    replacements.Add((reference, "data:" + metadata + "," + Convert.ToBase64String(nested.ToArray())));
                    foreach (OfficeProvenanceChange change in nested.Changes) {
                        changes.Add(new OfficeProvenanceChange(
                            change.Carrier,
                            $"HTML/{element.LocalName}[{reference.AttributeName}][{index}]/{change.Location}",
                            0));
                    }
                } catch (Exception exception) when (exception is InvalidDataException || exception is System.Xml.XmlException) {
                    // Preserve malformed embedded data; structural diagnostics are available through Inspect.
                }
            }
            ApplyEmbeddedImageReplacements(element, replacements);
        }
        if (srcDocDepth >= HtmlConversionInputGuard.MaxSrcDocDepth) return;
        foreach (IElement iframe in document.QuerySelectorAll("iframe[srcdoc]")) {
            string? srcdoc = iframe.GetAttribute("srcdoc");
            if (srcdoc == null || string.IsNullOrWhiteSpace(srcdoc)) continue;
            IHtmlDocument nested = HtmlDocumentParser.ParseDocument(srcdoc);
            int priorChanges = changes.Count;
            RemoveEmbeddedImages(nested, options, changes, ref count, srcDocDepth + 1);
            if (changes.Count != priorChanges) iframe.SetAttribute("srcdoc", nested.ToHtml());
        }
    }

    private static IEnumerable<EmbeddedImageReference> GetEmbeddedImageReferences(
        IElement element,
        ISet<string> usedImageProperties) {
        string localName = element.LocalName.ToLowerInvariant();
        if (localName == "style") {
            foreach (HtmlCssImageReference reference in HtmlResourcePipeline.EnumerateProvenanceCssImageReferences(
                element.TextContent,
                usedImageProperties)) {
                yield return new EmbeddedImageReference("css", reference.Value, reference.Start, reference.Length);
            }
            yield break;
        }

        string? inlineStyle = element.GetAttribute("style");
        if (inlineStyle != null) {
            foreach (HtmlCssImageReference reference in HtmlResourcePipeline.EnumerateProvenanceCssImageReferences(
                inlineStyle,
                usedImageProperties)) {
                yield return new EmbeddedImageReference("style", reference.Value, reference.Start, reference.Length);
            }
        }

        string? background = element.GetAttribute("background");
        if (background != null) yield return new EmbeddedImageReference("background", background, 0, background.Length);

        if (localName is "img" or "source") {
            foreach (string attributeName in EmbeddedImageSourceAttributes) {
                string? source = element.GetAttribute(attributeName);
                if (source != null) yield return new EmbeddedImageReference(attributeName, source, 0, source.Length);
            }
            foreach (string attributeName in EmbeddedImageSourceSetAttributes) {
                string? sourceSet = element.GetAttribute(attributeName);
                if (sourceSet == null) continue;
                foreach (EmbeddedImageReference reference in ParseSrcset(attributeName, sourceSet)) yield return reference;
            }
            yield break;
        }

        string[] attributeNames;
        if (localName == "video") {
            attributeNames = new[] { "poster", "data-poster" };
        } else if (localName == "input" && string.Equals(
            HtmlFormControlSemantics.GetEffectiveType("input", element.GetAttribute("type")),
            "image",
            StringComparison.Ordinal)) {
            attributeNames = new[] { "src", "data-src" };
        } else if (localName == "image") {
            attributeNames = new[] { "href", "xlink:href", "src" };
        } else if (localName is "feimage" or "use") {
            attributeNames = new[] { "href", "xlink:href" };
        } else if (localName == "link" && IsImageLink(element)) {
            attributeNames = new[] { "href" };
        } else {
            yield break;
        }

        foreach (string attributeName in attributeNames) {
            string? source = element.GetAttribute(attributeName);
            if (source != null) yield return new EmbeddedImageReference(attributeName, source, 0, source.Length);
        }
        if (localName == "link" && IsPreloadedImage(element)) {
            string? sourceSet = element.GetAttribute("imagesrcset");
            if (sourceSet != null) {
                foreach (EmbeddedImageReference reference in ParseSrcset("imagesrcset", sourceSet)) yield return reference;
            }
        }
    }

    private static IEnumerable<IElement> GetEmbeddedImageElements(IHtmlDocument document) =>
        document.QuerySelectorAll("img,source,video,input,image,feimage,use,link,[background],style,[style]").Distinct();

    private static HashSet<string> GetUsedCssImageCustomProperties(IEnumerable<IElement> elements) =>
        HtmlResourcePipeline.CollectProvenanceCssImageCustomProperties(elements.SelectMany(element => {
            var styles = new List<string>(2);
            if (string.Equals(element.LocalName, "style", StringComparison.OrdinalIgnoreCase)) styles.Add(element.TextContent);
            string? inlineStyle = element.GetAttribute("style");
            if (inlineStyle != null) styles.Add(inlineStyle);
            return styles;
        }));

    private static bool IsImageLink(IElement element) {
        string? rel = element.GetAttribute("rel");
        return HasRelationship(rel, "icon") || HasRelationship(rel, "apple-touch-icon") ||
            HasRelationship(rel, "shortcut icon") ||
            HasRelationship(rel, "shortcut") && HasRelationship(rel, "icon") || IsPreloadedImage(element);
    }

    private static bool IsPreloadedImage(IElement element) =>
        HasRelationship(element.GetAttribute("rel"), "preload") &&
        string.Equals(element.GetAttribute("as")?.Trim(), "image", StringComparison.OrdinalIgnoreCase);

    private static IEnumerable<EmbeddedImageReference> ParseSrcset(string attributeName, string sourceSet) {
        int searchOffset = 0;
        foreach (HtmlSrcSetCandidate candidate in HtmlSrcSetParser.Enumerate(sourceSet)) {
            int start = sourceSet.IndexOf(candidate.Url, searchOffset, StringComparison.Ordinal);
            if (start < 0) continue;
            yield return new EmbeddedImageReference(attributeName, candidate.Url, start, candidate.Url.Length);
            searchOffset = start + candidate.Url.Length;
        }
    }

    private static string CreateRewrittenDataUriMetadata(HtmlImageDataUri dataUri) {
        string[] parts = dataUri.Metadata.Split(';');
        var metadata = new List<string>(parts.Length + 1);
        bool svg = string.Equals(dataUri.MediaType, "image/svg+xml", StringComparison.OrdinalIgnoreCase);
        bool hasCharset = false;
        foreach (string part in parts) {
            string trimmed = part.Trim();
            if (trimmed.Equals("base64", StringComparison.OrdinalIgnoreCase)) continue;
            if (svg && trimmed.StartsWith("charset=", StringComparison.OrdinalIgnoreCase)) {
                metadata.Add("charset=utf-8");
                hasCharset = true;
            } else {
                metadata.Add(trimmed);
            }
        }
        if (svg && !hasCharset) metadata.Add("charset=utf-8");
        metadata.Add("base64");
        return string.Join(";", metadata);
    }

    private static bool TryDecodeEmbeddedImage(HtmlImageDataUri dataUri, out byte[] image) {
        if (!string.Equals(dataUri.MediaType, "image/svg+xml", StringComparison.OrdinalIgnoreCase)) {
            return dataUri.TryDecodeBytes(out image);
        }
        bool hasDeclaredCharset = dataUri.Metadata.Split(';')
            .Any(part => part.Trim().StartsWith("charset=", StringComparison.OrdinalIgnoreCase));
        if (!hasDeclaredCharset) return dataUri.TryDecodeBytes(out image);
        if (!dataUri.TryDecodeText(out string text)) {
            image = Array.Empty<byte>();
            return false;
        }
        int declarationEnd = text.StartsWith("<?xml", StringComparison.OrdinalIgnoreCase)
            ? text.IndexOf("?>", StringComparison.Ordinal)
            : -1;
        if (declarationEnd >= 0) {
            string declaration = text.Substring(0, declarationEnd + 2);
            string normalized = System.Text.RegularExpressions.Regex.Replace(
                declaration,
                "(\\bencoding\\s*=\\s*[\"'])[^\"']*([\"'])",
                "$1utf-8$2",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            text = normalized + text.Substring(declarationEnd + 2);
        }
        image = Encoding.UTF8.GetBytes(text);
        return true;
    }

    private static void ApplyEmbeddedImageReplacements(
        IElement element,
        List<(EmbeddedImageReference Reference, string Value)> replacements) {
        foreach (IGrouping<string, (EmbeddedImageReference Reference, string Value)> group in replacements.GroupBy(item => item.Reference.AttributeName)) {
            string value = group.Key == "css" ? element.TextContent : element.GetAttribute(group.Key) ?? string.Empty;
            foreach ((EmbeddedImageReference reference, string replacement) in group.OrderByDescending(item => item.Reference.Start)) {
                value = value.Substring(0, reference.Start) + replacement + value.Substring(reference.Start + reference.Length);
            }
            if (group.Key == "css") element.TextContent = value;
            else element.SetAttribute(group.Key, value);
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

    private static OfficeProvenanceOptions CreateInspectionOptions(OfficeProvenanceRemovalOptions source) => new OfficeProvenanceOptions {
        MaxAssetBytes = source.Limits.MaxAssetBytes,
        MaxManifestBytes = source.Limits.MaxManifestBytes,
        MaxCarriers = source.Limits.MaxCarriers,
        MaxContainerEntries = source.Limits.MaxContainerEntries,
        MaxExpandedContainerBytes = source.Limits.MaxExpandedContainerBytes,
        ProcessEmbeddedAssets = source.ProcessEmbeddedAssets && source.Limits.ProcessEmbeddedAssets,
        MaxEmbeddedAssets = Math.Min(source.MaxEmbeddedAssets, source.Limits.MaxEmbeddedAssets)
    };

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
        Encoding strictEncoding = (Encoding)encoding.Clone();
        strictEncoding.EncoderFallback = EncoderFallback.ExceptionFallback;
        string encodableHtml = EscapeUnencodableCharacters(html, strictEncoding);
        byte[] body = strictEncoding.GetBytes(encodableHtml);
        byte[] preamble = includePreamble ? encoding.GetPreamble() : Array.Empty<byte>();
        if (preamble.Length == 0) return body;
        byte[] output = new byte[preamble.Length + body.Length];
        Buffer.BlockCopy(preamble, 0, output, 0, preamble.Length);
        Buffer.BlockCopy(body, 0, output, preamble.Length, body.Length);
        return output;
    }

    private static string EscapeUnencodableCharacters(string value, Encoding encoding) {
        var builder = new StringBuilder(value.Length);
        char[] characters = value.ToCharArray();
        int cursor = 0;
        while (cursor < characters.Length) {
            try {
                _ = encoding.GetByteCount(characters, cursor, characters.Length - cursor);
                builder.Append(characters, cursor, characters.Length - cursor);
                break;
            } catch (EncoderFallbackException exception) {
                int invalidIndex = cursor + exception.Index;
                if (invalidIndex < cursor || invalidIndex >= characters.Length) invalidIndex = cursor;
                if (invalidIndex > cursor) builder.Append(characters, cursor, invalidIndex - cursor);
                int characterCount = char.IsHighSurrogate(characters[invalidIndex]) && invalidIndex + 1 < characters.Length &&
                    char.IsLowSurrogate(characters[invalidIndex + 1]) ? 2 : 1;
                int codePoint = characterCount == 2
                    ? char.ConvertToUtf32(characters[invalidIndex], characters[invalidIndex + 1])
                    : characters[invalidIndex];
                builder.Append("&#x").Append(codePoint.ToString("X", System.Globalization.CultureInfo.InvariantCulture)).Append(';');
                cursor = invalidIndex + characterCount;
            }
        }
        return builder.ToString();
    }

    private sealed class EmbeddedImageReference {
        internal EmbeddedImageReference(string attributeName, string value, int start, int length) {
            AttributeName = attributeName;
            Value = value;
            Start = start;
            Length = length;
        }

        internal string AttributeName { get; }
        internal string Value { get; }
        internal int Start { get; }
        internal int Length { get; }
    }
}
