using AngleSharp;
using AngleSharp.Html.Dom;
using OfficeIMO.Core.Internal;
using OfficeIMO.Provenance;
using System.Text.RegularExpressions;

namespace OfficeIMO.Html;

/// <summary>Inspects and selectively removes standards-defined provenance from HTML documents.</summary>
public static partial class HtmlProvenance {
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

        int structuralEntries = 0;
        IHtmlDocument document = ParseBoundedDocument(html, options.MaxContainerEntries, ref structuralEntries);
        var evidence = new List<OfficeProvenanceEvidence>();
        var diagnostics = new List<string>();
        long expandedBytes = 0;
        InspectManifestCarriers(document, options, evidence, diagnostics, "HTML", ref expandedBytes);
        int embeddedAssetCount = 0;
        InspectEmbeddedImages(document, options, evidence, diagnostics, ref embeddedAssetCount, ref structuralEntries,
            ref expandedBytes, "HTML", srcDocDepth: 0);
        return new OfficeProvenanceReport(OfficeProvenanceAssetFormat.Html, evidence.AsReadOnly(), diagnostics.AsReadOnly());
    }

    private static void InspectManifestCarriers(
        IHtmlDocument document,
        OfficeProvenanceOptions options,
        List<OfficeProvenanceEvidence> evidence,
        List<string> diagnostics,
        string documentLocation,
        ref long expandedBytes) {
        IElement? head = document.Head;
        if (head == null) return;
        IElement[] manifestElements = head.QuerySelectorAll("script[type],link[rel][href]")
            .Where(IsManifestElement)
            .ToArray();
        if (manifestElements.Length > 1) diagnostics.Add($"{documentLocation}: manifest.html.multipleManifests: the HTML head contains multiple C2PA manifest associations.");
        int carrierIndex = 0;
        foreach (IElement script in head.QuerySelectorAll("script[type]")) {
            if (!string.Equals(script.GetAttribute("type")?.Trim(), "application/c2pa", StringComparison.OrdinalIgnoreCase)) continue;
            TryDecodeManifest(
                script.TextContent,
                options.MaxManifestBytes,
                options.MaxExpandedContainerBytes,
                ref expandedBytes,
                out byte[] manifest);
            bool valid = manifestElements.Length == 1 && manifest.Length != 0 &&
                OfficeC2paManifestStore.IsValid(
                    manifest, 0, manifest.Length, options.MaxManifestBytes, options.MaxContainerEntries, out _);
            AddEvidence(evidence, options, new OfficeProvenanceEvidence(
                OfficeProvenanceCarrierKind.C2paManifest,
                $"{documentLocation}/script[type=application/c2pa][{carrierIndex++}]",
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
                $"{documentLocation}/link[rel=c2pa-manifest][{carrierIndex++}]",
                valid,
                0,
                valid && uri!.IsAbsoluteUri ? uri.AbsoluteUri : value));
        }

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
        return RemoveCore(html, options, enforceUtf8Size: true, outputEncoding: null, outputHadPreamble: false);
    }

    private static OfficeProvenanceRemovalResult RemoveCore(
        string html,
        OfficeProvenanceRemovalOptions options,
        bool enforceUtf8Size,
        Encoding? outputEncoding,
        bool outputHadPreamble) {
        OfficeProvenanceBinary.ValidateLimits(options.Limits);
        if (options.MaxEmbeddedAssets <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxEmbeddedAssets));
        OfficeProvenanceOptions inspectionOptions = CreateInspectionOptions(options);
        OfficeProvenanceReport before = InspectCore(html, inspectionOptions, enforceUtf8Size);
        int structuralEntries = 0;
        IHtmlDocument document = ParseBoundedDocument(html, options.Limits.MaxContainerEntries, ref structuralEntries);
        var changes = new List<OfficeProvenanceChange>();
        long expandedBytes = 0;
        RemoveManifestCarriers(document, options, changes, "HTML", ref expandedBytes);
        int embeddedAssetCount = 0;
        RemoveEmbeddedImages(document, options, changes, ref embeddedAssetCount, ref structuralEntries,
            ref expandedBytes, "HTML", srcDocDepth: 0);

        if (changes.Count == 0) {
            byte[] original;
            if (outputEncoding != null) {
                original = Array.Empty<byte>();
            } else {
                int byteCount = Encoding.UTF8.GetByteCount(html);
                if (byteCount > options.Limits.MaxAssetBytes) throw new InvalidDataException("The HTML document exceeds the configured asset limit.");
                original = Encoding.UTF8.GetBytes(html);
            }
            return new OfficeProvenanceRemovalResult(original, before, before, changes.AsReadOnly(), false);
        }

        if (enforceUtf8Size) NormalizeDeclaredEncodingToUtf8(document);
        string outputHtml = document.ToHtml();
        byte[] output;
        if (outputEncoding == null) {
            int byteCount = Encoding.UTF8.GetByteCount(outputHtml);
            if (byteCount > options.Limits.MaxAssetBytes) throw new InvalidDataException("The rewritten HTML document exceeds the configured asset limit.");
            output = Encoding.UTF8.GetBytes(outputHtml);
        } else {
            output = EncodeHtml(outputHtml, outputEncoding, outputHadPreamble, options.Limits.MaxAssetBytes);
        }
        OfficeProvenanceReport after = InspectCore(outputHtml, inspectionOptions, enforceUtf8Size);
        return new OfficeProvenanceRemovalResult(output, before, after, changes.AsReadOnly(), true);
    }

    private static void RemoveManifestCarriers(
        IHtmlDocument document,
        OfficeProvenanceRemovalOptions options,
        List<OfficeProvenanceChange> changes,
        string documentLocation,
        ref long expandedBytes) {
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
            TryDecodeManifest(
                script.TextContent,
                options.Limits.MaxManifestBytes,
                options.Limits.MaxExpandedContainerBytes,
                ref expandedBytes,
                out byte[] manifest);
            bool valid = manifestElementCount == 1 && manifest.Length != 0 &&
                OfficeC2paManifestStore.IsValid(
                    manifest, 0, manifest.Length, options.Limits.MaxManifestBytes, options.Limits.MaxContainerEntries, out _);
            string location = $"{documentLocation}/script[type=application/c2pa][{carrierIndex++}]";
            if (!options.RemoveC2paManifests || (!valid && options.RequireStructurallyValidCarrier)) continue;
            script.Remove();
            changes.Add(new OfficeProvenanceChange(OfficeProvenanceCarrierKind.C2paManifest, location, 0));
        }

        foreach (IElement link in links.ToArray()) {
            if (!HasRelationship(link.GetAttribute("rel"), "c2pa-manifest")) continue;
            string value = link.GetAttribute("href")?.Trim() ?? string.Empty;
            bool valid = manifestElementCount == 1 && IsSafeManifestReference(value, out _);
            string location = $"{documentLocation}/link[rel=c2pa-manifest][{carrierIndex++}]";
            if (!options.RemoveExternalC2paReferences || (!valid && options.RequireStructurallyValidCarrier)) continue;
            RemoveRelationship(link, "c2pa-manifest");
            changes.Add(new OfficeProvenanceChange(OfficeProvenanceCarrierKind.C2paExternalManifest, location, 0));
        }

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
        OfficeProvenanceRemovalResult result = RemoveCore(
            html, options, enforceUtf8Size: false, outputEncoding: encoding, outputHadPreamble: hadPreamble);
        if (!result.WasChanged) {
            OfficeFileCommit.WriteAllBytes(Path.GetFullPath(outputPath), input);
            return new OfficeProvenanceRemovalResult(
                (byte[])input.Clone(), result.Before, result.After, result.Changes, result.WasReserialized);
        }

        byte[] output = result.ToArray();
        OfficeFileCommit.WriteAllBytes(Path.GetFullPath(outputPath), output);
        return new OfficeProvenanceRemovalResult(
            output, result.Before, result.After, result.Changes, result.WasReserialized);
    }

    private static void InspectEmbeddedImages(
        IHtmlDocument document,
        OfficeProvenanceOptions options,
        List<OfficeProvenanceEvidence> evidence,
        List<string> diagnostics,
        ref int count,
        ref int structuralEntries,
        ref long expandedBytes,
        string documentLocation,
        int srcDocDepth) {
        if (options.ProcessEmbeddedAssets) {
            IElement[] elements = GetEmbeddedImageElements(document).ToArray();
            HtmlProvenanceCssScope cssScope = HtmlResourcePipeline.CollectProvenanceCssImageScope(document);
            foreach (IElement element in elements) {
                cssScope.UsedCustomPropertyDeclarations.TryGetValue(element, out HashSet<int>? usedDeclarations);
                cssScope.ResolvedVarFallbackStarts.TryGetValue(element, out HashSet<int>? resolvedFallbacks);
                foreach (EmbeddedImageReference reference in GetEmbeddedImageReferences(
                    document, element, usedDeclarations, resolvedFallbacks)) {
                    if (!HtmlImageDataUri.TryParse(reference.Value, out HtmlImageDataUri dataUri)) continue;
                    if (!IsSupportedProvenanceImage(dataUri.MediaType)) continue;
                    int index = count++;
                    string location = $"{documentLocation}/{element.LocalName}[{reference.AttributeName}][{index}]";
                    if (count > options.MaxEmbeddedAssets) throw new InvalidDataException("The HTML document exceeds the configured embedded-asset limit.");
                    if (!dataUri.TryEstimateDecodedByteCount(out long estimatedBytes)) {
                        diagnostics.Add($"{location}: embedded image data URI could not be decoded.");
                        continue;
                    }
                    if (estimatedBytes > options.MaxAssetBytes) throw new InvalidDataException("An embedded HTML image exceeds the configured asset limit.");
                    ReserveExpandedBytes(ref expandedBytes, estimatedBytes, options.MaxExpandedContainerBytes);
                    if (!TryDecodeEmbeddedImage(dataUri, options.MaxAssetBytes, out byte[] image)) {
                        diagnostics.Add($"{location}: embedded image data URI could not be decoded.");
                        continue;
                    }
                    if (image.LongLength > options.MaxAssetBytes) throw new InvalidDataException("An embedded HTML image exceeds the configured asset limit.");
                    if (image.LongLength > estimatedBytes) {
                        ReserveExpandedBytes(ref expandedBytes, image.LongLength - estimatedBytes, options.MaxExpandedContainerBytes);
                    }
                    try {
                        OfficeProvenanceReport nested = OfficeProvenanceInspector.Inspect(image, "asset" + dataUri.FileExtension, CreateNestedOptions(options));
                        foreach (OfficeProvenanceEvidence item in nested.Evidence) AddEvidence(evidence, options, Prefix(location, item));
                        foreach (string diagnostic in nested.Diagnostics) diagnostics.Add($"{location}: {diagnostic}");
                    } catch (Exception exception) when (exception is InvalidDataException || exception is System.Xml.XmlException) {
                        diagnostics.Add($"{location}: embedded image was preserved because inspection failed: {exception.Message}");
                    }
                }
            }
        }
        if (srcDocDepth >= HtmlConversionInputGuard.MaxSrcDocDepth) return;
        int iframeIndex = 0;
        foreach (IElement iframe in document.QuerySelectorAll("iframe[srcdoc]")) {
            string? srcdoc = iframe.GetAttribute("srcdoc");
            if (srcdoc == null || string.IsNullOrWhiteSpace(srcdoc)) continue;
            string location = $"{documentLocation}/iframe[srcdoc][{iframeIndex++}]";
            IHtmlDocument nested = ParseBoundedDocument(srcdoc, options.MaxContainerEntries, ref structuralEntries);
            InspectManifestCarriers(nested, options, evidence, diagnostics, location, ref expandedBytes);
            InspectEmbeddedImages(nested, options, evidence, diagnostics, ref count, ref structuralEntries,
                ref expandedBytes, location, srcDocDepth + 1);
        }
    }

    private static void RemoveEmbeddedImages(
        IHtmlDocument document,
        OfficeProvenanceRemovalOptions options,
        List<OfficeProvenanceChange> changes,
        ref int count,
        ref int structuralEntries,
        ref long expandedBytes,
        string documentLocation,
        int srcDocDepth) {
        if (options.ProcessEmbeddedAssets && options.Limits.ProcessEmbeddedAssets) {
            int maxEmbeddedAssets = Math.Min(options.MaxEmbeddedAssets, options.Limits.MaxEmbeddedAssets);
            IElement[] elements = GetEmbeddedImageElements(document).ToArray();
            HtmlProvenanceCssScope cssScope = HtmlResourcePipeline.CollectProvenanceCssImageScope(document);
            foreach (IElement element in elements) {
                cssScope.UsedCustomPropertyDeclarations.TryGetValue(element, out HashSet<int>? usedDeclarations);
                cssScope.ResolvedVarFallbackStarts.TryGetValue(element, out HashSet<int>? resolvedFallbacks);
                EmbeddedImageReference[] references = GetEmbeddedImageReferences(
                    document, element, usedDeclarations, resolvedFallbacks).ToArray();
                var replacements = new List<(EmbeddedImageReference Reference, string Value)>();
                foreach (EmbeddedImageReference reference in references) {
                    if (!HtmlImageDataUri.TryParse(reference.Value, out HtmlImageDataUri dataUri)) continue;
                    if (!IsSupportedProvenanceImage(dataUri.MediaType)) continue;
                    int index = count++;
                    if (count > maxEmbeddedAssets) throw new InvalidDataException("The HTML document exceeds the configured embedded-asset limit.");
                    if (!dataUri.TryEstimateDecodedByteCount(out long estimatedBytes)) continue;
                    if (estimatedBytes > options.Limits.MaxAssetBytes) throw new InvalidDataException("An embedded HTML image exceeds the configured asset limit.");
                    ReserveExpandedBytes(ref expandedBytes, estimatedBytes, options.Limits.MaxExpandedContainerBytes);
                    if (!TryDecodeEmbeddedImage(dataUri, options.Limits.MaxAssetBytes, out byte[] image)) continue;
                    if (image.LongLength > options.Limits.MaxAssetBytes) throw new InvalidDataException("An embedded HTML image exceeds the configured asset limit.");
                    if (image.LongLength > estimatedBytes) {
                        ReserveExpandedBytes(ref expandedBytes, image.LongLength - estimatedBytes, options.Limits.MaxExpandedContainerBytes);
                    }
                    try {
                        OfficeProvenanceRemovalResult nested = OfficeProvenanceRemover.Remove(
                            image,
                            "asset" + dataUri.FileExtension,
                            CreateNestedRemovalOptions(options));
                        if (!nested.WasChanged) continue;
                        string metadata = CreateRewrittenDataUriMetadata(dataUri);
                        replacements.Add((reference, "data:" + metadata + "," + Convert.ToBase64String(nested.ToArray()) + dataUri.Fragment));
                        foreach (OfficeProvenanceChange change in nested.Changes) {
                            changes.Add(new OfficeProvenanceChange(
                                change.Carrier,
                                $"{documentLocation}/{element.LocalName}[{reference.AttributeName}][{index}]/{change.Location}",
                                0));
                        }
                    } catch (Exception exception) when (exception is InvalidDataException || exception is System.Xml.XmlException) {
                        // Preserve malformed embedded data; structural diagnostics are available through Inspect.
                    }
                }
                ApplyEmbeddedImageReplacements(element, replacements);
            }
        }
        if (srcDocDepth >= HtmlConversionInputGuard.MaxSrcDocDepth) return;
        int iframeIndex = 0;
        foreach (IElement iframe in document.QuerySelectorAll("iframe[srcdoc]")) {
            string? srcdoc = iframe.GetAttribute("srcdoc");
            if (srcdoc == null || string.IsNullOrWhiteSpace(srcdoc)) continue;
            string location = $"{documentLocation}/iframe[srcdoc][{iframeIndex++}]";
            IHtmlDocument nested = ParseBoundedDocument(srcdoc, options.Limits.MaxContainerEntries, ref structuralEntries);
            int priorChanges = changes.Count;
            RemoveManifestCarriers(nested, options, changes, location, ref expandedBytes);
            RemoveEmbeddedImages(nested, options, changes, ref count, ref structuralEntries,
                ref expandedBytes, location, srcDocDepth + 1);
            if (changes.Count != priorChanges) iframe.SetAttribute("srcdoc", nested.ToHtml());
        }
    }

    private static IEnumerable<EmbeddedImageReference> GetEmbeddedImageReferences(
        IHtmlDocument document,
        IElement element,
        ISet<int>? usedImageProperties,
        ISet<int>? resolvedVarFallbacks) {
        string localName = element.LocalName.ToLowerInvariant();
        if (localName == "style") {
            if (!HtmlResourcePipeline.IsActiveProvenanceStyleElement(element)) yield break;
            foreach (HtmlCssImageReference reference in HtmlResourcePipeline.EnumerateProvenanceCssImageReferences(
                document,
                "css",
                element.TextContent,
                usedImageProperties,
                resolvedVarFallbacks)) {
                yield return new EmbeddedImageReference("css", reference.Value, reference.Start, reference.Length);
            }
            yield break;
        }

        string? inlineStyle = element.GetAttribute("style");
        if (inlineStyle != null) {
            foreach (HtmlCssImageReference reference in HtmlResourcePipeline.EnumerateProvenanceCssImageReferences(
                document,
                "style",
                inlineStyle,
                usedImageProperties,
                resolvedVarFallbacks)) {
                yield return new EmbeddedImageReference("style", reference.Value, reference.Start, reference.Length);
            }
        }

        string? background = element.GetAttribute("background");
        if (background != null && HtmlResourcePipeline.SupportsLegacyBackground(localName)) {
            yield return CreateDirectUrlReference("background", background);
        }

        if (localName == "source" && !IsSupportedPictureSource(element)) yield break;
        if (localName == "img" && !HtmlResourcePipeline.IsActivePictureFallbackImage(element)) yield break;
        if (localName is "img" or "source") {
            foreach (string attributeName in EmbeddedImageSourceAttributes) {
                string? source = element.GetAttribute(attributeName);
                if (source != null) yield return CreateDirectUrlReference(attributeName, source);
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
            if (source != null) yield return CreateDirectUrlReference(attributeName, source);
        }
        if (localName == "link" && IsPreloadedImage(element)) {
            string? sourceSet = element.GetAttribute("imagesrcset");
            if (sourceSet != null) {
                foreach (EmbeddedImageReference reference in ParseSrcset("imagesrcset", sourceSet)) yield return reference;
            }
        }
    }

    private static IEnumerable<IElement> GetEmbeddedImageElements(IHtmlDocument document) =>
        document.QuerySelectorAll("img,source,video,input,image,feImage,use,link,[background],style,[style]")
            .Where(element => !element.LocalName.Equals("style", StringComparison.OrdinalIgnoreCase) ||
                HtmlResourcePipeline.IsActiveProvenanceStyleElement(element))
            .Where(element => !element.LocalName.Equals("source", StringComparison.OrdinalIgnoreCase) ||
                IsSupportedPictureSource(element))
            .Distinct();

    private static bool IsSupportedPictureSource(IElement element) {
        string? parentName = element.ParentElement?.LocalName;
        if (parentName is "audio" or "video") return false;
        return !string.Equals(parentName, "picture", StringComparison.OrdinalIgnoreCase) ||
            HtmlResourcePipeline.IsActivePictureImageSource(element);
    }

    private static void NormalizeDeclaredEncodingToUtf8(IHtmlDocument document) {
        foreach (IElement meta in document.QuerySelectorAll("meta[charset]")) meta.SetAttribute("charset", "utf-8");
        foreach (IElement meta in document.QuerySelectorAll("meta[http-equiv][content]")) {
            if (!string.Equals(meta.GetAttribute("http-equiv")?.Trim(), "content-type", StringComparison.OrdinalIgnoreCase)) continue;
            string content = meta.GetAttribute("content") ?? string.Empty;
            meta.SetAttribute("content", Regex.Replace(content, "(?i)(charset\\s*=\\s*)[^;\\s]+", "$1utf-8"));
        }
    }

    private static EmbeddedImageReference CreateDirectUrlReference(string attributeName, string value) {
        int start = 0;
        while (start < value.Length && char.IsWhiteSpace(value[start])) start++;
        int end = value.Length;
        while (end > start && char.IsWhiteSpace(value[end - 1])) end--;
        return new EmbeddedImageReference(attributeName, value.Substring(start, end - start), start, end - start);
    }

    private static bool IsImageLink(IElement element) {
        string? rel = element.GetAttribute("rel");
        return HasRelationship(rel, "icon") || HasRelationship(rel, "apple-touch-icon") ||
            HasRelationship(rel, "shortcut icon") ||
            HasRelationship(rel, "shortcut") && HasRelationship(rel, "icon") || IsPreloadedImage(element);
    }

    private static bool IsPreloadedImage(IElement element) =>
        HasRelationship(element.GetAttribute("rel"), "preload") &&
        string.Equals(element.GetAttribute("as")?.Trim(), "image", StringComparison.OrdinalIgnoreCase) &&
        HtmlResourcePipeline.IsApplicableProvenanceMedia(element);

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

    private static bool TryDecodeEmbeddedImage(HtmlImageDataUri dataUri, long maximumBytes, out byte[] image) {
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
        int declarationStart = text.Length > 0 && text[0] == '\uFEFF' ? 1 : 0;
        int declarationEnd = text.IndexOf("?>", declarationStart, StringComparison.Ordinal);
        if (declarationEnd < 0 || text.IndexOf("<?xml", declarationStart, StringComparison.OrdinalIgnoreCase) != declarationStart) {
            declarationEnd = -1;
        }
        if (declarationEnd >= 0) {
            string declaration = text.Substring(declarationStart, declarationEnd + 2 - declarationStart);
            string normalized = System.Text.RegularExpressions.Regex.Replace(
                declaration,
                "(\\bencoding\\s*=\\s*[\"'])[^\"']*([\"'])",
                "$1utf-8$2",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            text = text.Substring(0, declarationStart) + normalized + text.Substring(declarationEnd + 2);
        }
        int byteCount = Encoding.UTF8.GetByteCount(text);
        if (byteCount > maximumBytes) throw new InvalidDataException("An embedded HTML image exceeds the configured asset limit.");
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

    private static bool TryDecodeManifest(
        string? value,
        long maximumBytes,
        long maximumExpandedBytes,
        ref long expandedBytes,
        out byte[] manifest) {
        manifest = Array.Empty<byte>();
        string encoded = (value ?? string.Empty).Trim();
        const string prefix = "data:application/c2pa;base64,";
        if (encoded.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) encoded = encoded.Substring(prefix.Length);
        if (encoded.Length == 0 || encoded.Length > maximumBytes * 2L || encoded.Length > int.MaxValue) return false;
        if (!TryEstimateBase64DecodedByteCount(encoded, out long estimatedBytes) || estimatedBytes > maximumBytes) return false;
        ReserveExpandedBytes(ref expandedBytes, estimatedBytes, maximumExpandedBytes);
        try {
            manifest = Convert.FromBase64String(encoded);
            if (manifest.LongLength > estimatedBytes) {
                ReserveExpandedBytes(ref expandedBytes, manifest.LongLength - estimatedBytes, maximumExpandedBytes);
            }
            return manifest.LongLength <= maximumBytes;
        } catch (FormatException) {
            manifest = Array.Empty<byte>();
            return false;
        }
    }

    private static bool TryEstimateBase64DecodedByteCount(string encoded, out long decodedBytes) {
        decodedBytes = 0;
        int characterCount = 0;
        int padding = 0;
        bool sawPadding = false;
        foreach (char character in encoded) {
            if (char.IsWhiteSpace(character)) continue;
            characterCount++;
            if (character == '=') {
                sawPadding = true;
                padding++;
                if (padding > 2) return false;
            } else if (sawPadding) {
                return false;
            }
        }
        if (characterCount == 0 || (characterCount & 3) != 0) return false;
        decodedBytes = (long)(characterCount / 4) * 3 - padding;
        return decodedBytes >= 0;
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

    private static IHtmlDocument ParseBoundedDocument(string html, int maximumEntries, ref int structuralEntries) {
        int remaining = maximumEntries - structuralEntries;
        if (remaining <= 0) throw new InvalidDataException("The HTML document exceeds the configured container-entry limit.");
        ValidatePotentialElementCount(html, remaining);
        IHtmlDocument document = HtmlDocumentParser.ParseDocument(html);
        int elementCount = document.All.Length;
        if (elementCount > remaining) throw new InvalidDataException("The HTML document exceeds the configured container-entry limit.");
        structuralEntries += elementCount;
        return document;
    }

    private static void ValidatePotentialElementCount(string html, int maximumEntries) {
        int count = 0;
        int index = 0;
        var openElements = new List<HtmlPreflightElement>();
        while (index < html.Length - 1) {
            int markup = html.IndexOf('<', index);
            if (markup < 0 || markup == html.Length - 1) break;
            if (markup <= html.Length - 4 && string.CompareOrdinal(html, markup, "<!--", 0, 4) == 0) {
                int commentEnd = FindHtmlCommentEnd(html, markup + 4);
                index = commentEnd < 0 ? html.Length : commentEnd;
                continue;
            }
            char next = html[markup + 1];
            if (next == '?') {
                int declarationEnd = html.IndexOf('>', markup + 2);
                index = declarationEnd < 0 ? html.Length : declarationEnd + 1;
                continue;
            }
            if (next == '!') {
                if (ChildNamespace(openElements) != HtmlPreflightNamespace.Html && markup <= html.Length - 9 &&
                    string.CompareOrdinal(html, markup, "<![CDATA[", 0, 9) == 0) {
                    int cdataEnd = html.IndexOf("]]>", markup + 9, StringComparison.Ordinal);
                    index = cdataEnd < 0 ? html.Length : cdataEnd + 3;
                } else if (markup <= html.Length - 9 &&
                    string.Compare(html, markup, "<!DOCTYPE", 0, 9, StringComparison.OrdinalIgnoreCase) == 0) {
                    int declarationEnd = FindTagEnd(html, markup + 9);
                    index = declarationEnd < 0 ? html.Length : declarationEnd + 1;
                } else {
                    int declarationEnd = html.IndexOf('>', markup + 2);
                    index = declarationEnd < 0 ? html.Length : declarationEnd + 1;
                }
                continue;
            }
            if (next == '/') {
                int nameStart = markup + 2;
                int closingNameEnd = nameStart;
                while (closingNameEnd < html.Length && (char.IsLetterOrDigit(html[closingNameEnd]) || html[closingNameEnd] is '-' or ':')) closingNameEnd++;
                string closingName = html.Substring(nameStart, closingNameEnd - nameStart);
                for (int elementIndex = openElements.Count - 1; elementIndex >= 0; elementIndex--) {
                    if (!openElements[elementIndex].Name.Equals(closingName, StringComparison.OrdinalIgnoreCase)) continue;
                    openElements.RemoveRange(elementIndex, openElements.Count - elementIndex);
                    break;
                }
                int declarationEnd = FindTagEnd(html, closingNameEnd);
                index = declarationEnd < 0 ? html.Length : declarationEnd + 1;
                continue;
            }
            if (!IsAsciiLetter(next)) {
                index = markup + 1;
                continue;
            }
            if (++count > maximumEntries) {
                throw new InvalidDataException("The HTML document exceeds the configured container-entry limit.");
            }
            int nameEnd = markup + 2;
            while (nameEnd < html.Length && (char.IsLetterOrDigit(html[nameEnd]) || html[nameEnd] is '-' or ':')) nameEnd++;
            string tagName = html.Substring(markup + 1, nameEnd - markup - 1);
            int tagEnd = FindTagEnd(html, nameEnd);
            if (tagEnd < 0) break;
            bool selfClosing = tagEnd > markup && html[tagEnd - 1] == '/';
            HtmlPreflightNamespace elementNamespace = ChildNamespace(openElements, tagName);
            bool childrenUseHtml = elementNamespace == HtmlPreflightNamespace.Html ||
                IsHtmlIntegrationPoint(html, tagName, elementNamespace, nameEnd, tagEnd);
            if (!selfClosing) openElements.Add(new HtmlPreflightElement(tagName, elementNamespace, childrenUseHtml));
            index = tagEnd + 1;
            if (elementNamespace == HtmlPreflightNamespace.Html && tagName.Equals("plaintext", StringComparison.OrdinalIgnoreCase)) return;
            if (elementNamespace == HtmlPreflightNamespace.Html && IsRawTextOrRcDataElement(tagName)) {
                int rawTextEnd = FindRawTextClosingTag(html, index, tagName);
                if (rawTextEnd < 0) break;
                index = rawTextEnd;
            }
        }
    }

    private static bool IsSupportedProvenanceImage(string mediaType) => mediaType.ToLowerInvariant() is
        "image/jpeg" or "image/jpg" or "image/png" or "image/gif" or "image/tiff" or "image/webp" or "image/svg+xml";

    private static HtmlPreflightNamespace ChildNamespace(List<HtmlPreflightElement> elements, string? tagName = null) {
        if (elements.Count == 0 || elements[elements.Count - 1].ChildrenUseHtml) {
            if (tagName?.Equals("svg", StringComparison.OrdinalIgnoreCase) == true) return HtmlPreflightNamespace.Svg;
            if (tagName?.Equals("math", StringComparison.OrdinalIgnoreCase) == true) return HtmlPreflightNamespace.MathMl;
            return HtmlPreflightNamespace.Html;
        }
        return elements[elements.Count - 1].Namespace;
    }

    private static bool IsHtmlIntegrationPoint(
        string html,
        string tagName,
        HtmlPreflightNamespace elementNamespace,
        int attributesStart,
        int tagEnd) {
        if (elementNamespace == HtmlPreflightNamespace.Svg) {
            return tagName.Equals("foreignObject", StringComparison.OrdinalIgnoreCase) ||
                tagName.Equals("desc", StringComparison.OrdinalIgnoreCase) ||
                tagName.Equals("title", StringComparison.OrdinalIgnoreCase);
        }
        if (elementNamespace != HtmlPreflightNamespace.MathMl) return false;
        if (tagName.Equals("mi", StringComparison.OrdinalIgnoreCase) ||
            tagName.Equals("mo", StringComparison.OrdinalIgnoreCase) ||
            tagName.Equals("mn", StringComparison.OrdinalIgnoreCase) ||
            tagName.Equals("ms", StringComparison.OrdinalIgnoreCase) ||
            tagName.Equals("mtext", StringComparison.OrdinalIgnoreCase)) return true;
        if (!tagName.Equals("annotation-xml", StringComparison.OrdinalIgnoreCase)) return false;
        string attributes = html.Substring(attributesStart, tagEnd - attributesStart);
        Match encoding = Regex.Match(
            attributes,
            "(?:^|\\s)encoding\\s*=\\s*(?:\"(?<value>[^\"]*)\"|'(?<value>[^']*)'|(?<value>[^\\s/>]+))",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant,
            TimeSpan.FromMilliseconds(100));
        string value = encoding.Success ? encoding.Groups["value"].Value : string.Empty;
        return value.Equals("text/html", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("application/xhtml+xml", StringComparison.OrdinalIgnoreCase);
    }

    private enum HtmlPreflightNamespace { Html, Svg, MathMl }

    private readonly struct HtmlPreflightElement {
        internal HtmlPreflightElement(string name, HtmlPreflightNamespace @namespace, bool childrenUseHtml) {
            Name = name;
            Namespace = @namespace;
            ChildrenUseHtml = childrenUseHtml;
        }
        internal string Name { get; }
        internal HtmlPreflightNamespace Namespace { get; }
        internal bool ChildrenUseHtml { get; }
    }

    private static bool IsRawTextOrRcDataElement(string tagName) =>
        tagName.Equals("script", StringComparison.OrdinalIgnoreCase) ||
        tagName.Equals("style", StringComparison.OrdinalIgnoreCase) ||
        tagName.Equals("xmp", StringComparison.OrdinalIgnoreCase) ||
        tagName.Equals("iframe", StringComparison.OrdinalIgnoreCase) ||
        tagName.Equals("noembed", StringComparison.OrdinalIgnoreCase) ||
        tagName.Equals("noframes", StringComparison.OrdinalIgnoreCase) ||
        tagName.Equals("textarea", StringComparison.OrdinalIgnoreCase) ||
        tagName.Equals("title", StringComparison.OrdinalIgnoreCase);

    private static bool IsAsciiLetter(char value) =>
        value >= 'A' && value <= 'Z' || value >= 'a' && value <= 'z';

    private static int FindRawTextClosingTag(string html, int offset, string tagName) {
        string closingPrefix = "</" + tagName;
        int candidate = offset;
        while (candidate < html.Length) {
            candidate = html.IndexOf(closingPrefix, candidate, StringComparison.OrdinalIgnoreCase);
            if (candidate < 0) return -1;
            int delimiter = candidate + closingPrefix.Length;
            if (delimiter >= html.Length || char.IsWhiteSpace(html[delimiter]) || html[delimiter] is '>' or '/') return candidate;
            candidate = delimiter;
        }
        return -1;
    }

    private static int FindTagEnd(string html, int offset) {
        char quote = '\0';
        for (int index = offset; index < html.Length; index++) {
            char current = html[index];
            if (quote != '\0') {
                if (current == quote) quote = '\0';
                continue;
            }
            if (current is '\'' or '"') quote = current;
            else if (current == '>') return index;
        }
        return -1;
    }

    private static void AddEvidence(List<OfficeProvenanceEvidence> evidence, OfficeProvenanceOptions options, OfficeProvenanceEvidence item) {
        if (evidence.Count >= options.MaxCarriers) throw new InvalidDataException($"The asset exceeds the configured carrier limit of {options.MaxCarriers}.");
        evidence.Add(item);
    }

    private static void ReserveExpandedBytes(ref long expandedBytes, long additionalBytes, long maximumBytes) {
        if (additionalBytes < 0 || expandedBytes > maximumBytes - additionalBytes) {
            throw new InvalidDataException("HTML provenance payloads exceed the configured expanded-container limit.");
        }
        expandedBytes += additionalBytes;
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

    private static byte[] EncodeHtml(string html, Encoding encoding, bool includePreamble, long maximumBytes) {
        Encoding strictEncoding = (Encoding)encoding.Clone();
        strictEncoding.EncoderFallback = EncoderFallback.ExceptionFallback;
        string encodableHtml = EscapeUnencodableCharacters(html, strictEncoding);
        byte[] preamble = includePreamble ? encoding.GetPreamble() : Array.Empty<byte>();
        int bodyLength = strictEncoding.GetByteCount(encodableHtml);
        if (bodyLength > maximumBytes - preamble.Length) {
            throw new InvalidDataException("The rewritten HTML document exceeds the configured asset limit.");
        }
        byte[] body = strictEncoding.GetBytes(encodableHtml);
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
