using AngleSharp;
using AngleSharp.Html.Dom;
using OfficeIMO.Core.Internal;
using OfficeIMO.Provenance;
using System.Text.RegularExpressions;
using System.Threading;

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
        return InspectCore(html, options, enforceUtf8Size: true, documentUri: null);
    }

    private static OfficeProvenanceReport InspectCore(
        string html,
        OfficeProvenanceOptions options,
        bool enforceUtf8Size,
        Uri? documentUri = null) {
        OfficeProvenanceBinary.ValidateLimits(options);
        options.CancellationToken.ThrowIfCancellationRequested();
        if (enforceUtf8Size && Encoding.UTF8.GetByteCount(html) > options.MaxAssetBytes) {
            throw new InvalidDataException("The HTML document exceeds the configured asset limit.");
        }

        int structuralEntries = 0;
        IHtmlDocument document = ParseBoundedDocument(html, options.MaxContainerEntries, ref structuralEntries, options.CancellationToken);
        var evidence = new List<OfficeProvenanceEvidence>();
        var diagnostics = new List<string>();
        var resolvedExternalManifestReferences = new Dictionary<OfficeProvenanceEvidence, string>();
        Uri? effectiveBaseUri = HtmlDocumentParser.ResolveEffectiveBaseUri(document, documentUri);
        Uri? sourceDirectoryUri = documentUri is null ? null : new Uri(documentUri, ".");
        long expandedBytes = 0;
        InspectManifestCarriers(
            document,
            options,
            evidence,
            diagnostics,
            resolvedExternalManifestReferences,
            effectiveBaseUri,
            sourceDirectoryUri,
            "HTML",
            ref expandedBytes);
        int embeddedAssetCount = 0;
        InspectEmbeddedImages(document, options, evidence, diagnostics, ref embeddedAssetCount, ref structuralEntries,
            ref expandedBytes, resolvedExternalManifestReferences, effectiveBaseUri, sourceDirectoryUri,
            "HTML", srcDocDepth: 0);
        return new OfficeProvenanceReport(
            OfficeProvenanceAssetFormat.Html,
            evidence.AsReadOnly(),
            diagnostics.AsReadOnly(),
            expandedBytes,
            resolvedExternalManifestReferences);
    }

    private static void InspectManifestCarriers(
        IHtmlDocument document,
        OfficeProvenanceOptions options,
        List<OfficeProvenanceEvidence> evidence,
        List<string> diagnostics,
        Dictionary<OfficeProvenanceEvidence, string> resolvedExternalManifestReferences,
        Uri? effectiveBaseUri,
        Uri? sourceDirectoryUri,
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
            options.CancellationToken.ThrowIfCancellationRequested();
            if (!string.Equals(TrimAsciiWhitespace(script.GetAttribute("type")), "application/c2pa", StringComparison.OrdinalIgnoreCase)) continue;
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
            options.CancellationToken.ThrowIfCancellationRequested();
            if (!HasRelationship(link.GetAttribute("rel"), "c2pa-manifest")) continue;
            string value = TrimAsciiWhitespace(link.GetAttribute("href"));
            bool safeReference = IsSafeManifestReference(value, out Uri? uri);
            bool valid = manifestElements.Length == 1 && safeReference;
            var item = new OfficeProvenanceEvidence(
                OfficeProvenanceCarrierKind.C2paExternalManifest,
                $"{documentLocation}/link[rel=c2pa-manifest][{carrierIndex++}]",
                valid,
                0,
                valid && uri!.IsAbsoluteUri ? uri.AbsoluteUri : value);
            AddEvidence(evidence, options, item);
            if (valid) {
                string resolved = ResolveExternalManifestReference(value, effectiveBaseUri, sourceDirectoryUri);
                if (!string.Equals(resolved, item.Value, StringComparison.Ordinal)) {
                    resolvedExternalManifestReferences.Add(item, resolved);
                }
            }
        }

    }

    private static string ResolveExternalManifestReference(
        string reference,
        Uri? effectiveBaseUri,
        Uri? sourceDirectoryUri) {
        if (effectiveBaseUri is null ||
            !Uri.TryCreate(effectiveBaseUri, reference, out Uri? resolved)) return reference;
        if (resolved.IsFile && sourceDirectoryUri?.IsFile == true) {
            return sourceDirectoryUri.MakeRelativeUri(resolved).ToString();
        }
        return resolved.IsAbsoluteUri ? resolved.AbsoluteUri : reference;
    }

    /// <summary>Inspects a bounded HTML file without resolving external resources.</summary>
    public static OfficeProvenanceReport InspectFile(string filePath, OfficeProvenanceOptions? options = null) {
        return InspectFile(filePath, filePath, options);
    }

    internal static OfficeProvenanceReport InspectFile(
        string filePath,
        string logicalFilePath,
        OfficeProvenanceOptions? options) {
        if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("A file path is required.", nameof(filePath));
        if (string.IsNullOrWhiteSpace(logicalFilePath)) throw new ArgumentException("A logical file path is required.", nameof(logicalFilePath));
        options ??= new OfficeProvenanceOptions();
        OfficeProvenanceBinary.ValidateLimits(options);
        byte[] data = ReadBounded(filePath, options.MaxAssetBytes, options.CancellationToken);
        return InspectCore(
            DecodeHtml(data, out _, out _),
            options,
            enforceUtf8Size: false,
            new Uri(Path.GetFullPath(logicalFilePath)));
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
        OfficeProvenanceBinary.ValidateRemovalOptions(options);
        options.Limits.CancellationToken.ThrowIfCancellationRequested();
        OfficeProvenanceOptions inspectionOptions = CreateInspectionOptions(options);
        OfficeProvenanceReport before = InspectCore(html, inspectionOptions, enforceUtf8Size);
        int structuralEntries = 0;
        IHtmlDocument document = ParseBoundedDocument(html, options.Limits.MaxContainerEntries, ref structuralEntries, options.Limits.CancellationToken);
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
                OfficeProvenanceBinary.EnsureOutputWithinLimit(byteCount, options.EffectiveMaxOutputBytes);
                original = Encoding.UTF8.GetBytes(html);
            }
            return new OfficeProvenanceRemovalResult(original, before, before, changes.AsReadOnly(), false);
        }

        if (enforceUtf8Size) NormalizeDeclaredEncodingToUtf8(document);
        string outputHtml = document.ToHtml();
        byte[] output;
        if (outputEncoding == null) {
            int byteCount = Encoding.UTF8.GetByteCount(outputHtml);
            OfficeProvenanceBinary.EnsureOutputWithinLimit(byteCount, options.EffectiveMaxOutputBytes);
            output = Encoding.UTF8.GetBytes(outputHtml);
        } else {
            output = EncodeHtml(outputHtml, outputEncoding, outputHadPreamble, options.EffectiveMaxOutputBytes);
        }
        OfficeProvenanceReport after = InspectCore(
            outputHtml,
            CreateOutputInspectionOptions(options),
            enforceUtf8Size);
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
            if (!string.Equals(TrimAsciiWhitespace(script.GetAttribute("type")), "application/c2pa", StringComparison.OrdinalIgnoreCase)) continue;
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
            string value = TrimAsciiWhitespace(link.GetAttribute("href"));
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
        byte[] input = ReadBounded(inputPath, options.Limits.MaxAssetBytes, options.Limits.CancellationToken);
        string html = DecodeHtml(input, out Encoding encoding, out bool hadPreamble);
        OfficeProvenanceRemovalResult result = RemoveCore(
            html, options, enforceUtf8Size: false, outputEncoding: encoding, outputHadPreamble: hadPreamble);
        if (!result.WasChanged) {
            OfficeProvenanceBinary.EnsureOutputWithinLimit(input.LongLength, options.EffectiveMaxOutputBytes);
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
        Dictionary<OfficeProvenanceEvidence, string> resolvedExternalManifestReferences,
        Uri? effectiveBaseUri,
        Uri? sourceDirectoryUri,
        string documentLocation,
        int srcDocDepth) {
        if (options.ProcessEmbeddedAssets) {
            IElement[] elements = GetEmbeddedImageElements(document, options).ToArray();
            HtmlProvenanceCssScope cssScope = HtmlResourcePipeline.CollectProvenanceCssImageScope(
                document,
                options.MaxAssetBytes,
                options.MaxExpandedContainerBytes - expandedBytes);
            ReserveExpandedBytes(ref expandedBytes, cssScope.DecodedStylesheetBytes, options.MaxExpandedContainerBytes);
            foreach (IElement element in elements) {
                options.CancellationToken.ThrowIfCancellationRequested();
                cssScope.UsedCustomPropertyDeclarations.TryGetValue(element, out HashSet<int>? usedDeclarations);
                cssScope.ResolvedVarFallbackStarts.TryGetValue(element, out HashSet<int>? resolvedFallbacks);
                cssScope.DataStylesheets.TryGetValue(element, out HtmlProvenanceDataStylesheet? dataStylesheet);
                foreach (EmbeddedImageReference reference in GetEmbeddedImageReferences(
                    document, element, usedDeclarations, resolvedFallbacks, cssScope.ComputedStyles,
                    cssScope.ComputedStyleSet, dataStylesheet)) {
                    options.CancellationToken.ThrowIfCancellationRequested();
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
                    } catch (Exception exception) when (
                        (exception is InvalidDataException || exception is System.Xml.XmlException) &&
                        !OfficeProvenanceLimitException.Is(exception)) {
                        diagnostics.Add($"{location}: embedded image was preserved because inspection failed: {exception.Message}");
                    }
                }
            }
        }
        if (srcDocDepth >= HtmlConversionInputGuard.MaxSrcDocDepth) {
            ThrowIfNestedSrcDocRemains(document);
            return;
        }
        int iframeIndex = 0;
        foreach (IElement iframe in document.QuerySelectorAll("iframe[srcdoc]").Where(IsHtmlIframe)) {
            options.CancellationToken.ThrowIfCancellationRequested();
            string? srcdoc = iframe.GetAttribute("srcdoc");
            if (srcdoc == null || string.IsNullOrWhiteSpace(srcdoc)) continue;
            string location = $"{documentLocation}/iframe[srcdoc][{iframeIndex++}]";
            IHtmlDocument nested = ParseBoundedDocument(srcdoc, options.MaxContainerEntries, ref structuralEntries, options.CancellationToken);
            Uri? nestedBaseUri = HtmlDocumentParser.ResolveEffectiveBaseUri(nested, effectiveBaseUri);
            InspectManifestCarriers(
                nested,
                options,
                evidence,
                diagnostics,
                resolvedExternalManifestReferences,
                nestedBaseUri,
                sourceDirectoryUri,
                location,
                ref expandedBytes);
            InspectEmbeddedImages(nested, options, evidence, diagnostics, ref count, ref structuralEntries,
                ref expandedBytes, resolvedExternalManifestReferences, nestedBaseUri, sourceDirectoryUri,
                location, srcDocDepth + 1);
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
            IElement[] elements = GetEmbeddedImageElements(document, options.Limits).ToArray();
            HtmlProvenanceCssScope cssScope = HtmlResourcePipeline.CollectProvenanceCssImageScope(
                document,
                options.Limits.MaxAssetBytes,
                options.Limits.MaxExpandedContainerBytes - expandedBytes);
            ReserveExpandedBytes(ref expandedBytes, cssScope.DecodedStylesheetBytes, options.Limits.MaxExpandedContainerBytes);
            foreach (IElement element in elements) {
                options.Limits.CancellationToken.ThrowIfCancellationRequested();
                cssScope.UsedCustomPropertyDeclarations.TryGetValue(element, out HashSet<int>? usedDeclarations);
                cssScope.ResolvedVarFallbackStarts.TryGetValue(element, out HashSet<int>? resolvedFallbacks);
                cssScope.DataStylesheets.TryGetValue(element, out HtmlProvenanceDataStylesheet? dataStylesheet);
                EmbeddedImageReference[] references = GetEmbeddedImageReferences(
                    document, element, usedDeclarations, resolvedFallbacks, cssScope.ComputedStyles,
                    cssScope.ComputedStyleSet, dataStylesheet).ToArray();
                var replacements = new List<(EmbeddedImageReference Reference, string Value)>();
                foreach (EmbeddedImageReference reference in references) {
                    options.Limits.CancellationToken.ThrowIfCancellationRequested();
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
                        byte[] nestedOutput = nested.ToArray();
                        long base64Characters = checked(((nestedOutput.LongLength + 2L) / 3L) * 4L);
                        long replacementCharacters = checked(
                            5L + metadata.Length + base64Characters + dataUri.Fragment.Length);
                        OfficeProvenanceBinary.EnsureOutputWithinLimit(
                            replacementCharacters,
                            options.EffectiveMaxOutputBytes);
                        replacements.Add((reference,
                            "data:" + metadata + "," + Convert.ToBase64String(nestedOutput) + dataUri.Fragment));
                        foreach (OfficeProvenanceChange change in nested.Changes) {
                            changes.Add(new OfficeProvenanceChange(
                                change.Carrier,
                                $"{documentLocation}/{element.LocalName}[{reference.AttributeName}][{index}]/{change.Location}",
                                0));
                        }
                    } catch (Exception exception) when (
                        (exception is InvalidDataException || exception is System.Xml.XmlException) &&
                        !OfficeProvenanceLimitException.Is(exception)) {
                        // Preserve malformed embedded data; structural diagnostics are available through Inspect.
                    }
                }
                ApplyEmbeddedImageReplacements(element, replacements);
            }
        }
        if (srcDocDepth >= HtmlConversionInputGuard.MaxSrcDocDepth) {
            ThrowIfNestedSrcDocRemains(document);
            return;
        }
        int iframeIndex = 0;
        foreach (IElement iframe in document.QuerySelectorAll("iframe[srcdoc]").Where(IsHtmlIframe)) {
            options.Limits.CancellationToken.ThrowIfCancellationRequested();
            string? srcdoc = iframe.GetAttribute("srcdoc");
            if (srcdoc == null || string.IsNullOrWhiteSpace(srcdoc)) continue;
            string location = $"{documentLocation}/iframe[srcdoc][{iframeIndex++}]";
            IHtmlDocument nested = ParseBoundedDocument(srcdoc, options.Limits.MaxContainerEntries, ref structuralEntries, options.Limits.CancellationToken);
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
        ISet<int>? resolvedVarFallbacks,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> computedStyles,
        HtmlComputedStyleSet computedStyleSet,
        HtmlProvenanceDataStylesheet? dataStylesheet) {
        const string htmlNamespace = "http://www.w3.org/1999/xhtml";
        const string svgNamespace = "http://www.w3.org/2000/svg";
        string localName = element.LocalName.ToLowerInvariant();
        bool isHtmlElement = string.Equals(element.NamespaceUri, htmlNamespace, StringComparison.Ordinal);
        bool isSvgElement = string.Equals(element.NamespaceUri, svgNamespace, StringComparison.Ordinal);
        if (dataStylesheet != null) {
            foreach (HtmlCssImageReference reference in HtmlResourcePipeline.EnumerateProvenanceCssImageReferences(
                document,
                "css",
                dataStylesheet.Css,
                usedImageProperties,
                resolvedVarFallbacks,
                computedStyles,
                computedStyleSet,
                element)) {
                yield return new EmbeddedImageReference(
                    "stylesheet-css", reference.Value, reference.Start, reference.Length,
                    dataStylesheet.Css, dataStylesheet.Metadata, dataStylesheet.Fragment);
            }
            yield break;
        }
        if (localName == "style") {
            if (!HtmlResourcePipeline.IsActiveProvenanceStyleElement(element)) yield break;
            foreach (HtmlCssImageReference reference in HtmlResourcePipeline.EnumerateProvenanceCssImageReferences(
                document,
                "css",
                element.TextContent,
                usedImageProperties,
                resolvedVarFallbacks,
                computedStyles,
                computedStyleSet,
                element)) {
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
                resolvedVarFallbacks,
                computedStyles,
                computedStyleSet,
                element)) {
                yield return new EmbeddedImageReference("style", reference.Value, reference.Start, reference.Length);
            }
        }

        string? background = element.GetAttribute("background");
        if (background != null && isHtmlElement && HtmlResourcePipeline.SupportsLegacyBackground(localName)) {
            yield return CreateDirectUrlReference("background", background);
        }

        if (localName is "img" or "source" or "video" or "input" or "link" or "object" or "embed" && !isHtmlElement) yield break;
        if (localName is "image" or "feimage" or "use" && !isSvgElement) yield break;
        if (localName == "source" && !IsSupportedPictureSource(element)) yield break;
        if (localName == "img" && !HtmlResourcePipeline.IsActivePictureFallbackImage(element)) yield break;
        if (localName is "img" or "source") {
            if (localName == "img") {
                foreach (string attributeName in EmbeddedImageSourceAttributes) {
                    string? source = element.GetAttribute(attributeName);
                    if (source != null && !SourceSetReplacesDirectSource(element, attributeName)) {
                        yield return CreateDirectUrlReference(attributeName, source);
                    }
                }
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
            attributeNames = HtmlDocumentParser.GetExactAttributeValue(element, "href") != null ? new[] { "href" } : new[] { "xlink:href" };
        } else if (localName is "feimage" or "use") {
            attributeNames = HtmlDocumentParser.GetExactAttributeValue(element, "href") != null ? new[] { "href" } : new[] { "xlink:href" };
        } else if (localName == "link" && IsImageLink(element)) {
            if (IsPreloadedImage(element) && HasUsableImageSourceSet(element)) attributeNames = Array.Empty<string>();
            else attributeNames = new[] { "href" };
        } else if (localName == "object" && HasSupportedDeclaredOrInferredImageType(element, "data")) {
            attributeNames = new[] { "data" };
        } else if (localName == "embed" && HasSupportedDeclaredOrInferredImageType(element, "src")) {
            attributeNames = new[] { "src" };
        } else {
            yield break;
        }

        foreach (string attributeName in attributeNames) {
            string? source = attributeName is "href" or "xlink:href"
                ? HtmlDocumentParser.GetExactAttributeValue(element, attributeName)
                : element.GetAttribute(attributeName);
            if (source != null) yield return CreateDirectUrlReference(attributeName, source);
        }
        if (localName == "link" && IsPreloadedImage(element)) {
            string? sourceSet = element.GetAttribute("imagesrcset");
            if (sourceSet != null) {
                foreach (EmbeddedImageReference reference in ParseSrcset("imagesrcset", sourceSet)) yield return reference;
            }
        }
    }

    private static void ThrowIfNestedSrcDocRemains(IHtmlDocument document) {
        if (document.QuerySelectorAll("iframe[srcdoc]").Where(IsHtmlIframe).Any(iframe =>
            !string.IsNullOrWhiteSpace(iframe.GetAttribute("srcdoc")))) {
            throw new InvalidDataException($"HTML iframe srcdoc nesting exceeds the supported depth of {HtmlConversionInputGuard.MaxSrcDocDepth}.");
        }
    }

    private static bool IsHtmlIframe(IElement element) =>
        string.Equals(element.NamespaceUri, "http://www.w3.org/1999/xhtml", StringComparison.Ordinal);

    private static IEnumerable<IElement> GetEmbeddedImageElements(
        IHtmlDocument document,
        OfficeProvenanceOptions options) {
        var activeObjectImages = new Dictionary<IElement, bool>();
        return document.QuerySelectorAll("img,source,video,input,object,embed,image,feImage,use,link,[background],style,[style]")
            .Where(element => !element.LocalName.Equals("style", StringComparison.OrdinalIgnoreCase) ||
                HtmlResourcePipeline.IsActiveProvenanceStyleElement(element))
            .Where(element => !element.LocalName.Equals("source", StringComparison.OrdinalIgnoreCase) ||
                IsSupportedPictureSource(element))
            .Where(element => !HasActiveObjectImageAncestor(element, options, activeObjectImages))
            .Distinct();
    }

    private static bool HasActiveObjectImageAncestor(
        IElement element,
        OfficeProvenanceOptions options,
        IDictionary<IElement, bool> cache) {
        for (IElement? ancestor = element.ParentElement; ancestor != null; ancestor = ancestor.ParentElement) {
            if (!ancestor.LocalName.Equals("object", StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(ancestor.NamespaceUri, "http://www.w3.org/1999/xhtml", StringComparison.Ordinal)) continue;
            if (!cache.TryGetValue(ancestor, out bool active)) {
                active = HasLoadableSupportedObjectImage(ancestor, options);
                cache.Add(ancestor, active);
            }
            if (active) return true;
        }
        return false;
    }

    private static bool HasLoadableSupportedObjectImage(IElement element, OfficeProvenanceOptions options) {
        if (!HasSupportedDeclaredOrInferredImageType(element, "data")) return false;
        string? source = element.GetAttribute("data");
        if (source == null || !HtmlImageDataUri.TryParse(source, out HtmlImageDataUri dataUri) ||
            !IsSupportedProvenanceImage(dataUri.MediaType) ||
            !dataUri.TryEstimateDecodedByteCount(out long estimatedBytes) ||
            estimatedBytes > options.MaxAssetBytes) return false;
        try {
            if (!TryDecodeEmbeddedImage(dataUri, options.MaxAssetBytes, out byte[] image) ||
                image.LongLength > options.MaxAssetBytes) return false;
            OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(
                image,
                "asset" + dataUri.FileExtension,
                CreateNestedOptions(options));
            return IsMatchingImageFormat(dataUri.MediaType, report.Format);
        } catch (Exception exception) when (
            (exception is InvalidDataException || exception is System.Xml.XmlException) &&
            !OfficeProvenanceLimitException.Is(exception)) {
            return false;
        }
    }

    private static bool IsMatchingImageFormat(string mediaType, OfficeProvenanceAssetFormat format) {
        switch (mediaType.ToLowerInvariant()) {
            case "image/jpeg":
            case "image/jpg": return format == OfficeProvenanceAssetFormat.Jpeg;
            case "image/png": return format == OfficeProvenanceAssetFormat.Png;
            case "image/gif": return format == OfficeProvenanceAssetFormat.Gif;
            case "image/tiff": return format == OfficeProvenanceAssetFormat.Tiff;
            case "image/webp": return format == OfficeProvenanceAssetFormat.Webp;
            case "image/svg+xml": return format == OfficeProvenanceAssetFormat.Svg;
            default: return false;
        }
    }

    private static bool IsSupportedPictureSource(IElement element) {
        string? parentName = element.ParentElement?.LocalName;
        return string.Equals(parentName, "picture", StringComparison.OrdinalIgnoreCase) &&
            HtmlResourcePipeline.IsActivePictureImageSource(element);
    }

    private static void NormalizeDeclaredEncodingToUtf8(IHtmlDocument document) {
        foreach (IElement meta in document.QuerySelectorAll("meta[charset]")) meta.SetAttribute("charset", "utf-8");
        foreach (IElement meta in document.QuerySelectorAll("meta[http-equiv][content]")) {
            if (!string.Equals(TrimAsciiWhitespace(meta.GetAttribute("http-equiv")), "content-type", StringComparison.OrdinalIgnoreCase)) continue;
            string content = meta.GetAttribute("content") ?? string.Empty;
            meta.SetAttribute("content", RewriteExactCharsetParameter(content));
        }
    }

    private static string RewriteExactCharsetParameter(string content) {
        int segmentStart = 0;
        char quote = '\0';
        for (int index = 0; index <= content.Length; index++) {
            char current = index < content.Length ? content[index] : ';';
            if (quote != '\0') {
                if (current == quote) quote = '\0';
                continue;
            }
            if (current is '\'' or '"') {
                quote = current;
                continue;
            }
            if (current != ';') continue;
            int equals = content.IndexOf('=', segmentStart, index - segmentStart);
            if (equals >= 0 && string.Equals(
                    TrimAsciiWhitespace(content.Substring(segmentStart, equals - segmentStart)),
                    "charset",
                    StringComparison.OrdinalIgnoreCase)) {
                int valueStart = equals + 1;
                while (valueStart < index && IsAsciiWhitespace(content[valueStart])) valueStart++;
                int valueEnd = index;
                while (valueEnd > valueStart && IsAsciiWhitespace(content[valueEnd - 1])) valueEnd--;
                string replacement = "utf-8";
                if (valueEnd - valueStart >= 2 && content[valueStart] is '\'' or '"' &&
                    content[valueEnd - 1] == content[valueStart]) {
                    replacement = content[valueStart] + replacement + content[valueEnd - 1];
                }
                return content.Substring(0, valueStart) + replacement + content.Substring(valueEnd);
            }
            segmentStart = index + 1;
        }
        return content;
    }

    private static EmbeddedImageReference CreateDirectUrlReference(string attributeName, string value) {
        int start = 0;
        while (start < value.Length && IsAsciiWhitespace(value[start])) start++;
        int end = value.Length;
        while (end > start && IsAsciiWhitespace(value[end - 1])) end--;
        string parsedValue = value.Substring(start, end - start)
            .Replace("\t", string.Empty)
            .Replace("\n", string.Empty)
            .Replace("\r", string.Empty);
        return new EmbeddedImageReference(attributeName, parsedValue, start, end - start);
    }

    private static bool IsImageLink(IElement element) {
        string? rel = element.GetAttribute("rel");
        return HtmlResourcePipeline.IsApplicableProvenanceMedia(element) &&
            (HasRelationship(rel, "icon") || HasRelationship(rel, "apple-touch-icon") ||
            HasRelationship(rel, "shortcut icon") ||
            HasRelationship(rel, "shortcut") && HasRelationship(rel, "icon") || IsPreloadedImage(element));
    }

    private static bool SourceSetReplacesDirectSource(IElement element, string sourceAttribute) {
        string? sourceSetAttribute = sourceAttribute switch {
            "src" => "srcset",
            "data-src" => "data-srcset",
            "data-original" or "data-original-src" => "data-original-srcset",
            "data-lazy-src" => "data-lazy-srcset",
            _ => null
        };
        if (sourceSetAttribute == null) return false;
        string? sourceSet = element.GetAttribute(sourceSetAttribute);
        if (string.IsNullOrEmpty(sourceSet)) return false;
        foreach (HtmlSrcSetCandidate candidate in HtmlSrcSetParser.Enumerate(sourceSet)) {
            string descriptor = candidate.Descriptor.Trim();
            if (descriptor.Length == 0 || descriptor.EndsWith("w", StringComparison.Ordinal)) return true;
            if (descriptor.EndsWith("x", StringComparison.Ordinal) && double.TryParse(
                    descriptor.Substring(0, descriptor.Length - 1),
                    System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out double density) && density == 1D) return true;
        }
        return false;
    }

    private static bool IsPreloadedImage(IElement element) =>
        HasRelationship(element.GetAttribute("rel"), "preload") &&
        string.Equals(TrimAsciiWhitespace(element.GetAttribute("as")), "image", StringComparison.OrdinalIgnoreCase) &&
        HtmlResourcePipeline.IsApplicableProvenanceMedia(element);

    private static bool HasUsableImageSourceSet(IElement element) {
        string? sourceSet = element.GetAttribute("imagesrcset");
        return !string.IsNullOrEmpty(sourceSet) && HtmlSrcSetParser.Enumerate(sourceSet).Any();
    }

    private static bool HasSupportedDeclaredImageType(IElement element) {
        string value = TrimAsciiWhitespace(element.GetAttribute("type"));
        int parameter = value.IndexOf(';');
        if (parameter >= 0) value = TrimAsciiWhitespace(value.Substring(0, parameter));
        return IsSupportedProvenanceImage(value);
    }

    private static bool HasSupportedDeclaredOrInferredImageType(IElement element, string sourceAttribute) {
        string declaredType = TrimAsciiWhitespace(element.GetAttribute("type"));
        if (declaredType.Length != 0) return HasSupportedDeclaredImageType(element);
        string? source = element.GetAttribute(sourceAttribute);
        return source != null && HtmlImageDataUri.TryParse(source, out HtmlImageDataUri dataUri) &&
               IsSupportedProvenanceImage(dataUri.MediaType);
    }

    private static IEnumerable<EmbeddedImageReference> ParseSrcset(string attributeName, string sourceSet) {
        foreach (HtmlSrcSetCandidate candidate in HtmlSrcSetParser.Enumerate(sourceSet)) {
            int start = candidate.UrlStart;
            if (start < 0 || start > sourceSet.Length - candidate.Url.Length) continue;
            yield return new EmbeddedImageReference(attributeName, candidate.Url, start, candidate.Url.Length);
        }
    }

    private static string CreateRewrittenDataUriMetadata(HtmlImageDataUri dataUri) {
        string[] parts = dataUri.Metadata.Split(';');
        var metadata = new List<string>(parts.Length + 1);
        bool svg = string.Equals(dataUri.MediaType, "image/svg+xml", StringComparison.OrdinalIgnoreCase);
        bool hasCharset = false;
        foreach (string part in parts) {
            string trimmed = TrimAsciiWhitespace(part);
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
            .Any(part => TrimAsciiWhitespace(part).StartsWith("charset=", StringComparison.OrdinalIgnoreCase));
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
            IAttr? exactAttribute = group.Key is "href" or "xlink:href" ? HtmlDocumentParser.GetExactAttribute(element, group.Key) : null;
            string value = group.Key == "css"
                ? element.TextContent
                : group.Key == "stylesheet-css"
                    ? group.First().Reference.ContainerText ?? string.Empty
                    : exactAttribute?.Value ?? element.GetAttribute(group.Key) ?? string.Empty;
            foreach ((EmbeddedImageReference reference, string replacement) in group.OrderByDescending(item => item.Reference.Start)) {
                value = value.Substring(0, reference.Start) + replacement + value.Substring(reference.Start + reference.Length);
            }
            if (group.Key == "css") element.TextContent = value;
            else if (group.Key == "stylesheet-css") {
                string metadata = CreateRewrittenCssDataUriMetadata(group.First().Reference.ContainerMetadata ?? "text/css");
                element.SetAttribute(
                    "href",
                    "data:" + metadata + "," + Convert.ToBase64String(Encoding.UTF8.GetBytes(value)) +
                    (group.First().Reference.ContainerFragment ?? string.Empty));
            }
            else if (exactAttribute != null) exactAttribute.Value = value;
            else element.SetAttribute(group.Key, value);
        }
    }

    private static string CreateRewrittenCssDataUriMetadata(string originalMetadata) {
        var metadata = new List<string>();
        foreach (string part in originalMetadata.Split(';')) {
            string trimmed = TrimAsciiWhitespace(part);
            if (trimmed.Length == 0 || trimmed.Equals("base64", StringComparison.OrdinalIgnoreCase) ||
                trimmed.StartsWith("charset=", StringComparison.OrdinalIgnoreCase)) continue;
            metadata.Add(trimmed);
        }
        if (metadata.Count == 0) metadata.Add("text/css");
        metadata.Add("charset=utf-8");
        metadata.Add("base64");
        return string.Join(";", metadata);
    }

    private static bool TryDecodeManifest(
        string? value,
        long maximumBytes,
        long maximumExpandedBytes,
        ref long expandedBytes,
        out byte[] manifest) {
        manifest = Array.Empty<byte>();
        string encoded = TrimAsciiWhitespace(value);
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
        SplitAsciiWhitespace(value)
            .Any(item => item.Equals(relationship, StringComparison.OrdinalIgnoreCase));

    private static void RemoveRelationship(IElement element, string relationship) {
        string[] retained = SplitAsciiWhitespace(element.GetAttribute("rel"))
            .Where(item => !item.Equals(relationship, StringComparison.OrdinalIgnoreCase))
            .ToArray();
        if (retained.Length == 0) element.Remove();
        else element.SetAttribute("rel", string.Join(" ", retained));
    }

    private static string[] SplitAsciiWhitespace(string? value) =>
        (value ?? string.Empty).Split(new[] { '\t', '\n', '\f', '\r', ' ' }, StringSplitOptions.RemoveEmptyEntries);

    private static string TrimAsciiWhitespace(string? value) {
        if (string.IsNullOrEmpty(value)) return string.Empty;
        int start = 0;
        int end = value!.Length;
        while (start < end && IsAsciiWhitespace(value[start])) start++;
        while (end > start && IsAsciiWhitespace(value[end - 1])) end--;
        return start == 0 && end == value.Length ? value : value.Substring(start, end - start);
    }

    private static bool IsManifestElement(IElement element) =>
        element.LocalName.Equals("script", StringComparison.OrdinalIgnoreCase)
            ? string.Equals(TrimAsciiWhitespace(element.GetAttribute("type")), "application/c2pa", StringComparison.OrdinalIgnoreCase)
            : element.LocalName.Equals("link", StringComparison.OrdinalIgnoreCase) &&
                HasRelationship(element.GetAttribute("rel"), "c2pa-manifest");

    private static bool IsSafeManifestReference(string value, out Uri? uri) {
        uri = null;
        string preprocessed = PreprocessHtmlUrl(value);
        if (preprocessed.Length == 0 || !Uri.TryCreate(preprocessed, UriKind.RelativeOrAbsolute, out Uri? parsed)) return false;
        if (parsed.IsAbsoluteUri && parsed.Scheme != Uri.UriSchemeHttp && parsed.Scheme != Uri.UriSchemeHttps) return false;
        uri = parsed;
        return true;
    }

    private static string PreprocessHtmlUrl(string? value) =>
        TrimAsciiWhitespace(value)
            .Replace("\t", string.Empty)
            .Replace("\n", string.Empty)
            .Replace("\r", string.Empty);

    private static IHtmlDocument ParseBoundedDocument(
        string html,
        int maximumEntries,
        ref int structuralEntries,
        CancellationToken cancellationToken) {
        int remaining = maximumEntries - structuralEntries;
        if (remaining <= 0) throw new InvalidDataException("The HTML document exceeds the configured container-entry limit.");
        ValidatePotentialElementCountCore(html, remaining, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        IHtmlDocument document = HtmlDocumentParser.ParseDocument(html);
        int elementCount = document.All.Length;
        if (elementCount > remaining) throw new InvalidDataException("The HTML document exceeds the configured container-entry limit.");
        structuralEntries += elementCount;
        return document;
    }

    private static void ValidatePotentialElementCount(string html, int maximumEntries) =>
        ValidatePotentialElementCountCore(html, maximumEntries, CancellationToken.None);

    private static void ValidatePotentialElementCountCore(
        string html,
        int maximumEntries,
        CancellationToken cancellationToken) {
        int count = 0;
        int index = 0;
        bool sawHtmlElement = false;
        bool sawHeadElement = false;
        bool sawBodyElement = false;
        var openElements = new List<HtmlPreflightElement>();
        while (index < html.Length - 1) {
            cancellationToken.ThrowIfCancellationRequested();
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
                    int declarationEnd = FindDoctypeEnd(html, markup + 9);
                    index = declarationEnd < 0 ? html.Length : declarationEnd + 1;
                } else {
                    int declarationEnd = html.IndexOf('>', markup + 2);
                    index = declarationEnd < 0 ? html.Length : declarationEnd + 1;
                }
                continue;
            }
            if (next == '/') {
                int nameStart = markup + 2;
                int closingNameEnd = FindHtmlTagNameEnd(html, nameStart);
                string closingName = html.Substring(nameStart, closingNameEnd - nameStart);
                if (!HandleEndTagInSelect(openElements, closingName)) {
                    for (int elementIndex = openElements.Count - 1; elementIndex >= 0; elementIndex--) {
                        if (!openElements[elementIndex].Name.Equals(closingName, StringComparison.OrdinalIgnoreCase)) continue;
                        openElements.RemoveRange(elementIndex, openElements.Count - elementIndex);
                        break;
                    }
                }
                int declarationEnd = FindStartTagEnd(html, closingNameEnd, out _);
                index = declarationEnd < 0 ? html.Length : declarationEnd + 1;
                continue;
            }
            if (!IsAsciiLetter(next)) {
                index = markup + 1;
                continue;
            }
            int nameEnd = FindHtmlTagNameEnd(html, markup + 2);
            string tagName = html.Substring(markup + 1, nameEnd - markup - 1);
            int tagEnd = FindStartTagEnd(html, nameEnd, out bool selfClosing);
            if (tagEnd < 0) break;
            if (ChildNamespace(openElements) != HtmlPreflightNamespace.Html &&
                IsForeignContentHtmlBreakout(html, tagName, nameEnd, tagEnd)) {
                while (openElements.Count > 0 && ChildNamespace(openElements) != HtmlPreflightNamespace.Html) {
                    openElements.RemoveAt(openElements.Count - 1);
                }
            }
            if (ShouldIgnoreStartTagInSelect(openElements, tagName)) {
                index = tagEnd + 1;
                continue;
            }
            HtmlPreflightNamespace elementNamespace = ChildNamespace(openElements, tagName);
            if (elementNamespace == HtmlPreflightNamespace.Html &&
                ShouldIgnoreNestedFormStart(openElements, tagName)) {
                index = tagEnd + 1;
                continue;
            }
            if (elementNamespace == HtmlPreflightNamespace.Html &&
                ShouldIgnoreDuplicateDocumentElement(
                    tagName,
                    ref sawHtmlElement,
                    ref sawHeadElement,
                    ref sawBodyElement)) {
                index = tagEnd + 1;
                continue;
            }
            if (++count > maximumEntries) {
                throw new InvalidDataException("The HTML document exceeds the configured container-entry limit.");
            }
            bool childrenUseHtml = elementNamespace == HtmlPreflightNamespace.Html ||
                IsHtmlIntegrationPoint(html, tagName, elementNamespace, nameEnd, tagEnd);
            if (!selfClosing && !(elementNamespace == HtmlPreflightNamespace.Html && IsHtmlVoidElement(tagName))) {
                openElements.Add(new HtmlPreflightElement(tagName, elementNamespace, childrenUseHtml));
            }
            index = tagEnd + 1;
            if (elementNamespace == HtmlPreflightNamespace.Html && tagName.Equals("plaintext", StringComparison.OrdinalIgnoreCase)) return;
            if (elementNamespace == HtmlPreflightNamespace.Html && IsRawTextOrRcDataElement(tagName)) {
                int rawTextEnd = HtmlRawTextScanner.FindClosingTag(html, index, tagName);
                if (rawTextEnd < 0) break;
                index = rawTextEnd;
            }
        }
    }

    private static bool ShouldIgnoreDuplicateDocumentElement(
        string tagName,
        ref bool sawHtmlElement,
        ref bool sawHeadElement,
        ref bool sawBodyElement) {
        if (tagName.Equals("html", StringComparison.OrdinalIgnoreCase)) {
            if (sawHtmlElement) return true;
            sawHtmlElement = true;
            return false;
        }
        if (tagName.Equals("head", StringComparison.OrdinalIgnoreCase)) {
            if (sawHeadElement || sawBodyElement) return true;
            sawHeadElement = true;
            return false;
        }
        if (!tagName.Equals("body", StringComparison.OrdinalIgnoreCase)) return false;
        if (sawBodyElement) return true;
        sawBodyElement = true;
        return false;
    }

    private static bool ShouldIgnoreNestedFormStart(List<HtmlPreflightElement> elements, string tagName) {
        if (!tagName.Equals("form", StringComparison.OrdinalIgnoreCase)) return false;
        bool hasForm = false;
        foreach (HtmlPreflightElement element in elements) {
            if (element.Namespace != HtmlPreflightNamespace.Html) continue;
            if (element.Name.Equals("template", StringComparison.OrdinalIgnoreCase)) return false;
            if (element.Name.Equals("form", StringComparison.OrdinalIgnoreCase)) hasForm = true;
        }
        return hasForm;
    }

    private static bool ShouldIgnoreStartTagInSelect(List<HtmlPreflightElement> elements, string tagName) {
        int selectIndex = -1;
        for (int index = elements.Count - 1; index >= 0; index--) {
            HtmlPreflightElement element = elements[index];
            if (element.Namespace != HtmlPreflightNamespace.Html) continue;
            if (element.Name.Equals("template", StringComparison.OrdinalIgnoreCase)) return false;
            if (!element.Name.Equals("select", StringComparison.OrdinalIgnoreCase)) continue;
            selectIndex = index;
            break;
        }
        if (selectIndex < 0) return false;

        if (tagName.Equals("option", StringComparison.OrdinalIgnoreCase)) {
            RemoveCurrentSelectChild(elements, "option");
            return false;
        }
        if (tagName.Equals("optgroup", StringComparison.OrdinalIgnoreCase)) {
            RemoveCurrentSelectChild(elements, "option");
            RemoveCurrentSelectChild(elements, "optgroup");
            return false;
        }
        if (tagName.Equals("hr", StringComparison.OrdinalIgnoreCase)) {
            RemoveCurrentSelectChild(elements, "option");
            RemoveCurrentSelectChild(elements, "optgroup");
            return false;
        }
        if (tagName.Equals("script", StringComparison.OrdinalIgnoreCase) ||
            tagName.Equals("template", StringComparison.OrdinalIgnoreCase)) return false;
        if (tagName.Equals("input", StringComparison.OrdinalIgnoreCase) ||
            tagName.Equals("keygen", StringComparison.OrdinalIgnoreCase) ||
            tagName.Equals("textarea", StringComparison.OrdinalIgnoreCase)) {
            elements.RemoveRange(selectIndex, elements.Count - selectIndex);
            return false;
        }
        if (tagName.Equals("select", StringComparison.OrdinalIgnoreCase)) {
            elements.RemoveRange(selectIndex, elements.Count - selectIndex);
        }
        return true;
    }

    private static bool HandleEndTagInSelect(List<HtmlPreflightElement> elements, string tagName) {
        int selectIndex = -1;
        for (int index = elements.Count - 1; index >= 0; index--) {
            HtmlPreflightElement element = elements[index];
            if (element.Namespace != HtmlPreflightNamespace.Html) continue;
            if (element.Name.Equals("template", StringComparison.OrdinalIgnoreCase)) return false;
            if (!element.Name.Equals("select", StringComparison.OrdinalIgnoreCase)) continue;
            selectIndex = index;
            break;
        }
        if (selectIndex < 0) return false;

        if (tagName.Equals("option", StringComparison.OrdinalIgnoreCase)) {
            RemoveCurrentSelectChild(elements, "option");
            return true;
        }
        if (tagName.Equals("optgroup", StringComparison.OrdinalIgnoreCase)) {
            if (elements.Count >= 2 &&
                elements[elements.Count - 1].Name.Equals("option", StringComparison.OrdinalIgnoreCase) &&
                elements[elements.Count - 2].Name.Equals("optgroup", StringComparison.OrdinalIgnoreCase)) {
                elements.RemoveAt(elements.Count - 1);
            }
            RemoveCurrentSelectChild(elements, "optgroup");
            return true;
        }
        if (tagName.Equals("select", StringComparison.OrdinalIgnoreCase)) {
            elements.RemoveRange(selectIndex, elements.Count - selectIndex);
            return true;
        }
        if (tagName.Equals("template", StringComparison.OrdinalIgnoreCase)) return false;
        return true;
    }

    private static void RemoveCurrentSelectChild(List<HtmlPreflightElement> elements, string name) {
        if (elements.Count != 0 && elements[elements.Count - 1].Name.Equals(name, StringComparison.OrdinalIgnoreCase)) {
            elements.RemoveAt(elements.Count - 1);
        }
    }

    private static bool IsHtmlVoidElement(string tagName) => tagName.Equals("area", StringComparison.OrdinalIgnoreCase) ||
        tagName.Equals("base", StringComparison.OrdinalIgnoreCase) ||
        tagName.Equals("br", StringComparison.OrdinalIgnoreCase) ||
        tagName.Equals("col", StringComparison.OrdinalIgnoreCase) ||
        tagName.Equals("embed", StringComparison.OrdinalIgnoreCase) ||
        tagName.Equals("hr", StringComparison.OrdinalIgnoreCase) ||
        tagName.Equals("img", StringComparison.OrdinalIgnoreCase) ||
        tagName.Equals("input", StringComparison.OrdinalIgnoreCase) ||
        tagName.Equals("link", StringComparison.OrdinalIgnoreCase) ||
        tagName.Equals("meta", StringComparison.OrdinalIgnoreCase) ||
        tagName.Equals("source", StringComparison.OrdinalIgnoreCase) ||
        tagName.Equals("track", StringComparison.OrdinalIgnoreCase) ||
        tagName.Equals("wbr", StringComparison.OrdinalIgnoreCase);

    private static int FindHtmlTagNameEnd(string html, int start) {
        int index = start;
        while (index < html.Length) {
            char character = html[index];
            if (character == '>' || character == '/' || character is '\t' or '\n' or '\f' or '\r' or ' ') break;
            index++;
        }
        return index;
    }

    private static bool IsSupportedProvenanceImage(string mediaType) => mediaType.ToLowerInvariant() is
        "image/jpeg" or "image/jpg" or "image/png" or "image/gif" or "image/tiff" or "image/webp" or "image/svg+xml";

    private static HtmlPreflightNamespace ChildNamespace(List<HtmlPreflightElement> elements, string? tagName = null) {
        if (elements.Count > 0 && elements[elements.Count - 1].Namespace == HtmlPreflightNamespace.MathMl &&
            elements[elements.Count - 1].ChildrenUseHtml && tagName != null &&
            (tagName.Equals("mglyph", StringComparison.OrdinalIgnoreCase) ||
             tagName.Equals("malignmark", StringComparison.OrdinalIgnoreCase))) {
            return HtmlPreflightNamespace.MathMl;
        }
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
        string value = encoding.Success
            ? System.Net.WebUtility.HtmlDecode(encoding.Groups["value"].Value)
            : string.Empty;
        return value.Equals("text/html", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("application/xhtml+xml", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsForeignContentHtmlBreakout(string html, string tagName, int attributesStart, int tagEnd) {
        switch (tagName.ToLowerInvariant()) {
            case "b": case "big": case "blockquote": case "body": case "br": case "center": case "code":
            case "dd": case "div": case "dl": case "dt": case "em": case "embed": case "h1": case "h2":
            case "h3": case "h4": case "h5": case "h6": case "head": case "hr": case "i": case "img":
            case "li": case "listing": case "menu": case "meta": case "nobr": case "ol": case "p":
            case "pre": case "ruby": case "s": case "small": case "span": case "strong": case "strike":
            case "sub": case "sup": case "table": case "tt": case "u": case "ul": case "var":
                return true;
            case "font":
                return HasForeignContentFontBreakoutAttribute(html, attributesStart, tagEnd);
            default:
                return false;
        }
    }

    private static bool HasForeignContentFontBreakoutAttribute(string html, int offset, int tagEnd) {
        while (offset < tagEnd) {
            while (offset < tagEnd && (IsAsciiWhitespace(html[offset]) || html[offset] == '/')) offset++;
            int nameStart = offset;
            while (offset < tagEnd && !IsAsciiWhitespace(html[offset]) && html[offset] is not '=' and not '/' and not '>') offset++;
            if (offset == nameStart) break;
            string name = html.Substring(nameStart, offset - nameStart);
            if (name.Equals("color", StringComparison.OrdinalIgnoreCase) ||
                name.Equals("face", StringComparison.OrdinalIgnoreCase) ||
                name.Equals("size", StringComparison.OrdinalIgnoreCase)) {
                return true;
            }

            while (offset < tagEnd && IsAsciiWhitespace(html[offset])) offset++;
            if (offset >= tagEnd || html[offset] != '=') continue;
            offset++;
            while (offset < tagEnd && IsAsciiWhitespace(html[offset])) offset++;
            if (offset >= tagEnd) break;
            char quote = html[offset];
            if (quote is '\'' or '"') {
                offset++;
                while (offset < tagEnd && html[offset] != quote) offset++;
                if (offset < tagEnd) offset++;
            } else {
                while (offset < tagEnd && !IsAsciiWhitespace(html[offset]) && html[offset] != '>') offset++;
            }
        }
        return false;
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

    private static int FindDoctypeEnd(string html, int offset) {
        int cursor = SkipAsciiWhitespace(html, offset);
        while (cursor < html.Length && html[cursor] != '>' && !IsAsciiWhitespace(html[cursor])) cursor++;
        if (cursor >= html.Length || html[cursor] == '>') return cursor < html.Length ? cursor : -1;
        cursor = SkipAsciiWhitespace(html, cursor);
        bool isPublic = StartsWithDoctypeKeyword(html, cursor, "PUBLIC");
        bool isSystem = StartsWithDoctypeKeyword(html, cursor, "SYSTEM");
        if (!isPublic && !isSystem) return html.IndexOf('>', cursor);
        cursor = SkipAsciiWhitespace(html, cursor + 6);
        if (cursor >= html.Length || html[cursor] is not ('\'' or '"')) return html.IndexOf('>', cursor);
        int firstIdentifierEnd = html.IndexOf(html[cursor], cursor + 1);
        if (firstIdentifierEnd < 0) return -1;
        cursor = SkipAsciiWhitespace(html, firstIdentifierEnd + 1);
        if (cursor >= html.Length || html[cursor] == '>') return cursor < html.Length ? cursor : -1;
        if (!isPublic || html[cursor] is not ('\'' or '"')) return html.IndexOf('>', cursor);
        int secondIdentifierEnd = html.IndexOf(html[cursor], cursor + 1);
        if (secondIdentifierEnd < 0) return -1;
        cursor = SkipAsciiWhitespace(html, secondIdentifierEnd + 1);
        return cursor < html.Length && html[cursor] == '>' ? cursor : html.IndexOf('>', cursor);
    }

    private static int SkipAsciiWhitespace(string value, int offset) {
        while (offset < value.Length && IsAsciiWhitespace(value[offset])) offset++;
        return offset;
    }

    private static bool IsAsciiWhitespace(char value) => value is '\t' or '\n' or '\f' or '\r' or ' ';

    private static bool StartsWithDoctypeKeyword(string html, int offset, string keyword) =>
        offset <= html.Length - keyword.Length &&
        string.Compare(html, offset, keyword, 0, keyword.Length, StringComparison.OrdinalIgnoreCase) == 0 &&
        offset + keyword.Length < html.Length && IsAsciiWhitespace(html[offset + keyword.Length]);

    private static int FindStartTagEnd(string html, int offset, out bool selfClosing) {
        selfClosing = false;
        char quote = '\0';
        bool afterEquals = false;
        bool unquotedValue = false;
        for (int index = offset; index < html.Length; index++) {
            char current = html[index];
            if (quote != '\0') {
                if (current == quote) quote = '\0';
                continue;
            }
            if (current == '>') return index;
            if (unquotedValue) {
                if (IsAsciiWhitespace(current)) unquotedValue = false;
                continue;
            }
            if (afterEquals) {
                if (IsAsciiWhitespace(current)) continue;
                afterEquals = false;
                if (current is '\'' or '"') quote = current;
                else unquotedValue = true;
                continue;
            }
            if (current == '=') afterEquals = true;
            else if (current == '/' && index + 1 < html.Length && html[index + 1] == '>') selfClosing = true;
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
        CancellationToken = source.Limits.CancellationToken,
        ProcessEmbeddedAssets = source.ProcessEmbeddedAssets && source.Limits.ProcessEmbeddedAssets,
        MaxEmbeddedAssets = Math.Min(source.MaxEmbeddedAssets, source.Limits.MaxEmbeddedAssets)
    };

    private static OfficeProvenanceOptions CreateOutputInspectionOptions(OfficeProvenanceRemovalOptions source) => new OfficeProvenanceOptions {
        MaxAssetBytes = source.EffectiveMaxOutputBytes,
        MaxManifestBytes = Math.Min(source.Limits.MaxManifestBytes, source.EffectiveMaxOutputBytes),
        MaxCarriers = source.Limits.MaxCarriers,
        MaxContainerEntries = source.Limits.MaxContainerEntries,
        MaxExpandedContainerBytes = source.Limits.MaxExpandedContainerBytes,
        CancellationToken = source.Limits.CancellationToken,
        ProcessEmbeddedAssets = source.ProcessEmbeddedAssets && source.Limits.ProcessEmbeddedAssets,
        MaxEmbeddedAssets = Math.Min(source.MaxEmbeddedAssets, source.Limits.MaxEmbeddedAssets)
    };

    private static OfficeProvenanceOptions CreateNestedOptions(OfficeProvenanceOptions source) => new OfficeProvenanceOptions {
        MaxAssetBytes = source.MaxAssetBytes,
        MaxManifestBytes = source.MaxManifestBytes,
        MaxCarriers = source.MaxCarriers,
        MaxContainerEntries = source.MaxContainerEntries,
        MaxExpandedContainerBytes = source.MaxExpandedContainerBytes,
        CancellationToken = source.CancellationToken,
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
            MaxEmbeddedAssets = source.MaxEmbeddedAssets,
            MaxOutputBytes = Math.Min(source.EffectiveMaxOutputBytes, source.Limits.MaxAssetBytes)
        };
        nested.Limits.MaxAssetBytes = source.Limits.MaxAssetBytes;
        nested.Limits.MaxManifestBytes = source.Limits.MaxManifestBytes;
        nested.Limits.MaxCarriers = source.Limits.MaxCarriers;
        nested.Limits.MaxContainerEntries = source.Limits.MaxContainerEntries;
        nested.Limits.MaxExpandedContainerBytes = source.Limits.MaxExpandedContainerBytes;
        nested.Limits.CancellationToken = source.Limits.CancellationToken;
        nested.Limits.ProcessEmbeddedAssets = false;
        nested.Limits.MaxEmbeddedAssets = source.Limits.MaxEmbeddedAssets;
        return nested;
    }

    private static byte[] ReadBounded(string filePath, long maximumBytes, CancellationToken cancellationToken) {
        string fullPath = Path.GetFullPath(filePath);
        using var stream = File.OpenRead(fullPath);
        return OfficeProvenanceBinary.ReadBounded(stream, maximumBytes, cancellationToken);
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
            throw OfficeProvenanceLimitException.CreateOutput(
                $"The rewritten HTML document exceeds the configured output limit of {maximumBytes} bytes.");
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
        internal EmbeddedImageReference(
            string attributeName,
            string value,
            int start,
            int length,
            string? containerText = null,
            string? containerMetadata = null,
            string? containerFragment = null) {
            AttributeName = attributeName;
            Value = value;
            Start = start;
            Length = length;
            ContainerText = containerText;
            ContainerMetadata = containerMetadata;
            ContainerFragment = containerFragment;
        }

        internal string AttributeName { get; }
        internal string Value { get; }
        internal int Start { get; }
        internal int Length { get; }
        internal string? ContainerText { get; }
        internal string? ContainerMetadata { get; }
        internal string? ContainerFragment { get; }
    }
}
