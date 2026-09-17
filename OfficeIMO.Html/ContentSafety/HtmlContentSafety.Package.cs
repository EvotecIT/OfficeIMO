using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using AngleSharp;
using AngleSharp.Dom;
using AngleSharp.Html;
using AngleSharp.Html.Dom;
using AngleSharp.Xhtml;
using OfficeIMO.ContentSafety;
using OfficeIMO.Html.Providers;

namespace OfficeIMO.Html;

internal sealed class HtmlContentSafetyPackagePart {
    internal HtmlContentSafetyPackagePart(
        string key,
        string html,
        string locationPrefix,
        HtmlRenderOptions renderOptions,
        bool serializeAsXhtml = false) {
        Key = string.IsNullOrWhiteSpace(key) ? throw new ArgumentException("A package-part key is required.", nameof(key)) : key;
        Html = html ?? throw new ArgumentNullException(nameof(html));
        LocationPrefix = string.IsNullOrWhiteSpace(locationPrefix)
            ? throw new ArgumentException("A package-part location prefix is required.", nameof(locationPrefix))
            : locationPrefix;
        RenderOptions = renderOptions?.Clone() ?? throw new ArgumentNullException(nameof(renderOptions));
        SerializeAsXhtml = serializeAsXhtml;
    }

    internal string Key { get; }
    internal string Html { get; }
    internal string LocationPrefix { get; }
    internal HtmlRenderOptions RenderOptions { get; }
    internal bool SerializeAsXhtml { get; }
}

internal sealed class HtmlContentSafetyPackageCleanupResult {
    internal HtmlContentSafetyPackageCleanupResult(
        IReadOnlyDictionary<string, string> parts,
        IReadOnlyCollection<string> changedParts,
        OfficeContentSafetyReport before,
        OfficeContentSafetyReport after,
        IReadOnlyList<OfficeContentCleanupChange> changes) {
        Parts = parts;
        ChangedParts = changedParts;
        Before = before;
        After = after;
        Changes = changes;
    }

    internal IReadOnlyDictionary<string, string> Parts { get; }
    internal IReadOnlyCollection<string> ChangedParts { get; }
    internal OfficeContentSafetyReport Before { get; }
    internal OfficeContentSafetyReport After { get; }
    internal IReadOnlyList<OfficeContentCleanupChange> Changes { get; }
}

public static partial class HtmlContentSafety {
    private const string XhtmlNamespace = "http://www.w3.org/1999/xhtml";

    internal static async Task<OfficeContentSafetyReport> InspectPackagePartsAsync(
        string format,
        IReadOnlyList<HtmlContentSafetyPackagePart> parts,
        OfficeContentSafetyOptions? options,
        CancellationToken cancellationToken) {
        OfficeContentSafetyOptions effective = options ?? new OfficeContentSafetyOptions();
        effective.Validate();
        ValidateAggregatePackageText(parts, effective);
        var builder = new OfficeContentSafetyBuilder(format, effective);
        var resourceBudget = new PackageResourceBudget(effective);
        HtmlCssByteBudget cssBudget = CreatePackageCssBudget(effective);
        foreach (HtmlContentSafetyPackagePart part in parts) {
            cancellationToken.ThrowIfCancellationRequested();
            PreparedPackagePart prepared = await PreparePackagePartAsync(
                part,
                effective,
                resourceBudget,
                cssBudget,
                cancellationToken).ConfigureAwait(false);
            InspectDocument(prepared.Document, builder, targets: null, part.LocationPrefix, prepared.SyntheticStyles, prepared.Limits);
        }
        return builder.Build();
    }

    internal static async Task<HtmlContentSafetyPackageCleanupResult> RemoveSelectedPackagePartsAsync(
        string format,
        IReadOnlyList<HtmlContentSafetyPackagePart> parts,
        OfficeContentCleanupSelection selection,
        OfficeContentSafetyOptions? options,
        CancellationToken cancellationToken) {
        if (selection == null) throw new ArgumentNullException(nameof(selection));
        OfficeContentSafetyOptions effective = options ?? new OfficeContentSafetyOptions();
        effective.Validate();
        ValidateAggregatePackageText(parts, effective);
        var builder = new OfficeContentSafetyBuilder(format, effective);
        var preparedParts = new List<PreparedPackagePart>(parts.Count);
        var targets = new Dictionary<string, PackageCleanupTarget>(StringComparer.Ordinal);
        var resourceBudget = new PackageResourceBudget(effective);
        HtmlCssByteBudget cssBudget = CreatePackageCssBudget(effective);
        foreach (HtmlContentSafetyPackagePart part in parts) {
            cancellationToken.ThrowIfCancellationRequested();
            PreparedPackagePart prepared = await PreparePackagePartAsync(
                part,
                effective,
                resourceBudget,
                cssBudget,
                cancellationToken).ConfigureAwait(false);
            var partTargets = new Dictionary<string, HtmlCleanupTarget>(StringComparer.Ordinal);
            InspectDocument(prepared.Document, builder, partTargets, part.LocationPrefix, prepared.SyntheticStyles, prepared.Limits);
            foreach (KeyValuePair<string, HtmlCleanupTarget> target in partTargets) {
                targets[target.Key] = new PackageCleanupTarget(prepared, target.Value);
            }
            preparedParts.Add(prepared);
        }

        OfficeContentSafetyReport before = builder.Build();
        IReadOnlyList<OfficeContentSafetyFinding> selected = OfficeContentSafetyBuilder.ResolveSelection(before, selection);
        if (selected.Count == 0) {
            var unchanged = parts.ToDictionary(part => part.Key, part => part.Html, StringComparer.Ordinal);
            return new HtmlContentSafetyPackageCleanupResult(
                unchanged,
                Array.Empty<string>(),
                before,
                before,
                Array.Empty<OfficeContentCleanupChange>());
        }

        foreach (IGrouping<PackageCleanupTarget, OfficeContentSafetyFinding> group in selected
            .OrderByDescending(item => item.SourceTextOffset ?? -1)
            .GroupBy(item => targets[item.Id])) {
            group.Key.Part.Changed = true;
            group.Key.Target.Remove();
        }

        var output = new Dictionary<string, string>(StringComparer.Ordinal);
        var afterParts = new List<HtmlContentSafetyPackagePart>(parts.Count);
        foreach (PreparedPackagePart prepared in preparedParts) {
            cancellationToken.ThrowIfCancellationRequested();
            string html;
            if (prepared.Changed) {
                foreach (IElement style in prepared.SyntheticStyles) style.Remove();
                foreach (KeyValuePair<IElement, string> original in prepared.OriginalInlineStyles) {
                    original.Key.TextContent = original.Value;
                }
                html = prepared.Source.SerializeAsXhtml
                    ? prepared.Document.ToHtml(XhtmlMarkupFormatter.Instance)
                    : prepared.Document.ToHtml(HtmlMarkupFormatter.Instance);
            } else {
                html = prepared.Source.Html;
            }
            output.Add(prepared.Source.Key, html);
            afterParts.Add(new HtmlContentSafetyPackagePart(
                prepared.Source.Key,
                html,
                prepared.Source.LocationPrefix,
                prepared.Source.RenderOptions,
                prepared.Source.SerializeAsXhtml));
        }

        OfficeContentSafetyReport after = await InspectPackagePartsAsync(format, afterParts, effective, cancellationToken).ConfigureAwait(false);
        OfficeContentCleanupChange[] changes = selected
            .Select(item => new OfficeContentCleanupChange(item.Id, item.Location, item.CleanupCapability))
            .ToArray();
        string[] changedParts = preparedParts.Where(item => item.Changed).Select(item => item.Source.Key).ToArray();
        return new HtmlContentSafetyPackageCleanupResult(output, changedParts, before, after, changes);
    }

    private static async Task<PreparedPackagePart> PreparePackagePartAsync(
        HtmlContentSafetyPackagePart part,
        OfficeContentSafetyOptions safetyOptions,
        PackageResourceBudget resourceBudget,
        HtmlCssByteBudget cssBudget,
        CancellationToken cancellationToken) {
        OfficeContentSafetyInputGuard.ValidateText(part.Html, safetyOptions);
        cancellationToken.ThrowIfCancellationRequested();

        HtmlRenderOptions renderOptions = part.RenderOptions.Clone();
        ApplySafetyLimits(renderOptions, safetyOptions);
        var limits = HtmlConversionLimits.CreateUntrustedProfile();
        limits.MaxInputCharacters = Math.Min(limits.MaxInputCharacters ?? int.MaxValue, safetyOptions.MaxCharacters);
        limits.MaxCssBytes = Math.Min(limits.MaxCssBytes ?? long.MaxValue, renderOptions.MaxResourceBytes);
        limits.MaxTotalCssBytes = Math.Min(limits.MaxTotalCssBytes ?? long.MaxValue, renderOptions.MaxTotalResourceBytes);

        IHtmlDocument document = ParsePackageDocument(part, renderOptions.BaseUri, limits, safetyOptions, cancellationToken);
        if (document.QuerySelectorAll("meta[http-equiv]").Any(meta =>
                string.Equals(meta.GetAttribute("http-equiv")?.Trim(), "Content-Security-Policy", StringComparison.OrdinalIgnoreCase))) {
            throw new InvalidDataException(
                "Package content-safety inspection does not apply stylesheets when the document declares a Content Security Policy.");
        }
        string[] titledStylesheetSets = document.QuerySelectorAll("link[href]")
            .Where(link => HtmlRenderStylesheetApplier.IsApplicableStylesheetLink(link, renderOptions))
            .Select(link => link.GetAttribute("title")?.Trim())
            .Where(title => !string.IsNullOrEmpty(title))
            .Select(title => title!)
            .Distinct(StringComparer.Ordinal)
            .ToArray();
        if (titledStylesheetSets.Length > 1) {
            throw new InvalidDataException(
                "Package content-safety inspection does not support conflicting preferred stylesheet sets: "
                + string.Join(", ", titledStylesheetSets));
        }
        foreach (IElement style in document.QuerySelectorAll("style")) {
            if (!HtmlRenderStylesheetApplier.IsApplicableStyleElement(style, renderOptions)) continue;
            cssBudget.ReserveOrThrow(style.TextContent ?? string.Empty);
        }
        var resourceOptions = new HtmlResourcePipelineOptions {
            BaseUri = renderOptions.BaseUri,
            UrlPolicy = renderOptions.UrlPolicy.Clone(),
            ResourceUrlPolicy = renderOptions.ResourceUrlPolicy?.Clone(),
            Limits = limits.Clone(),
            MaxResponsiveImageCandidates = renderOptions.ResponsiveImageCandidateLimit,
            MediaContext = renderOptions.MediaContext,
            MediaWidth = renderOptions.ViewportWidth,
            MediaHeight = renderOptions.ViewportHeight,
            MediaFeatures = renderOptions.MediaFeatures.Clone()
        };
        HtmlResourceManifest discovered = HtmlResourcePipeline.BuildManifest(document, resourceOptions);
        var stylesheets = new HtmlResourceManifest();
        foreach (HtmlResourceReference reference in discovered.Resources.Where(item => item.Kind == HtmlResourceKind.Stylesheet)) {
            stylesheets.Add(reference);
            if (!reference.IsAllowed) {
                throw new InvalidDataException("A package stylesheet reference was rejected by the configured URL policy: " + reference.Source);
            }
            if (reference.ResolvedSource.StartsWith("data:", StringComparison.OrdinalIgnoreCase)) {
                throw new InvalidDataException("Linked data-URI stylesheets are not supported by package content-safety inspection.");
            }
        }

        var requestedStylesheets = new HashSet<string>(HtmlResourceIdentityComparer.Instance);
        var requestSync = new object();
        HtmlRenderResourceResolver? resolver = renderOptions.ResourceResolver;
        if (resolver != null) {
            renderOptions.ResourceResolver = async (request, token) => {
                if (request.Kind == HtmlResourceKind.Stylesheet) {
                    lock (requestSync) requestedStylesheets.Add(request.Uri.AbsoluteUri);
                }
                return await resolver(request, token).ConfigureAwait(false);
            };
        }

        var diagnostics = new HtmlDiagnosticReport();
        var existingStyles = new HashSet<IElement>(document.QuerySelectorAll("style"));
        var originalInlineStyles = existingStyles.ToDictionary(
            style => style,
            style => style.TextContent ?? string.Empty);
        if (stylesheets.Resources.Count > 0) resourceBudget.Apply(renderOptions);
        HtmlResourceSession resources = await HtmlRenderResourceLoader.LoadAsync(
            stylesheets,
            renderOptions,
            diagnostics,
            limits,
            cancellationToken,
            cssBudget).ConfigureAwait(false);
        resourceBudget.Reserve(resources);

        var acceptedStylesheets = new HashSet<string>(
            resources.Resources
                .Where(item => item.Kind == HtmlResourceKind.Stylesheet)
                .Select(item => item.CanonicalSource),
            HtmlResourceIdentityComparer.Instance);
        string? missing;
        lock (requestSync) missing = requestedStylesheets.FirstOrDefault(uri => !acceptedStylesheets.Contains(uri));
        missing ??= stylesheets.Resources
            .Where(item => item.IsAllowed && item.ResolvedSource.Length > 0)
            .Select(item => item.ResolvedSource)
            .FirstOrDefault(uri => !acceptedStylesheets.Contains(uri));
        if (missing != null) throw new InvalidDataException("A package stylesheet could not be resolved: " + missing);
        if (diagnostics.HasErrors) throw CreateStylesheetException(diagnostics);

        HtmlRenderStylesheetApplier.Apply(document, resources, renderOptions, limits, cssBudget, diagnostics);
        HtmlDiagnostic? unsafeDiagnostic = diagnostics.FirstOrDefault(item =>
            item.Severity == HtmlDiagnosticSeverity.Error
            || item.Code == "StylesheetResourceRejectedByPolicy"
            || item.Code == HtmlRenderDiagnosticCodes.StylesheetEncodingUnsupported
            || item.Code == HtmlRenderDiagnosticCodes.StylesheetImportCycle
            || item.Code == HtmlRenderDiagnosticCodes.StylesheetImportDepthExceeded);
        if (unsafeDiagnostic != null) throw CreateStylesheetException(diagnostics);

        var syntheticStyles = new HashSet<IElement>(document.QuerySelectorAll("style").Where(item => !existingStyles.Contains(item)));
        return new PreparedPackagePart(part, document, syntheticStyles, originalInlineStyles, limits);
    }

    private static IHtmlDocument ParsePackageDocument(
        HtmlContentSafetyPackagePart part,
        Uri? baseUri,
        HtmlConversionLimits limits,
        OfficeContentSafetyOptions safetyOptions,
        CancellationToken cancellationToken) {
        if (!part.SerializeAsXhtml) {
            HtmlConversionDocument conversion = HtmlConversionDocument.Parse(part.Html, new HtmlConversionDocumentOptions {
                BaseUri = baseUri,
                Limits = limits.Clone()
            });
            return conversion.CreateSourceDocumentForConversion();
        }

        var settings = new XmlReaderSettings {
            DtdProcessing = DtdProcessing.Parse,
            XmlResolver = null,
            MaxCharactersInDocument = safetyOptions.MaxCharacters,
            MaxCharactersFromEntities = safetyOptions.MaxCharacters
        };
        XDocument xml;
        try {
            using var text = new StringReader(part.Html);
            using XmlReader reader = XmlReader.Create(text, settings);
            xml = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        } catch (XmlException exception) {
            throw new InvalidDataException("An EPUB XHTML content document is not well-formed XML.", exception);
        }

        XNamespace xhtml = "http://www.w3.org/1999/xhtml";
        if (xml.Root == null || xml.Root.Name.LocalName != "html" || xml.Root.Name.Namespace != xhtml) {
            throw new InvalidDataException("An EPUB XHTML content document must have an html root in the XHTML namespace.");
        }
        XDocumentType? documentType = xml.Nodes().OfType<XDocumentType>().SingleOrDefault();
        if (documentType != null
            && (!documentType.Name.Equals("html", StringComparison.OrdinalIgnoreCase)
                || !string.IsNullOrWhiteSpace(documentType.PublicId)
                || !string.IsNullOrWhiteSpace(documentType.SystemId)
                || !string.IsNullOrWhiteSpace(documentType.InternalSubset))) {
            throw new InvalidDataException("EPUB XHTML content-safety inspection accepts only the entity-free HTML5 document type.");
        }
        if (xml.DescendantNodes().Any(node => node is XProcessingInstruction)
            || xml.Nodes().Any(node => node is XProcessingInstruction)) {
            throw new InvalidDataException("EPUB XHTML content-safety inspection does not accept processing instructions.");
        }

        var owned = new OfficeIMO.Html.Dom.HtmlDocument(
            AngleSharpDomServices.Instance,
            AngleSharpHtmlParser.Instance.Id);
        int nodeCount = 0;
        foreach (XNode node in xml.Nodes()) {
            if (node is XDocumentType) {
                owned.AppendChild(owned.CreateDocumentType("html"));
                nodeCount++;
                continue;
            }
            AppendXmlNode(node, owned, owned, limits, ref nodeCount, depth: 1, cancellationToken);
        }
        cancellationToken.ThrowIfCancellationRequested();
        return NativeDomBridge.GetNativeDocument(owned, cancellationToken);
    }

    private static void AppendXmlNode(
        XNode source,
        OfficeIMO.Html.Dom.HtmlDocument document,
        OfficeIMO.Html.Dom.HtmlNode parent,
        HtmlConversionLimits limits,
        ref int nodeCount,
        int depth,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        nodeCount = checked(nodeCount + 1);
        if (limits.MaxHtmlNodes.HasValue && nodeCount > limits.MaxHtmlNodes.Value) {
            throw new InvalidDataException("An EPUB XHTML content document exceeds the configured HTML node limit.");
        }
        if (limits.MaxHtmlDepth.HasValue && depth > limits.MaxHtmlDepth.Value) {
            throw new InvalidDataException("An EPUB XHTML content document exceeds the configured HTML depth limit.");
        }

        if (source is XElement element) {
            string namespaceUri = element.Name.NamespaceName;
            if (namespaceUri == XhtmlNamespace
                && element.Name.LocalName != element.Name.LocalName.ToLowerInvariant()) {
                throw new InvalidDataException(
                    "EPUB XHTML content-safety inspection requires canonical lowercase XHTML element names.");
            }
            string? prefix = element.GetPrefixOfNamespace(element.Name.Namespace);
            OfficeIMO.Html.Dom.HtmlElement target = document.CreateElement(element.Name.LocalName, namespaceUri, prefix);
            foreach (XAttribute attribute in element.Attributes()) {
                cancellationToken.ThrowIfCancellationRequested();
                string name;
                if (attribute.IsNamespaceDeclaration) {
                    name = attribute.Name.LocalName == "xmlns" ? "xmlns" : "xmlns:" + attribute.Name.LocalName;
                } else {
                    if (namespaceUri == XhtmlNamespace
                        && string.IsNullOrEmpty(attribute.Name.NamespaceName)
                        && attribute.Name.LocalName != attribute.Name.LocalName.ToLowerInvariant()) {
                        throw new InvalidDataException(
                            "EPUB XHTML content-safety inspection requires canonical lowercase unqualified XHTML attribute names.");
                    }
                    string? attributePrefix = element.GetPrefixOfNamespace(attribute.Name.Namespace);
                    name = string.IsNullOrEmpty(attributePrefix)
                        ? attribute.Name.LocalName
                        : attributePrefix + ":" + attribute.Name.LocalName;
                }
                target.SetAttribute(name, attribute.Value, attribute.Name.NamespaceName);
            }
            parent.AppendChild(target);
            foreach (XNode child in element.Nodes()) {
                AppendXmlNode(child, document, target, limits, ref nodeCount, depth + 1, cancellationToken);
            }
            return;
        }
        if (source is XText text) {
            parent.AppendChild(document.CreateTextNode(text.Value));
            return;
        }
        if (source is XComment comment) {
            parent.AppendChild(document.CreateComment(comment.Value));
            return;
        }
        throw new InvalidDataException("EPUB XHTML content-safety inspection encountered an unsupported XML node.");
    }

    private static HtmlCssByteBudget CreatePackageCssBudget(OfficeContentSafetyOptions options) {
        HtmlConversionLimits limits = HtmlConversionLimits.CreateUntrustedProfile();
        limits.MaxCssBytes = Math.Min(limits.MaxCssBytes ?? long.MaxValue, options.MaxInputBytes);
        limits.MaxTotalCssBytes = Math.Min(limits.MaxTotalCssBytes ?? long.MaxValue, options.MaxExpandedPackageBytes);
        return new HtmlCssByteBudget(limits);
    }

    private static void ValidateAggregatePackageText(
        IEnumerable<HtmlContentSafetyPackagePart> parts,
        OfficeContentSafetyOptions options) {
        long characters = 0;
        foreach (HtmlContentSafetyPackagePart part in parts) {
            if (part.Html.Length > options.MaxCharacters - characters) {
                throw new InvalidDataException("The package exceeds the configured decoded-character limit.");
            }
            characters += part.Html.Length;
        }
    }

    private static void ApplySafetyLimits(HtmlRenderOptions options, OfficeContentSafetyOptions safetyOptions) {
        options.MaxInputCharacters = Math.Min(options.MaxInputCharacters, safetyOptions.MaxCharacters);
        options.MaxResourceBytes = Math.Min(options.MaxResourceBytes, safetyOptions.MaxInputBytes);
        options.MaxTotalResourceBytes = Math.Min(options.MaxTotalResourceBytes, safetyOptions.MaxExpandedPackageBytes);
        options.MaxResourceCount = Math.Min(options.MaxResourceCount, safetyOptions.MaxPackageEntries);
        options.MaxResourceRequests = Math.Min(options.MaxResourceRequests, safetyOptions.MaxPackageEntries);
    }

    private static InvalidDataException CreateStylesheetException(HtmlDiagnosticReport diagnostics) {
        string detail = string.Join("; ", diagnostics.Diagnostics.Select(item => item.Code + ": " + item.Message));
        return new InvalidDataException("Package stylesheet resolution was not complete and safe. " + detail);
    }

    private sealed class PreparedPackagePart {
        internal PreparedPackagePart(
            HtmlContentSafetyPackagePart source,
            IHtmlDocument document,
            ISet<IElement> syntheticStyles,
            IReadOnlyDictionary<IElement, string> originalInlineStyles,
            HtmlConversionLimits limits) {
            Source = source;
            Document = document;
            SyntheticStyles = syntheticStyles;
            OriginalInlineStyles = originalInlineStyles;
            Limits = limits;
        }

        internal HtmlContentSafetyPackagePart Source { get; }
        internal IHtmlDocument Document { get; }
        internal ISet<IElement> SyntheticStyles { get; }
        internal IReadOnlyDictionary<IElement, string> OriginalInlineStyles { get; }
        internal HtmlConversionLimits Limits { get; }
        internal bool Changed { get; set; }
    }

    private sealed class PackageCleanupTarget {
        internal PackageCleanupTarget(PreparedPackagePart part, HtmlCleanupTarget target) {
            Part = part;
            Target = target;
        }

        internal PreparedPackagePart Part { get; }
        internal HtmlCleanupTarget Target { get; }
    }

    private sealed class PackageResourceBudget {
        private readonly long _maximumBytes;
        private readonly int _maximumResources;
        private long _acceptedBytes;
        private int _acceptedResources;
        private int _resolverRequests;

        internal PackageResourceBudget(OfficeContentSafetyOptions options) {
            _maximumBytes = options.MaxExpandedPackageBytes;
            _maximumResources = options.MaxPackageEntries;
        }

        internal void Apply(HtmlRenderOptions options) {
            long remainingBytes = _maximumBytes - _acceptedBytes;
            int remainingResources = _maximumResources - _acceptedResources;
            int remainingRequests = _maximumResources - _resolverRequests;
            if (remainingBytes <= 0 || remainingResources <= 0 || remainingRequests <= 0) {
                throw new InvalidDataException("Package stylesheet resources exceed the configured aggregate limits.");
            }
            options.MaxResourceBytes = Math.Min(options.MaxResourceBytes, remainingBytes);
            options.MaxTotalResourceBytes = Math.Min(options.MaxTotalResourceBytes, remainingBytes);
            options.MaxResourceCount = Math.Min(options.MaxResourceCount, remainingResources);
            options.MaxResourceRequests = Math.Min(options.MaxResourceRequests, remainingRequests);
            if (options.MaxResourceRequests < options.MaxResourceCount) {
                options.MaxResourceCount = options.MaxResourceRequests;
            }
        }

        internal void Reserve(HtmlResourceSession resources) {
            _acceptedBytes = checked(_acceptedBytes + resources.AcceptedResourceBytes);
            _acceptedResources = checked(_acceptedResources + resources.AcceptedResourceCount);
            _resolverRequests = checked(_resolverRequests + resources.ResolverRequestCount);
            if (_acceptedBytes > _maximumBytes
                || _acceptedResources > _maximumResources
                || _resolverRequests > _maximumResources) {
                throw new InvalidDataException("Package stylesheet resources exceed the configured aggregate limits.");
            }
        }
    }
}
