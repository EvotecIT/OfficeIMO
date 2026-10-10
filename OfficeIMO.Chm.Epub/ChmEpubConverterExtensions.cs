using OfficeIMO.Epub;

namespace OfficeIMO.Chm;

/// <summary>Creates reflowable EPUB publications using the shared manuscript importer and archive-only resources.</summary>
public static class ChmEpubConverterExtensions {
    /// <summary>Imports a linked topic book and its embedded resources into an editable EPUB publication.</summary>
    public static EpubManuscriptResult ToEpubPublicationResult(this ChmDocument document, ChmConversionOptions? options = null,
        EpubManuscriptOptions? epubOptions = null, CancellationToken cancellationToken = default) =>
        Task.Run(() => document.ToEpubPublicationResultAsync(options, epubOptions, cancellationToken), cancellationToken).GetAwaiter().GetResult();

    /// <summary>Asynchronously imports a linked topic book. Resource resolution never contacts the network or opens external files.</summary>
    public static async Task<EpubManuscriptResult> ToEpubPublicationResultAsync(this ChmDocument document, ChmConversionOptions? options = null,
        EpubManuscriptOptions? epubOptions = null, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        ChmConversionOptions configured = options?.Clone() ?? new ChmConversionOptions(); configured.EmbedImages = false;
        ChmConversionResult<HtmlConversionDocument> projected = document.ToHtmlDocumentResult(configured, cancellationToken);
        EpubManuscriptOptions conversion = epubOptions?.Clone() ?? new EpubManuscriptOptions();
        conversion.Title = conversion.Title ?? document.Title ?? "Compiled help";
        conversion.ResourceResolver = document.CreateResourceResolver();
        EpubManuscriptResult result = await EpubManuscript.ImportHtmlAsync(projected.Value, conversion, cancellationToken).ConfigureAwait(false);
        var diagnostics = projected.Report.FidelityDiagnostics.ToList();
        if (result.Succeeded) diagnostics.AddRange(RestoreNavigation(document, projected.Value, result.Publication, cancellationToken));
        return result.WithSourceReports(new ChmConversionReport(projected.Report.TopicPaths, diagnostics));
    }

    /// <summary>Creates serialized EPUB bytes and retains import and package-write fidelity reports.</summary>
    public static ChmConversionResult<byte[]> ToEpubBytesResult(this ChmDocument document, ChmConversionOptions? options = null,
        EpubManuscriptOptions? epubOptions = null, EpubWriteOptions? writeOptions = null, CancellationToken cancellationToken = default) {
        ChmConversionOptions configured = options?.Clone() ?? new ChmConversionOptions();
        EpubManuscriptResult imported = document.ToEpubPublicationResult(configured, epubOptions, cancellationToken);
        EpubWriteOptions requested = writeOptions ?? new EpubWriteOptions();
        var writing = new EpubWriteOptions {
            CompressEntries = requested.CompressEntries, MaxOutputBytes = Math.Min(requested.MaxOutputBytes, configured.MaxOutputBytes),
            MaxExpandedBytes = requested.MaxExpandedBytes, MaxEntries = requested.MaxEntries,
            RemoveInvalidatedSignatures = requested.RemoveInvalidatedSignatures, ModifiedAt = requested.ModifiedAt
        };
        var written = imported.RequireValue().Write(writing, cancellationToken);
        ChmDocument.EnforceOutput(written.Value.LongLength, configured);
        return new ChmConversionResult<byte[]>(written.Value, new ChmConversionReport(document.SelectTopics(configured).Select(topic => topic.Path),
            imported.Report.FidelityDiagnostics.Concat(written.Report.FidelityDiagnostics)));
    }

    private static IReadOnlyList<OfficeConversionFidelityDiagnostic> RestoreNavigation(ChmDocument source, HtmlConversionDocument projection, EpubPublication publication, CancellationToken token) {
        var diagnostics = new List<OfficeConversionFidelityDiagnostic>();
        var targets = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var anchorTargets = new Dictionary<string, string>(StringComparer.Ordinal);
        // Only direct book sections are generated topic boundaries. Retained authored
        // sections may use the same annotation and must not supply navigation ownership.
        var paths = projection.Document.Body!.Children.Where(section => section.NamespaceUri == "http://www.w3.org/1999/xhtml" &&
                section.LocalName == "section" && section.HasAttribute("data-chm-topic"))
            .ToDictionary(section => section.GetAttribute("id")!, section => section.GetAttribute("data-chm-topic")!, StringComparer.Ordinal);
        foreach (var item in publication.Manifest.Where(item => item.MediaType == "application/xhtml+xml")) {
            token.ThrowIfCancellationRequested();
            foreach (var element in publication.GetContentXml(item.Id).Descendants()) {
                string? anchor = (string?)element.Attribute("id");
                if (anchor == null) continue;
                string target = item.Reference.ContainerPath + "#" + Uri.EscapeDataString(anchor);
                anchorTargets[anchor] = target;
                if (paths.TryGetValue(anchor, out string? path)) targets[path] = target;
            }
        }
        EpubNavigationEntry? Map(ChmNavigationItem item) {
            token.ThrowIfCancellationRequested();
            var children = item.Children.Select(Map).Where(child => child != null).Cast<EpubNavigationEntry>().ToArray();
            string? target = null;
            foreach (ChmLink link in item.Links) {
                ChmEntry? entry = source.FindEntry(link.Target);
                if (entry == null || !targets.TryGetValue(entry.Path, out string? topicTarget)) continue;
                int fragment = link.Target.IndexOf('#');
                string sectionAnchor = Uri.UnescapeDataString(topicTarget.Substring(topicTarget.IndexOf('#') + 1));
                string anchor = fragment < 0 ? sectionAnchor : sectionAnchor + "-" + Uri.UnescapeDataString(link.Target.Substring(fragment + 1));
                if (!anchorTargets.TryGetValue(anchor, out string? resolved)) {
                    resolved = topicTarget;
                    diagnostics.Add(new OfficeConversionFidelityDiagnostic("CHM_EPUB_NAVIGATION_FRAGMENT", "A contents fragment is absent from the projected topic; navigation uses the topic heading.", OfficeConversionLossKind.Approximation, "OfficeIMO.Chm.Epub", link.Target));
                }
                target = resolved; break;
            }
            if (item.Links.Count > 1) diagnostics.Add(new OfficeConversionFidelityDiagnostic("CHM_EPUB_NAVIGATION_TARGETS", "EPUB contents use one destination for a help contents item with several targets.", OfficeConversionLossKind.Omission, "OfficeIMO.Chm.Epub", item.Name));
            target = target ?? children.FirstOrDefault()?.Target;
            return target == null ? null : new EpubNavigationEntry(item.Name, target, children);
        }
        var navigation = source.TableOfContents.Select(Map).Where(item => item != null).Cast<EpubNavigationEntry>().ToList();
        var listed = new HashSet<string>(ChmDocument.EnumerateNavigation(source.TableOfContents).SelectMany(item => item.Links)
            .Select(link => source.FindEntry(link.Target)?.Path).Where(path => path != null).Cast<string>(), StringComparer.OrdinalIgnoreCase);
        foreach (ChmTopic topic in source.Topics) if (!listed.Contains(topic.Path) && targets.TryGetValue(topic.Path, out string? target))
            navigation.Add(new EpubNavigationEntry(topic.Title, target));
        if (navigation.Count != 0) publication.SetNavigation(navigation);
        return diagnostics;
    }
}
