using OfficeIMO.Core.Internal;
using OfficeIMO.Epub;

namespace OfficeIMO.Chm;

/// <summary>Creates reflowable EPUB publications using the shared manuscript importer and archive-only resources.</summary>
public static class ChmEpubConverterExtensions {
    /// <summary>Imports a linked topic book and its embedded resources into an editable EPUB publication.</summary>
    /// <remarks>Topics resolve through the asynchronous archive resolver; the synchronous form waits for it on the thread pool.</remarks>
    public static EpubPublication ToEpub(this ChmDocument document, ChmEpubExportOptions? options = null,
        CancellationToken cancellationToken = default) => document.ToEpub(options, out _, cancellationToken);

    /// <summary>Imports a linked topic book into an EPUB publication and returns the source and import findings.</summary>
    public static EpubPublication ToEpub(this ChmDocument document, ChmEpubExportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        EpubManuscriptResult result = Task.Run(() => document.ConvertToEpubWithEvidenceAsync(options, cancellationToken),
            cancellationToken).GetAwaiter().GetResult();
        return OfficeConversionOutput.Complete(result.RequireValue(), result.Report, options?.FailOnLoss == true, out report);
    }

    /// <summary>Asynchronously imports a linked topic book. Resource resolution never contacts the network or opens external files.</summary>
    public static async Task<OfficeConversionResult<EpubPublication>> ToEpubAsync(this ChmDocument document,
        ChmEpubExportOptions? options = null, CancellationToken cancellationToken = default) {
        EpubManuscriptResult result = await document.ConvertToEpubWithEvidenceAsync(options, cancellationToken).ConfigureAwait(false);
        EpubPublication value = OfficeConversionOutput.Complete(result.RequireValue(), result.Report, options?.FailOnLoss == true, out OfficeConversionReport report);
        return new OfficeConversionResult<EpubPublication>(value, report);
    }

    /// <summary>Writes the converted EPUB package to a path and returns the import and package-write report.</summary>
    public static OfficeConversionReport SaveAsEpub(this ChmDocument document, string path, ChmEpubExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        var (bytes, report) = document.WriteEpub(options, cancellationToken);
        return OfficeConversionOutput.Save(report, path, bytes);
    }

    /// <summary>Writes the converted EPUB package to a caller-owned stream and returns the import and package-write report.</summary>
    public static OfficeConversionReport SaveAsEpub(this ChmDocument document, Stream stream, ChmEpubExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        var (bytes, report) = document.WriteEpub(options, cancellationToken);
        return OfficeConversionOutput.Save(report, stream, bytes);
    }

    /// <summary>Asynchronously writes the converted EPUB package to a path.</summary>
    public static async Task<OfficeConversionReport> SaveAsEpubAsync(this ChmDocument document, string path,
        ChmEpubExportOptions? options = null, CancellationToken cancellationToken = default) {
        var (bytes, report) = await document.WriteEpubAsync(options, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, path, bytes, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously writes the converted EPUB package to a caller-owned stream.</summary>
    public static async Task<OfficeConversionReport> SaveAsEpubAsync(this ChmDocument document, Stream stream,
        ChmEpubExportOptions? options = null, CancellationToken cancellationToken = default) {
        var (bytes, report) = await document.WriteEpubAsync(options, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, stream, bytes, cancellationToken).ConfigureAwait(false);
    }

    internal static EpubManuscriptResult ConvertToEpubWithEvidence(this ChmDocument document, ChmConversionOptions? options = null,
        EpubManuscriptOptions? epubOptions = null, CancellationToken cancellationToken = default) =>
        Task.Run(() => document.ConvertToEpubWithEvidenceAsync(Combine(options, epubOptions, null), cancellationToken), cancellationToken).GetAwaiter().GetResult();

    internal static Task<EpubManuscriptResult> ConvertToEpubWithEvidenceAsync(this ChmDocument document, ChmConversionOptions? options = null,
        EpubManuscriptOptions? epubOptions = null, CancellationToken cancellationToken = default) =>
        document.ConvertToEpubWithEvidenceAsync(Combine(options, epubOptions, null), cancellationToken);

    internal static ChmConversionResult<byte[]> ConvertToEpubBytesWithEvidence(this ChmDocument document, ChmConversionOptions? options = null,
        EpubManuscriptOptions? epubOptions = null, EpubSaveOptions? writeOptions = null, CancellationToken cancellationToken = default) {
        var (bytes, report) = document.WriteEpub(Combine(options, epubOptions, writeOptions), cancellationToken);
        return new ChmConversionResult<byte[]>(bytes, (ChmConversionReport)report);
    }

    private static ChmEpubExportOptions Combine(ChmConversionOptions? options, EpubManuscriptOptions? epubOptions, EpubSaveOptions? writeOptions) =>
        new ChmEpubExportOptions { ConversionOptions = options ?? new ChmConversionOptions(), EpubOptions = epubOptions, SaveOptions = writeOptions };

    internal static async Task<EpubManuscriptResult> ConvertToEpubWithEvidenceAsync(this ChmDocument document,
        ChmEpubExportOptions? options, CancellationToken cancellationToken) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        ChmConversionOptions configured = options?.ConversionOptions.Clone() ?? new ChmConversionOptions(); configured.EmbedImages = false;
        ChmConversionResult<HtmlConversionDocument> projected = document.ToHtmlDocumentResult(configured, cancellationToken);
        EpubManuscriptOptions conversion = options?.EpubOptions?.Clone() ?? new EpubManuscriptOptions();
        conversion.Title = conversion.Title ?? document.Title ?? "Compiled help";
        conversion.ResourceResolver = document.CreateResourceResolver();
        EpubManuscriptResult result = await EpubManuscript.ImportHtmlAsync(projected.Value, conversion, cancellationToken).ConfigureAwait(false);
        var diagnostics = projected.Report.Diagnostics.ToList();
        if (result.Succeeded) diagnostics.AddRange(RestoreNavigation(document, projected.Value, result.Publication, cancellationToken));
        return result.WithSourceReports(new ChmConversionReport(projected.Report.TopicPaths, diagnostics));
    }

    private static (byte[] Bytes, OfficeConversionReport Report) WriteEpub(this ChmDocument document, ChmEpubExportOptions? options,
        CancellationToken cancellationToken) =>
        Task.Run(() => document.WriteEpubAsync(options, cancellationToken), cancellationToken).GetAwaiter().GetResult();

    private static async Task<(byte[] Bytes, OfficeConversionReport Report)> WriteEpubAsync(this ChmDocument document,
        ChmEpubExportOptions? options, CancellationToken cancellationToken) {
        ChmConversionOptions configured = options?.ConversionOptions.Clone() ?? new ChmConversionOptions();
        EpubManuscriptResult imported = await document.ConvertToEpubWithEvidenceAsync(options, cancellationToken).ConfigureAwait(false);
        EpubSaveOptions requested = options?.SaveOptions ?? new EpubSaveOptions();
        var writing = new EpubSaveOptions {
            CompressEntries = requested.CompressEntries, MaxOutputBytes = Math.Min(requested.MaxOutputBytes, configured.MaxOutputBytes),
            MaxExpandedBytes = requested.MaxExpandedBytes, MaxEntries = requested.MaxEntries,
            RemoveInvalidatedSignatures = requested.RemoveInvalidatedSignatures, ModifiedAt = requested.ModifiedAt
        };
        var written = imported.RequireValue().Write(writing, cancellationToken);
        ChmDocument.EnforceOutput(written.Value.LongLength, configured);
        var report = new ChmConversionReport(document.SelectTopics(configured).Select(topic => topic.Path),
            imported.Report.Diagnostics.Concat(written.Report.Diagnostics));
        if (options?.FailOnLoss == true) report.ThrowIfLoss();
        return (written.Value, report);
    }

    private static IReadOnlyList<OfficeDiagnostic> RestoreNavigation(ChmDocument source, HtmlConversionDocument projection, EpubPublication publication, CancellationToken token) {
        var diagnostics = new List<OfficeDiagnostic>();
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
                    diagnostics.Add(new global::OfficeIMO.OfficeDiagnostic("CHM_EPUB_NAVIGATION_FRAGMENT", global::OfficeIMO.OfficeDiagnosticSeverity.Warning, "A contents fragment is absent from the projected topic; navigation uses the topic heading.", link.Target, "OfficeIMO.Chm.Epub", OfficeConversionLossKind.Approximation));
                }
                target = resolved; break;
            }
            if (item.Links.Count > 1) diagnostics.Add(new global::OfficeIMO.OfficeDiagnostic("CHM_EPUB_NAVIGATION_TARGETS", global::OfficeIMO.OfficeDiagnosticSeverity.Warning, "EPUB contents use one destination for a help contents item with several targets.", item.Name, "OfficeIMO.Chm.Epub", OfficeConversionLossKind.Omission));
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
