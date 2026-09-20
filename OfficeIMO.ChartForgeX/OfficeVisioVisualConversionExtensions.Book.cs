using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using global::ChartForgeX.VisualArtifacts;
using OfficeIMO.Visio;

namespace OfficeIMO.ChartForgeX;

public static partial class OfficeVisioVisualConversionExtensions {
    /// <summary>
    /// Projects a sequence of semantic envelopes into one editable multi-page Visio document.
    /// Input order becomes page order; every page retains its own fidelity report.
    /// </summary>
    public static OfficeVisioVisualBookResult ToOfficeVisioBook(
        this IEnumerable<VisualArtifactInterchangeEnvelope> envelopes, OfficeVisioVisualOptions? options = null) {
        return CreateOfficeVisioBook(envelopes, Array.Empty<OfficeVisioVisualBookLink>(), new OfficeVisioVisualBookOptions(), options);
    }

    /// <summary>
    /// Projects semantic envelopes into one editable multi-page Visio document and adds bounded,
    /// reciprocal navigation for relationships that cross page boundaries.
    /// </summary>
    public static OfficeVisioVisualBookResult ToOfficeVisioBook(
        this IEnumerable<VisualArtifactInterchangeEnvelope> envelopes,
        IEnumerable<OfficeVisioVisualBookLink> links,
        OfficeVisioVisualBookOptions? bookOptions = null,
        OfficeVisioVisualOptions? options = null) {
        if (links == null) throw new ArgumentNullException(nameof(links));
        return CreateOfficeVisioBook(envelopes, links, bookOptions ?? new OfficeVisioVisualBookOptions(), options);
    }

    private static OfficeVisioVisualBookResult CreateOfficeVisioBook(
        IEnumerable<VisualArtifactInterchangeEnvelope> envelopes,
        IEnumerable<OfficeVisioVisualBookLink> links,
        OfficeVisioVisualBookOptions bookOptions,
        OfficeVisioVisualOptions? options) {
        if (envelopes == null) throw new ArgumentNullException(nameof(envelopes));
        bookOptions.Validate();
        options ??= new OfficeVisioVisualOptions();
        var source = envelopes.ToList();
        var requestedLinks = links.ToList();
        if (source.Count == 0) throw new ArgumentException("At least one page is required.", nameof(envelopes));
        foreach (var envelope in source) {
            if (envelope == null) throw new ArgumentException("Pages cannot contain a null envelope.", nameof(envelopes));
            envelope.Validate();
        }
        var document = VisioDocument.Create();
        var pages = new List<OfficeVisioVisualConversionResult>();
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var envelope in source) {
            string baseName = string.IsNullOrWhiteSpace(envelope.Title) ? options.PageName : envelope.Title;
            string name = baseName;
            int suffix = 2;
            while (!names.Add(name)) name = baseName + " (" + suffix++ + ")";
            pages.Add(ProjectPage(envelope, document, options.ForPage(name)));
        }

        List<PendingNavigation> requestedNavigations = BuildRequestedNavigations(requestedLinks, pages, bookOptions.IncludeReturnLinks);
        List<PendingNavigation> distinctNavigations = CoalesceNavigations(requestedNavigations);
        var applied = new List<OfficeVisioVisualBookNavigationResult>();
        int omitted = ApplyNavigationLinks(distinctNavigations, pages, bookOptions, applied);
        return new OfficeVisioVisualBookResult(
            document,
            pages,
            applied,
            requestedNavigations.Count,
            requestedNavigations.Count - distinctNavigations.Count,
            omitted);
    }

    private static List<PendingNavigation> BuildRequestedNavigations(
        IReadOnlyList<OfficeVisioVisualBookLink> links,
        IReadOnlyList<OfficeVisioVisualConversionResult> pages,
        bool includeReturnLinks) {
        var result = new List<PendingNavigation>();
        for (int index = 0; index < links.Count; index++) {
            OfficeVisioVisualBookLink link = links[index] ?? throw new ArgumentException("Links cannot contain a null value.", nameof(links));
            ValidateLink(link, pages, index);
            result.Add(new PendingNavigation(
                link.SourcePageNumber,
                link.SourceEntityId,
                link.TargetPageNumber,
                link.TargetEntityId,
                link.RelationshipId,
                link.Description,
                false));
            if (includeReturnLinks) {
                result.Add(new PendingNavigation(
                    link.TargetPageNumber,
                    link.TargetEntityId,
                    link.SourcePageNumber,
                    link.SourceEntityId,
                    link.RelationshipId,
                    link.ReturnDescription,
                    true));
            }
        }
        return result;
    }

    private static void ValidateLink(
        OfficeVisioVisualBookLink link,
        IReadOnlyList<OfficeVisioVisualConversionResult> pages,
        int index) {
        if (link.SourcePageNumber < 1 || link.SourcePageNumber > pages.Count) {
            throw new ArgumentOutOfRangeException(nameof(link.SourcePageNumber), $"Link {index + 1} references source page {link.SourcePageNumber}; the book contains {pages.Count} pages.");
        }
        if (link.TargetPageNumber < 1 || link.TargetPageNumber > pages.Count) {
            throw new ArgumentOutOfRangeException(nameof(link.TargetPageNumber), $"Link {index + 1} references target page {link.TargetPageNumber}; the book contains {pages.Count} pages.");
        }
        if (link.SourcePageNumber == link.TargetPageNumber) {
            throw new ArgumentException($"Link {index + 1} must cross a page boundary.", nameof(link));
        }
        if (string.IsNullOrWhiteSpace(link.SourceEntityId)) {
            throw new ArgumentException($"Link {index + 1} has an empty source entity identifier.", nameof(link));
        }
        if (string.IsNullOrWhiteSpace(link.TargetEntityId)) {
            throw new ArgumentException($"Link {index + 1} has an empty target entity identifier.", nameof(link));
        }
        ResolveShape(pages[link.SourcePageNumber - 1].Page, link.SourceEntityId, "source", index);
        ResolveShape(pages[link.TargetPageNumber - 1].Page, link.TargetEntityId, "target", index);
    }

    private static List<PendingNavigation> CoalesceNavigations(IEnumerable<PendingNavigation> requested) {
        var result = new List<PendingNavigation>();
        foreach (IGrouping<string, PendingNavigation> group in requested.GroupBy(
                     item => NavigationKey(item.SourcePageNumber, item.SourceEntityId),
                     StringComparer.Ordinal)) {
            var distinct = new Dictionary<string, PendingNavigation>(StringComparer.Ordinal);
            foreach (PendingNavigation item in group) {
                string targetKey = NavigationKey(item.TargetPageNumber, item.TargetEntityId);
                if (!distinct.TryGetValue(targetKey, out PendingNavigation? existing)) {
                    distinct.Add(targetKey, item);
                    result.Add(item);
                } else {
                    foreach (string relationshipId in item.RelationshipIds) {
                        if (!existing.RelationshipIds.Contains(relationshipId, StringComparer.Ordinal)) {
                            existing.RelationshipIds.Add(relationshipId);
                        }
                    }
                }
            }
        }
        return result;
    }

    private static int ApplyNavigationLinks(
        IReadOnlyList<PendingNavigation> navigations,
        IReadOnlyList<OfficeVisioVisualConversionResult> pages,
        OfficeVisioVisualBookOptions options,
        ICollection<OfficeVisioVisualBookNavigationResult> applied) {
        int omittedTotal = 0;
        foreach (IGrouping<string, PendingNavigation> group in navigations.GroupBy(
                     item => NavigationKey(item.SourcePageNumber, item.SourceEntityId),
                     StringComparer.Ordinal)) {
            List<PendingNavigation> sourceNavigations = group.ToList();
            int appliedCount = Math.Min(sourceNavigations.Count, options.MaximumNavigationLinksPerEntity);
            int omitted = sourceNavigations.Count - appliedCount;
            omittedTotal += omitted;
            VisioPage sourcePage = pages[sourceNavigations[0].SourcePageNumber - 1].Page;
            VisioShape sourceShape = ResolveShape(sourcePage, sourceNavigations[0].SourceEntityId, "source", 0);
            for (int index = 0; index < appliedCount; index++) {
                PendingNavigation navigation = sourceNavigations[index];
                VisioPage targetPage = pages[navigation.TargetPageNumber - 1].Page;
                VisioShape targetShape = ResolveShape(targetPage, navigation.TargetEntityId, "target", 0);
                string description = string.IsNullOrWhiteSpace(navigation.Description)
                    ? (navigation.IsReturnLink ? "Back to " : "Open ") + targetPage.Name + ": " + DisplayShapeName(targetShape, navigation.TargetEntityId)
                    : navigation.Description!;
                VisioHyperlink hyperlink = sourceShape.AddPageHyperlink(targetPage.Name, description);
                if (options.IncludeNavigationShapeData) AddNavigationShapeData(sourceShape, index + 1, targetPage.Name, navigation);
                applied.Add(new OfficeVisioVisualBookNavigationResult(
                    navigation.SourcePageNumber,
                    navigation.SourceEntityId,
                    navigation.TargetPageNumber,
                    navigation.TargetEntityId,
                    new ReadOnlyCollection<string>(navigation.RelationshipIds),
                    navigation.IsReturnLink,
                    hyperlink));
            }
            if (options.IncludeNavigationShapeData && omitted > 0) {
                sourceShape.SetShapeData("CFX.BookLink.Omitted", omitted.ToString(CultureInfo.InvariantCulture));
            }
        }
        return omittedTotal;
    }

    private static void AddNavigationShapeData(
        VisioShape shape,
        int number,
        string targetPageName,
        PendingNavigation navigation) {
        string prefix = "CFX.BookLink." + number.ToString(CultureInfo.InvariantCulture) + ".";
        shape.SetShapeData(prefix + "TargetPage", targetPageName);
        shape.SetShapeData(prefix + "TargetEntityId", navigation.TargetEntityId);
        if (navigation.RelationshipIds.Count > 0) {
            shape.SetShapeData(prefix + "RelationshipIds", string.Join(",", navigation.RelationshipIds));
        }
    }

    private static VisioShape ResolveShape(VisioPage page, string entityId, string role, int linkIndex) {
        VisioShape[] direct = page.Shapes.Where(shape => string.Equals(shape.Id, entityId, StringComparison.Ordinal)).ToArray();
        if (direct.Length == 1) return direct[0];
        VisioShape[] semantic = page.Shapes.Where(shape =>
            string.Equals(shape.GetShapeDataValue("CFX.Id"), entityId, StringComparison.Ordinal)).ToArray();
        if (semantic.Length == 1) return semantic[0];
        VisioShape[] source = page.Shapes.Where(shape =>
            string.Equals(shape.GetShapeDataValue("Extension.chartforgex.sourceId"), entityId, StringComparison.Ordinal)).ToArray();
        if (source.Length == 1) return source[0];
        if (direct.Length + semantic.Length + source.Length == 0) {
            throw new ArgumentException($"Link {linkIndex + 1} {role} entity '{entityId}' was not projected on page '{page.Name}'.", "links");
        }
        throw new ArgumentException($"Link {linkIndex + 1} {role} entity '{entityId}' is ambiguous on page '{page.Name}'.", "links");
    }

    private static string DisplayShapeName(VisioShape shape, string fallback) =>
        string.IsNullOrWhiteSpace(shape.Text) ? fallback : shape.Text!.Replace(Environment.NewLine, " — ");

    private static string NavigationKey(int pageNumber, string entityId) =>
        pageNumber.ToString(CultureInfo.InvariantCulture) + "\u001f" + entityId;

    private sealed class PendingNavigation {
        public PendingNavigation(
            int sourcePageNumber,
            string sourceEntityId,
            int targetPageNumber,
            string targetEntityId,
            string? relationshipId,
            string? description,
            bool isReturnLink) {
            SourcePageNumber = sourcePageNumber;
            SourceEntityId = sourceEntityId;
            TargetPageNumber = targetPageNumber;
            TargetEntityId = targetEntityId;
            Description = description;
            IsReturnLink = isReturnLink;
            if (!string.IsNullOrWhiteSpace(relationshipId)) RelationshipIds.Add(relationshipId!);
        }

        public int SourcePageNumber { get; }
        public string SourceEntityId { get; }
        public int TargetPageNumber { get; }
        public string TargetEntityId { get; }
        public string? Description { get; }
        public bool IsReturnLink { get; }
        public List<string> RelationshipIds { get; } = new();
    }
}

/// <summary>A multi-page native Visio document with a fidelity report for every input envelope.</summary>
public sealed class OfficeVisioVisualBookResult {
    internal OfficeVisioVisualBookResult(
        VisioDocument document,
        List<OfficeVisioVisualConversionResult> pages,
        List<OfficeVisioVisualBookNavigationResult> navigations,
        int requestedNavigationCount,
        int coalescedNavigationCount,
        int omittedNavigationCount) {
        Document = document;
        Pages = pages.AsReadOnly();
        Navigations = navigations.AsReadOnly();
        RequestedNavigationCount = requestedNavigationCount;
        CoalescedNavigationCount = coalescedNavigationCount;
        OmittedNavigationCount = omittedNavigationCount;
    }
    /// <summary>Gets the document containing every projected page.</summary>
    public VisioDocument Document { get; }
    /// <summary>Gets page results in input order, including each page's semantic fidelity diagnostics.</summary>
    public IReadOnlyList<OfficeVisioVisualConversionResult> Pages { get; }
    /// <summary>Gets the internal page hyperlinks applied after deduplication and per-entity bounds.</summary>
    public IReadOnlyList<OfficeVisioVisualBookNavigationResult> Navigations { get; }
    /// <summary>Gets the requested navigation count, including generated reciprocal links.</summary>
    public int RequestedNavigationCount { get; }
    /// <summary>Gets the number of duplicate navigation requests coalesced into existing links.</summary>
    public int CoalescedNavigationCount { get; }
    /// <summary>Gets the number of distinct navigation requests omitted by per-entity bounds.</summary>
    public int OmittedNavigationCount { get; }
}
