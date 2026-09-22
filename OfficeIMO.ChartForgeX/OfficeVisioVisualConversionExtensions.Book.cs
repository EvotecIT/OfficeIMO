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
    public static OfficeVisioVisualBookResult ToOfficeVisioBookWithNavigation(
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
        var requestedLinks = new List<OfficeVisioVisualBookLink>();
        foreach (OfficeVisioVisualBookLink link in links) {
            if (requestedLinks.Count >= bookOptions.MaximumRequestedLinks)
                throw new ArgumentException("The book exceeds MaximumRequestedLinks.", nameof(links));
            requestedLinks.Add(link);
        }
        if (source.Count == 0) throw new ArgumentException("At least one page is required.", nameof(envelopes));
        foreach (var envelope in source) {
            if (envelope == null) throw new ArgumentException("Pages cannot contain a null envelope.", nameof(envelopes));
            envelope.Validate();
        }
        var document = VisioDocument.Create();
        var pages = new List<OfficeVisioVisualConversionResult>();
        var projections = new List<BookPageProjection>();
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var envelope in source) {
            string baseName = string.IsNullOrWhiteSpace(envelope.Title) ? options.PageName : envelope.Title;
            string name = baseName;
            int suffix = 2;
            while (!names.Add(name)) name = baseName + " (" + suffix++ + ")";
            OfficeVisioVisualConversionResult page = ProjectPage(envelope, document, options.ForPage(name));
            pages.Add(page);
            projections.Add(new BookPageProjection(envelope, page));
        }

        List<PendingNavigation> requestedNavigations = BuildRequestedNavigations(requestedLinks, projections, bookOptions);
        List<PendingNavigation> distinctNavigations = CoalesceNavigations(requestedNavigations, bookOptions);
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
        IReadOnlyList<BookPageProjection> pages,
        OfficeVisioVisualBookOptions options) {
        var result = new List<PendingNavigation>();
        for (int index = 0; index < links.Count; index++) {
            OfficeVisioVisualBookLink link = links[index] ?? throw new ArgumentException("Links cannot contain a null value.", nameof(links));
            (VisioShape sourceShape, VisioShape targetShape) = ResolveLink(link, pages, index);
            result.Add(new PendingNavigation(
                link.SourcePageNumber,
                link.SourceEntityId,
                sourceShape,
                link.TargetPageNumber,
                link.TargetEntityId,
                targetShape,
                link.RelationshipId,
                link.Description,
                false, options));
            if (options.IncludeReturnLinks) {
                result.Add(new PendingNavigation(
                    link.TargetPageNumber,
                    link.TargetEntityId,
                    targetShape,
                    link.SourcePageNumber,
                    link.SourceEntityId,
                    sourceShape,
                    link.RelationshipId,
                    link.ReturnDescription,
                    true, options));
            }
        }
        return result;
    }

    private static (VisioShape Source, VisioShape Target) ResolveLink(
        OfficeVisioVisualBookLink link,
        IReadOnlyList<BookPageProjection> pages,
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
        VisioShape source = pages[link.SourcePageNumber - 1].Resolve(link.SourceEntityId, "source", index);
        VisioShape target = pages[link.TargetPageNumber - 1].Resolve(link.TargetEntityId, "target", index);
        return (source, target);
    }

    private static List<PendingNavigation> CoalesceNavigations(IEnumerable<PendingNavigation> requested,
        OfficeVisioVisualBookOptions options) {
        var result = new List<PendingNavigation>();
        foreach (IGrouping<string, PendingNavigation> group in requested.GroupBy(
                     item => NavigationKey(item.SourcePageNumber, item.SourceShape.Id),
                     StringComparer.Ordinal)) {
            var distinct = new Dictionary<string, PendingNavigation>(StringComparer.Ordinal);
            foreach (PendingNavigation item in group) {
                string targetKey = NavigationKey(item.TargetPageNumber, item.TargetShape.Id);
                if (!distinct.TryGetValue(targetKey, out PendingNavigation? existing)) {
                    distinct.Add(targetKey, item);
                    result.Add(item);
                } else {
                    foreach (string relationshipId in item.RelationshipIds) {
                        existing.AddRelationshipId(relationshipId, options);
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
                     item => NavigationKey(item.SourcePageNumber, item.SourceShape.Id),
                     StringComparer.Ordinal)) {
            List<PendingNavigation> sourceNavigations = group.ToList();
            int appliedCount = Math.Min(sourceNavigations.Count, options.MaximumNavigationLinksPerEntity);
            int omitted = sourceNavigations.Count - appliedCount;
            omittedTotal += omitted;
            VisioShape sourceShape = sourceNavigations[0].SourceShape;
            for (int index = 0; index < appliedCount; index++) {
                PendingNavigation navigation = sourceNavigations[index];
                VisioPage targetPage = pages[navigation.TargetPageNumber - 1].Page;
                VisioShape targetShape = navigation.TargetShape;
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

    private static string DisplayShapeName(VisioShape shape, string fallback) =>
        string.IsNullOrWhiteSpace(shape.Text) ? fallback : shape.Text!.Replace(Environment.NewLine, " — ");

    private static string NavigationKey(int pageNumber, string entityId) =>
        pageNumber.ToString(CultureInfo.InvariantCulture) + "\u001f" + entityId;

    private sealed class PendingNavigation {
        public PendingNavigation(
            int sourcePageNumber,
            string sourceEntityId,
            VisioShape sourceShape,
            int targetPageNumber,
            string targetEntityId,
            VisioShape targetShape,
            string? relationshipId,
            string? description,
            bool isReturnLink,
            OfficeVisioVisualBookOptions options) {
            SourcePageNumber = sourcePageNumber;
            SourceEntityId = sourceEntityId;
            SourceShape = sourceShape;
            TargetPageNumber = targetPageNumber;
            TargetEntityId = targetEntityId;
            TargetShape = targetShape;
            Description = description;
            IsReturnLink = isReturnLink;
            if (!string.IsNullOrWhiteSpace(relationshipId)) AddRelationshipId(relationshipId!, options);
        }

        public int SourcePageNumber { get; }
        public string SourceEntityId { get; }
        public VisioShape SourceShape { get; }
        public int TargetPageNumber { get; }
        public string TargetEntityId { get; }
        public VisioShape TargetShape { get; }
        public string? Description { get; }
        public bool IsReturnLink { get; }
        public List<string> RelationshipIds { get; } = new();
        private readonly HashSet<string> _relationshipIds = new(StringComparer.Ordinal);
        private int _relationshipIdCharacters;

        public void AddRelationshipId(string relationshipId, OfficeVisioVisualBookOptions options) {
            if (_relationshipIds.Contains(relationshipId)) return;
            if (RelationshipIds.Count >= options.MaximumRelationshipIdsPerNavigation ||
                relationshipId.Length > options.MaximumRelationshipIdCharactersPerNavigation - _relationshipIdCharacters)
                throw new ArgumentException("Book navigation exceeds its relationship identifier limit.", nameof(relationshipId));
            _relationshipIds.Add(relationshipId);
            RelationshipIds.Add(relationshipId);
            _relationshipIdCharacters += relationshipId.Length;
        }
    }

    private sealed class BookPageProjection {
        private readonly Dictionary<string, List<VisioShape>> _aliases = new(StringComparer.Ordinal);

        public BookPageProjection(VisualArtifactInterchangeEnvelope envelope, OfficeVisioVisualConversionResult result) {
            Result = result;
            foreach (VisioShape shape in result.Page.Shapes) {
                AddAlias(shape.Id, shape);
                AddAlias(shape.GetShapeDataValue("CFX.Id"), shape);
            }
            foreach (VisualArtifactInterchangeNode node in envelope.Nodes) {
                if (!_aliases.TryGetValue(node.Id, out List<VisioShape>? shapes)) continue;
                if (node.Extensions.TryGetValue("chartforgex.sourceId", out string? sourceId)) {
                    foreach (VisioShape shape in shapes.ToArray()) AddAlias(sourceId, shape);
                }
            }
        }

        public OfficeVisioVisualConversionResult Result { get; }

        public VisioShape Resolve(string entityId, string role, int linkIndex) {
            if (!_aliases.TryGetValue(entityId, out List<VisioShape>? matches) || matches.Count == 0) {
                throw new ArgumentException($"Link {linkIndex + 1} {role} entity '{entityId}' was not projected on page '{Result.Page.Name}'.", "links");
            }
            if (matches.Count != 1) {
                throw new ArgumentException($"Link {linkIndex + 1} {role} entity '{entityId}' is ambiguous on page '{Result.Page.Name}'.", "links");
            }
            return matches[0];
        }

        private void AddAlias(string? alias, VisioShape shape) {
            if (string.IsNullOrWhiteSpace(alias)) return;
            if (!_aliases.TryGetValue(alias!, out List<VisioShape>? matches)) {
                matches = new List<VisioShape>();
                _aliases.Add(alias!, matches);
            }
            if (!matches.Contains(shape)) matches.Add(shape);
        }
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
