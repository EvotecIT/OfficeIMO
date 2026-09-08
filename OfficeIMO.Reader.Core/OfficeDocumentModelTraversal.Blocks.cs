using System;
using System.Collections.Generic;

namespace OfficeIMO.Reader;

internal static partial class OfficeDocumentModelTraversal {
    /// <summary>Reconciles aggregate block projections with inspected page blocks without changing either source.</summary>
    /// <remarks>Callers control enumeration and inspection limits; this index never enumerates the document.</remarks>
    internal sealed class BlockProjectionIndex {
        private readonly Dictionary<OfficeDocumentBlock, ReaderLocation> _locations =
            new(ReferenceIdentityComparer<OfficeDocumentBlock>.Instance);
        private readonly Dictionary<string, List<ReaderLocation>> _identityLocations = new(StringComparer.Ordinal);
        private readonly Dictionary<ReaderLocation, OfficeDocumentRegion?> _regions = new();
        internal bool IsComplete { get; set; } = true;

        private readonly List<OfficeDocumentPage> _inspectedPages = new();
        internal IReadOnlyList<OfficeDocumentPage> InspectedPages => _inspectedPages;
        internal void RegisterPage(OfficeDocumentPage page) => _inspectedPages.Add(page);

        internal void Add(OfficeDocumentBlock block, OfficeDocumentPage page) {
            ReaderLocation fallback = BuildPageLocation(page);
            // Explicit block container coordinates remain authoritative.
            if (block.Location?.Slide.HasValue == true || !string.IsNullOrWhiteSpace(block.Location?.Sheet)) fallback.Page = null;
            ReaderLocation location = MergeLocation(block.Location, fallback, null);
            _regions.Add(location, block.Region);
            if (!_locations.ContainsKey(block)) _locations.Add(block, location);
            if (string.IsNullOrWhiteSpace(block.Id) && string.IsNullOrWhiteSpace(block.Location?.BlockAnchor)) return;
            string identity = BuildBlockProjectionKey(block);
            if (!_identityLocations.TryGetValue(identity, out var matches)) _identityLocations.Add(identity, matches = new());
            matches.Add(location);
        }

        internal ReaderLocation? ResolveLocation(OfficeDocumentBlock block) {
            if (_locations.TryGetValue(block, out ReaderLocation? location)) return location;
            if (!IsComplete) return block.Location;
            if (_identityLocations.TryGetValue(BuildBlockProjectionKey(block), out var matches)) {
                // Missing container information must not select an arbitrary page that reused an anchor.
                var compatible = matches.Where(candidate => block.Location == null || SameContainerWhenKnown(block.Location, candidate))
                    .GroupBy(candidate => BuildBlockIdentity(block, candidate), StringComparer.Ordinal).Take(2).ToArray();
                if (compatible.Length == 1) return MergeLocation(block.Location, compatible[0].First(), null);
            }
            return block.Location;
        }

        // An unscoped aggregate spanning several source containers cannot supply a truthful
        // page location or region. Its page fragments retain the observed text and geometry.
        internal bool HasPageFragments(OfficeDocumentBlock block) {
            // Local anchors may be reused by unrelated containers; only a stable source ID
            // establishes that page blocks are fragments replacing this aggregate observation.
            if (!IsComplete || string.IsNullOrWhiteSpace(block.Id) || _locations.ContainsKey(block)
                || !_identityLocations.TryGetValue(BuildBlockProjectionKey(block), out var matches)) return false;
            return matches.Where(candidate => block.Location == null || SameContainerWhenKnown(block.Location, candidate))
                .Select(candidate => BuildBlockIdentity(block, candidate)).Distinct(StringComparer.Ordinal).Take(2).Count() > 1;
        }

        internal OfficeDocumentRegion? ResolveRegion(OfficeDocumentBlock block) {
            if (block.Region != null || !IsComplete
                || !_identityLocations.TryGetValue(BuildBlockProjectionKey(block), out var matches)) return block.Region;
            var compatible = matches.Where(candidate => block.Location == null || SameContainerWhenKnown(block.Location, candidate))
                .GroupBy(candidate => BuildBlockIdentity(block, candidate), StringComparer.Ordinal).Take(2).ToArray();
            if (compatible.Length != 1) return null;
            var regions = compatible[0].Select(location => _regions[location]).Where(region => region != null)
                .GroupBy(region => (region!.X, region.Y, region.Width, region.Height)).Take(2).ToArray();
            return regions.Length == 1 ? regions[0].First() : null;
        }
    }
}
