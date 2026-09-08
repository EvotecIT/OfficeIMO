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

        private readonly List<OfficeDocumentPage> _inspectedPages = new();
        internal IReadOnlyList<OfficeDocumentPage> InspectedPages => _inspectedPages;
        internal void RegisterPage(OfficeDocumentPage page) => _inspectedPages.Add(page);

        internal void Add(OfficeDocumentBlock block, OfficeDocumentPage page) {
            ReaderLocation fallback = BuildPageLocation(page);
            // Explicit block container coordinates remain authoritative.
            if (block.Location?.Slide.HasValue == true || !string.IsNullOrWhiteSpace(block.Location?.Sheet)) fallback.Page = null;
            ReaderLocation location = MergeLocation(block.Location, fallback, null);
            if (!_locations.ContainsKey(block)) _locations.Add(block, location);
            if (string.IsNullOrWhiteSpace(block.Id) && string.IsNullOrWhiteSpace(block.Location?.BlockAnchor)) return;
            string identity = BuildBlockProjectionKey(block);
            if (!_identityLocations.TryGetValue(identity, out var matches)) _identityLocations.Add(identity, matches = new());
            matches.Add(location);
        }

        internal ReaderLocation? ResolveLocation(OfficeDocumentBlock block) {
            if (_locations.TryGetValue(block, out ReaderLocation? location)) return location;
            if (_identityLocations.TryGetValue(BuildBlockProjectionKey(block), out var matches)) {
                // Missing container information must not select an arbitrary page that reused an anchor.
                var compatible = matches.Where(candidate => block.Location == null || SameContainerWhenKnown(block.Location, candidate))
                    .GroupBy(candidate => BuildBlockIdentity(block, candidate), StringComparer.Ordinal).Take(2).ToArray();
                if (compatible.Length == 1) return MergeLocation(block.Location, compatible[0].First(), null);
            }
            return block.Location;
        }
    }
}
