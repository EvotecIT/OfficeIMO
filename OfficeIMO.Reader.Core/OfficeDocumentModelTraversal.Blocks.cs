using System;
using System.Collections.Generic;
using System.Globalization;

namespace OfficeIMO.Reader;

internal static partial class OfficeDocumentModelTraversal {
    /// <summary>Reconciles aggregate block projections with inspected page blocks without changing either source.</summary>
    /// <remarks>Callers control enumeration and inspection limits; this index never enumerates the document.</remarks>
    internal sealed class BlockProjectionIndex {
        private readonly Dictionary<OfficeDocumentBlock, ReaderLocation> _locations =
            new(ReferenceIdentityComparer<OfficeDocumentBlock>.Instance);
        private readonly Dictionary<string, List<ReaderLocation>> _identityLocations = new(StringComparer.Ordinal);

        internal void Add(OfficeDocumentBlock block, OfficeDocumentPage page) {
            ReaderLocation fallback = BuildPageLocation(page);
            // A sheet or slide number denotes that container, not a PDF-style page number.
            fallback.Page = page.Location?.Page;
            string? kind = page.Location?.SourceBlockKind?.Trim();
            if (string.Equals(kind, "sheet", StringComparison.OrdinalIgnoreCase) && string.IsNullOrWhiteSpace(fallback.Sheet))
                fallback.Sheet = !string.IsNullOrWhiteSpace(page.Name) ? page.Name
                    : page.Number > 0 ? "Sheet " + page.Number.Value.ToString(CultureInfo.InvariantCulture) : null;
            if (!fallback.Slide.HasValue && string.IsNullOrWhiteSpace(fallback.Sheet)) {
                int? number = page.Number > 0 ? page.Number : fallback.Page;
                if (string.Equals(kind, "slide", StringComparison.OrdinalIgnoreCase)) {
                    fallback.Slide = number;
                    fallback.Page = null;
                } else {
                    fallback.Page = number;
                }
            }
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
