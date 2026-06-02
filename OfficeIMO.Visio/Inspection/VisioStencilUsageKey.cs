using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeIMO.Visio {
    internal sealed class VisioStencilUsageKey {
        public VisioStencilUsageKey(
            VisioStencilProfileUsageKind kind,
            string key,
            string? masterId,
            string? masterNameU,
            string? shapeNameU,
            string? semanticKind,
            string? stencilId,
            string? stencilName,
            string? stencilCategory,
            string? stencilCatalogName,
            string? stencilSourcePackagePath,
            IReadOnlyList<string> stencilKeywords,
            IReadOnlyList<string> stencilAliases,
            IReadOnlyList<string> stencilTags,
            string? stencilIconNameU,
            double? sourceDefaultWidth,
            double? sourceDefaultHeight,
            string? stencilDefaultUnit,
            string? stencilPreviewImageContentType,
            string? stencilPreviewImageExtension) {
            Kind = kind;
            Key = key;
            MasterId = masterId;
            MasterNameU = masterNameU;
            ShapeNameU = shapeNameU;
            SemanticKind = semanticKind;
            StencilId = stencilId;
            StencilName = stencilName;
            StencilCategory = stencilCategory;
            StencilCatalogName = stencilCatalogName;
            StencilSourcePackagePath = stencilSourcePackagePath;
            StencilKeywords = stencilKeywords;
            StencilAliases = stencilAliases;
            StencilTags = stencilTags;
            StencilIconNameU = stencilIconNameU;
            SourceDefaultWidth = sourceDefaultWidth;
            SourceDefaultHeight = sourceDefaultHeight;
            StencilDefaultUnit = stencilDefaultUnit;
            StencilPreviewImageContentType = stencilPreviewImageContentType;
            StencilPreviewImageExtension = stencilPreviewImageExtension;
        }

        public VisioStencilProfileUsageKind Kind { get; }

        public string Key { get; }

        public string? MasterId { get; }

        public string? MasterNameU { get; }

        public string? ShapeNameU { get; }

        public string? SemanticKind { get; }

        public string? StencilId { get; }

        public string? StencilName { get; }

        public string? StencilCategory { get; }

        public string? StencilCatalogName { get; }

        public string? StencilSourcePackagePath { get; }

        public IReadOnlyList<string> StencilKeywords { get; }

        public IReadOnlyList<string> StencilAliases { get; }

        public IReadOnlyList<string> StencilTags { get; }

        public string? StencilIconNameU { get; }

        public double? SourceDefaultWidth { get; }

        public double? SourceDefaultHeight { get; }

        public string? StencilDefaultUnit { get; }

        public string? StencilPreviewImageContentType { get; }

        public string? StencilPreviewImageExtension { get; }

        public static IEqualityComparer<VisioStencilUsageKey> Comparer { get; } = new VisioStencilUsageKeyComparer();

        private sealed class VisioStencilUsageKeyComparer : IEqualityComparer<VisioStencilUsageKey> {
            public bool Equals(VisioStencilUsageKey? x, VisioStencilUsageKey? y) {
                if (ReferenceEquals(x, y)) {
                    return true;
                }

                if (x == null || y == null) {
                    return false;
                }

                return x.Kind == y.Kind &&
                       string.Equals(x.Key, y.Key, StringComparison.OrdinalIgnoreCase) &&
                       string.Equals(x.MasterId, y.MasterId, StringComparison.OrdinalIgnoreCase) &&
                       string.Equals(x.MasterNameU, y.MasterNameU, StringComparison.OrdinalIgnoreCase) &&
                       string.Equals(x.ShapeNameU, y.ShapeNameU, StringComparison.OrdinalIgnoreCase) &&
                       string.Equals(x.SemanticKind, y.SemanticKind, StringComparison.OrdinalIgnoreCase) &&
                       string.Equals(x.StencilId, y.StencilId, StringComparison.OrdinalIgnoreCase) &&
                       string.Equals(x.StencilName, y.StencilName, StringComparison.OrdinalIgnoreCase) &&
                       string.Equals(x.StencilCategory, y.StencilCategory, StringComparison.OrdinalIgnoreCase) &&
                       string.Equals(x.StencilCatalogName, y.StencilCatalogName, StringComparison.OrdinalIgnoreCase) &&
                       string.Equals(x.StencilSourcePackagePath, y.StencilSourcePackagePath, StringComparison.OrdinalIgnoreCase) &&
                       SequenceEqual(x.StencilKeywords, y.StencilKeywords) &&
                       SequenceEqual(x.StencilAliases, y.StencilAliases) &&
                       SequenceEqual(x.StencilTags, y.StencilTags) &&
                       string.Equals(x.StencilIconNameU, y.StencilIconNameU, StringComparison.OrdinalIgnoreCase) &&
                       Nullable.Equals(x.SourceDefaultWidth, y.SourceDefaultWidth) &&
                       Nullable.Equals(x.SourceDefaultHeight, y.SourceDefaultHeight) &&
                       string.Equals(x.StencilDefaultUnit, y.StencilDefaultUnit, StringComparison.OrdinalIgnoreCase) &&
                       string.Equals(x.StencilPreviewImageContentType, y.StencilPreviewImageContentType, StringComparison.OrdinalIgnoreCase) &&
                       string.Equals(x.StencilPreviewImageExtension, y.StencilPreviewImageExtension, StringComparison.OrdinalIgnoreCase);
            }

            public int GetHashCode(VisioStencilUsageKey obj) {
                unchecked {
                    int hash = 17;
                    hash = (hash * 31) + obj.Kind.GetHashCode();
                    hash = (hash * 31) + StringComparer.OrdinalIgnoreCase.GetHashCode(obj.Key);
                    hash = (hash * 31) + (obj.MasterId == null ? 0 : StringComparer.OrdinalIgnoreCase.GetHashCode(obj.MasterId));
                    hash = (hash * 31) + (obj.MasterNameU == null ? 0 : StringComparer.OrdinalIgnoreCase.GetHashCode(obj.MasterNameU));
                    hash = (hash * 31) + (obj.ShapeNameU == null ? 0 : StringComparer.OrdinalIgnoreCase.GetHashCode(obj.ShapeNameU));
                    hash = (hash * 31) + (obj.SemanticKind == null ? 0 : StringComparer.OrdinalIgnoreCase.GetHashCode(obj.SemanticKind));
                    hash = (hash * 31) + (obj.StencilId == null ? 0 : StringComparer.OrdinalIgnoreCase.GetHashCode(obj.StencilId));
                    hash = (hash * 31) + (obj.StencilName == null ? 0 : StringComparer.OrdinalIgnoreCase.GetHashCode(obj.StencilName));
                    hash = (hash * 31) + (obj.StencilCategory == null ? 0 : StringComparer.OrdinalIgnoreCase.GetHashCode(obj.StencilCategory));
                    hash = (hash * 31) + (obj.StencilCatalogName == null ? 0 : StringComparer.OrdinalIgnoreCase.GetHashCode(obj.StencilCatalogName));
                    hash = (hash * 31) + (obj.StencilSourcePackagePath == null ? 0 : StringComparer.OrdinalIgnoreCase.GetHashCode(obj.StencilSourcePackagePath));
                    hash = AddListHash(hash, obj.StencilKeywords);
                    hash = AddListHash(hash, obj.StencilAliases);
                    hash = AddListHash(hash, obj.StencilTags);
                    hash = (hash * 31) + (obj.StencilIconNameU == null ? 0 : StringComparer.OrdinalIgnoreCase.GetHashCode(obj.StencilIconNameU));
                    hash = (hash * 31) + (obj.SourceDefaultWidth?.GetHashCode() ?? 0);
                    hash = (hash * 31) + (obj.SourceDefaultHeight?.GetHashCode() ?? 0);
                    hash = (hash * 31) + (obj.StencilDefaultUnit == null ? 0 : StringComparer.OrdinalIgnoreCase.GetHashCode(obj.StencilDefaultUnit));
                    hash = (hash * 31) + (obj.StencilPreviewImageContentType == null ? 0 : StringComparer.OrdinalIgnoreCase.GetHashCode(obj.StencilPreviewImageContentType));
                    hash = (hash * 31) + (obj.StencilPreviewImageExtension == null ? 0 : StringComparer.OrdinalIgnoreCase.GetHashCode(obj.StencilPreviewImageExtension));
                    return hash;
                }
            }

            private static bool SequenceEqual(IReadOnlyList<string> left, IReadOnlyList<string> right) {
                if (left.Count != right.Count) {
                    return false;
                }

                for (int i = 0; i < left.Count; i++) {
                    if (!string.Equals(left[i], right[i], StringComparison.OrdinalIgnoreCase)) {
                        return false;
                    }
                }

                return true;
            }

            private static int AddListHash(int seed, IReadOnlyList<string> values) {
                int hash = seed;
                foreach (string value in values) {
                    hash = (hash * 31) + StringComparer.OrdinalIgnoreCase.GetHashCode(value);
                }

                return hash;
            }
        }
    }
}
