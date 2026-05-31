using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Creates stencil usage profiles from Visio documents and inspection snapshots.
    /// </summary>
    public static class VisioStencilProfileExtensions {
        /// <summary>
        /// Creates a deterministic stencil usage profile for the document.
        /// </summary>
        public static VisioStencilProfile CreateStencilProfile(this VisioDocument document) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            return document.CreateInspectionSnapshot().CreateStencilProfile();
        }

        /// <summary>
        /// Creates a deterministic stencil usage profile from an inspection snapshot.
        /// </summary>
        public static VisioStencilProfile CreateStencilProfile(this VisioInspectionSnapshot snapshot) {
            if (snapshot == null) {
                throw new ArgumentNullException(nameof(snapshot));
            }

            return VisioStencilProfile.FromSnapshot(snapshot);
        }
    }

    /// <summary>
    /// Aggregated view of which stencil masters, semantic shape kinds, and metadata keys a diagram uses.
    /// </summary>
    public sealed class VisioStencilProfile {
        private VisioStencilProfile(
            int totalShapes,
            int connectorCount,
            IReadOnlyList<VisioStencilUsageProfile> usages,
            IReadOnlyList<string> shapeDataKeys,
            IReadOnlyList<string> connectorShapeDataKeys,
            IReadOnlyList<string> semanticKinds) {
            TotalShapes = totalShapes;
            ConnectorCount = connectorCount;
            Usages = usages;
            ShapeDataKeys = shapeDataKeys;
            ConnectorShapeDataKeys = connectorShapeDataKeys;
            SemanticKinds = semanticKinds;
        }

        /// <summary>Total number of inspected shapes.</summary>
        public int TotalShapes { get; }

        /// <summary>Total number of inspected connectors.</summary>
        public int ConnectorCount { get; }

        /// <summary>Number of shapes backed by any registered master.</summary>
        public int MasterBackedShapeCount => Usages
            .Where(usage => usage.Kind == VisioStencilProfileUsageKind.PackageBackedMaster ||
                            usage.Kind == VisioStencilProfileUsageKind.GeneratedMaster)
            .Sum(usage => usage.Count);

        /// <summary>Number of shapes backed by masters imported from a stencil package or document package.</summary>
        public int PackageBackedShapeCount => Usages
            .Where(usage => usage.Kind == VisioStencilProfileUsageKind.PackageBackedMaster)
            .Sum(usage => usage.Count);

        /// <summary>Number of shapes backed by generated OfficeIMO masters.</summary>
        public int GeneratedMasterBackedShapeCount => Usages
            .Where(usage => usage.Kind == VisioStencilProfileUsageKind.GeneratedMaster)
            .Sum(usage => usage.Count);

        /// <summary>Number of direct geometry shapes not backed by a registered master.</summary>
        public int BasicGeometryShapeCount => Usages
            .Where(usage => usage.Kind == VisioStencilProfileUsageKind.BasicGeometry)
            .Sum(usage => usage.Count);

        /// <summary>Number of shapes grouped only by OfficeIMO semantic kind.</summary>
        public int SemanticOnlyShapeCount => Usages
            .Where(usage => usage.Kind == VisioStencilProfileUsageKind.SemanticOnly)
            .Sum(usage => usage.Count);

        /// <summary>Stencil, master, geometry, and semantic shape usage groups.</summary>
        public IReadOnlyList<VisioStencilUsageProfile> Usages { get; }

        /// <summary>Distinct Shape Data keys used by inspected shapes.</summary>
        public IReadOnlyList<string> ShapeDataKeys { get; }

        /// <summary>Distinct Shape Data keys used by inspected connectors.</summary>
        public IReadOnlyList<string> ConnectorShapeDataKeys { get; }

        /// <summary>Distinct OfficeIMO semantic kind values used by inspected shapes.</summary>
        public IReadOnlyList<string> SemanticKinds { get; }

        /// <summary>
        /// Creates a stencil profile from an inspection snapshot.
        /// </summary>
        public static VisioStencilProfile FromSnapshot(VisioInspectionSnapshot snapshot) {
            if (snapshot == null) {
                throw new ArgumentNullException(nameof(snapshot));
            }

            Dictionary<string, VisioInspectionMasterSnapshot> masters = snapshot.Masters
                .GroupBy(master => master.Id, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);
            List<VisioInspectionShapeSnapshot> shapes = snapshot.Pages
                .SelectMany(page => page.Shapes)
                .Where(shape => !IsMasterArtworkChild(shape))
                .ToList();
            List<VisioInspectionConnectorSnapshot> connectors = snapshot.Pages
                .SelectMany(page => page.Connectors)
                .ToList();
            List<VisioStencilUsageProfile> usages = shapes
                .GroupBy(shape => CreateUsageKey(shape, masters), VisioStencilUsageKey.Comparer)
                .Select(group => VisioStencilUsageProfile.FromShapes(group.Key, group, snapshot.Pages))
                .OrderBy(usage => usage.Kind)
                .ThenBy(usage => usage.Key, StringComparer.OrdinalIgnoreCase)
                .ToList();

            return new VisioStencilProfile(
                shapes.Count,
                connectors.Count,
                usages.AsReadOnly(),
                CollectShapeDataKeys(shapes).AsReadOnly(),
                CollectConnectorShapeDataKeys(connectors).AsReadOnly(),
                CollectSemanticKinds(shapes).AsReadOnly());
        }

        /// <summary>
        /// Writes a stable line-oriented representation suitable for profile snapshots and review diffs.
        /// </summary>
        public string ToText() {
            StringBuilder builder = new();
            AppendLine(builder, "profile.totalShapes", TotalShapes);
            AppendLine(builder, "profile.connectorCount", ConnectorCount);
            AppendLine(builder, "profile.masterBackedShapeCount", MasterBackedShapeCount);
            AppendLine(builder, "profile.packageBackedShapeCount", PackageBackedShapeCount);
            AppendLine(builder, "profile.generatedMasterBackedShapeCount", GeneratedMasterBackedShapeCount);
            AppendLine(builder, "profile.basicGeometryShapeCount", BasicGeometryShapeCount);
            AppendLine(builder, "profile.semanticOnlyShapeCount", SemanticOnlyShapeCount);
            AppendLine(builder, "profile.shapeDataKeys", string.Join(",", ShapeDataKeys));
            AppendLine(builder, "profile.connectorShapeDataKeys", string.Join(",", ConnectorShapeDataKeys));
            AppendLine(builder, "profile.semanticKinds", string.Join(",", SemanticKinds));
            AppendLine(builder, "profile.usageCount", Usages.Count);

            foreach (VisioStencilUsageProfile usage in Usages) {
                usage.AppendText(builder);
            }

            return builder.ToString();
        }

        /// <inheritdoc />
        public override string ToString() {
            return ToText();
        }

        private static VisioStencilUsageKey CreateUsageKey(
            VisioInspectionShapeSnapshot shape,
            IReadOnlyDictionary<string, VisioInspectionMasterSnapshot> masters) {
            string? semanticKind = GetSemanticKind(shape);
            VisioInspectionMasterSnapshot? master = null;
            if (!string.IsNullOrWhiteSpace(shape.MasterId)) {
                masters.TryGetValue(shape.MasterId!, out master);
            }

            if (master != null) {
                VisioStencilProfileUsageKind kind = master.IsPackageBacked
                    ? VisioStencilProfileUsageKind.PackageBackedMaster
                    : VisioStencilProfileUsageKind.GeneratedMaster;
                return new VisioStencilUsageKey(
                    kind,
                    "master:" + master.NameU,
                    master.Id,
                    master.NameU,
                    master.ShapeNameU,
                    semanticKind);
            }

            if (!string.IsNullOrWhiteSpace(shape.NameU)) {
                return new VisioStencilUsageKey(
                    VisioStencilProfileUsageKind.BasicGeometry,
                    "geometry:" + shape.NameU,
                    null,
                    null,
                    shape.NameU,
                    semanticKind);
            }

            if (!string.IsNullOrWhiteSpace(semanticKind)) {
                return new VisioStencilUsageKey(
                    VisioStencilProfileUsageKind.SemanticOnly,
                    "semantic:" + semanticKind,
                    null,
                    null,
                    null,
                    semanticKind);
            }

            return new VisioStencilUsageKey(
                VisioStencilProfileUsageKind.BasicGeometry,
                "geometry:unknown",
                null,
                null,
                null,
                null);
        }

        internal static string? GetSemanticKind(VisioInspectionShapeSnapshot shape) {
            return shape.UserCells
                .FirstOrDefault(cell => string.Equals(cell.Name, VisioSemanticUserCells.Kind, StringComparison.OrdinalIgnoreCase))
                ?.Value;
        }

        private static bool IsMasterArtworkChild(VisioInspectionShapeSnapshot shape) {
            return !string.IsNullOrWhiteSpace(shape.ParentId) &&
                   !string.IsNullOrWhiteSpace(shape.MasterShapeId);
        }

        internal static void AppendLine(StringBuilder builder, string key, object? value) {
            builder.Append(key);
            builder.Append('=');
            builder.Append(VisioInspectionSnapshot.FormatValue(value));
            builder.AppendLine();
        }

        private static List<string> CollectShapeDataKeys(IEnumerable<VisioInspectionShapeSnapshot> shapes) {
            return shapes
                .SelectMany(shape => shape.ShapeData.Select(row => row.Name))
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static List<string> CollectConnectorShapeDataKeys(IEnumerable<VisioInspectionConnectorSnapshot> connectors) {
            return connectors
                .SelectMany(connector => connector.ShapeData.Select(row => row.Name))
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static List<string> CollectSemanticKinds(IEnumerable<VisioInspectionShapeSnapshot> shapes) {
            return shapes
                .Select(GetSemanticKind)
                .Where(kind => !string.IsNullOrWhiteSpace(kind))
                .Select(kind => kind!)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(kind => kind, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }
    }

    /// <summary>
    /// One stencil, master, geometry, or semantic shape usage group in a stencil profile.
    /// </summary>
    public sealed class VisioStencilUsageProfile {
        private VisioStencilUsageProfile(
            string key,
            VisioStencilProfileUsageKind kind,
            string? masterId,
            string? masterNameU,
            string? shapeNameU,
            string? semanticKind,
            int count,
            IReadOnlyList<string> shapeIds,
            IReadOnlyList<string> pageNames,
            IReadOnlyList<string> shapeDataKeys) {
            Key = key;
            Kind = kind;
            MasterId = masterId;
            MasterNameU = masterNameU;
            ShapeNameU = shapeNameU;
            SemanticKind = semanticKind;
            Count = count;
            ShapeIds = shapeIds;
            PageNames = pageNames;
            ShapeDataKeys = shapeDataKeys;
        }

        /// <summary>Stable usage key.</summary>
        public string Key { get; }

        /// <summary>Usage classification.</summary>
        public VisioStencilProfileUsageKind Kind { get; }

        /// <summary>Referenced master identifier, when available.</summary>
        public string? MasterId { get; }

        /// <summary>Referenced master universal name, when available.</summary>
        public string? MasterNameU { get; }

        /// <summary>Shape universal name used by the grouped shapes.</summary>
        public string? ShapeNameU { get; }

        /// <summary>OfficeIMO semantic kind assigned to the grouped shapes, when consistent.</summary>
        public string? SemanticKind { get; }

        /// <summary>Number of shapes in this usage group.</summary>
        public int Count { get; }

        /// <summary>Shape identifiers included in this usage group.</summary>
        public IReadOnlyList<string> ShapeIds { get; }

        /// <summary>Page names where this usage appears.</summary>
        public IReadOnlyList<string> PageNames { get; }

        /// <summary>Distinct Shape Data keys used by shapes in this group.</summary>
        public IReadOnlyList<string> ShapeDataKeys { get; }

        internal static VisioStencilUsageProfile FromShapes(
            VisioStencilUsageKey key,
            IEnumerable<VisioInspectionShapeSnapshot> shapes,
            IReadOnlyList<VisioInspectionPageSnapshot> pages) {
            List<VisioInspectionShapeSnapshot> shapeList = shapes.ToList();
            Dictionary<string, string> pageByShapeId = pages
                .SelectMany(page => page.Shapes.Select(shape => new { shape.Id, Page = page.Name }))
                .GroupBy(item => item.Id, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First().Page, StringComparer.OrdinalIgnoreCase);

            IReadOnlyList<string> shapeIds = shapeList
                .Select(shape => shape.Id)
                .OrderBy(id => id, StringComparer.OrdinalIgnoreCase)
                .ToList()
                .AsReadOnly();
            IReadOnlyList<string> pageNames = shapeList
                .Select(shape => pageByShapeId.TryGetValue(shape.Id, out string? pageName) ? pageName : string.Empty)
                .Where(pageName => !string.IsNullOrWhiteSpace(pageName))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(pageName => pageName, StringComparer.OrdinalIgnoreCase)
                .ToList()
                .AsReadOnly();
            IReadOnlyList<string> shapeDataKeys = shapeList
                .SelectMany(shape => shape.ShapeData.Select(row => row.Name))
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                .ToList()
                .AsReadOnly();

            return new VisioStencilUsageProfile(
                key.Key,
                key.Kind,
                key.MasterId,
                key.MasterNameU,
                key.ShapeNameU,
                key.SemanticKind,
                shapeList.Count,
                shapeIds,
                pageNames,
                shapeDataKeys);
        }

        internal void AppendText(StringBuilder builder) {
            string prefix = "usage[" + VisioInspectionSnapshot.EscapeKey(Key) + "]";
            VisioStencilProfile.AppendLine(builder, prefix + ".kind", Kind);
            VisioStencilProfile.AppendLine(builder, prefix + ".masterId", MasterId);
            VisioStencilProfile.AppendLine(builder, prefix + ".masterNameU", MasterNameU);
            VisioStencilProfile.AppendLine(builder, prefix + ".shapeNameU", ShapeNameU);
            VisioStencilProfile.AppendLine(builder, prefix + ".semanticKind", SemanticKind);
            VisioStencilProfile.AppendLine(builder, prefix + ".count", Count);
            VisioStencilProfile.AppendLine(builder, prefix + ".shapeIds", string.Join(",", ShapeIds));
            VisioStencilProfile.AppendLine(builder, prefix + ".pages", string.Join(",", PageNames));
            VisioStencilProfile.AppendLine(builder, prefix + ".shapeDataKeys", string.Join(",", ShapeDataKeys));
        }
    }

    /// <summary>
    /// Classification for a stencil profile usage group.
    /// </summary>
    public enum VisioStencilProfileUsageKind {
        /// <summary>Shape uses a master imported from a stencil package or document package.</summary>
        PackageBackedMaster = 0,

        /// <summary>Shape uses a generated OfficeIMO master.</summary>
        GeneratedMaster = 1,

        /// <summary>Shape is direct geometry rather than a registered master instance.</summary>
        BasicGeometry = 2,

        /// <summary>Shape has no useful geometry or master identity and is grouped by semantic kind.</summary>
        SemanticOnly = 3
    }

    internal sealed class VisioStencilUsageKey {
        public VisioStencilUsageKey(
            VisioStencilProfileUsageKind kind,
            string key,
            string? masterId,
            string? masterNameU,
            string? shapeNameU,
            string? semanticKind) {
            Kind = kind;
            Key = key;
            MasterId = masterId;
            MasterNameU = masterNameU;
            ShapeNameU = shapeNameU;
            SemanticKind = semanticKind;
        }

        public VisioStencilProfileUsageKind Kind { get; }

        public string Key { get; }

        public string? MasterId { get; }

        public string? MasterNameU { get; }

        public string? ShapeNameU { get; }

        public string? SemanticKind { get; }

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
                       string.Equals(x.SemanticKind, y.SemanticKind, StringComparison.OrdinalIgnoreCase);
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
                    return hash;
                }
            }
        }
    }
}
