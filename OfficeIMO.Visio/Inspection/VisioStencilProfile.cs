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
            IReadOnlyList<string> semanticKinds,
            IReadOnlyList<string> stencilCatalogs,
            IReadOnlyList<string> stencilCategories,
            IReadOnlyList<string> stencilSourcePackagePaths,
            IReadOnlyList<string> stencilTags) {
            TotalShapes = totalShapes;
            ConnectorCount = connectorCount;
            Usages = usages;
            ShapeDataKeys = shapeDataKeys;
            ConnectorShapeDataKeys = connectorShapeDataKeys;
            SemanticKinds = semanticKinds;
            StencilCatalogs = stencilCatalogs;
            StencilCategories = stencilCategories;
            StencilSourcePackagePaths = stencilSourcePackagePaths;
            StencilTags = stencilTags;
        }

        /// <summary>Total number of inspected shapes.</summary>
        public int TotalShapes { get; }

        /// <summary>Total number of inspected connectors.</summary>
        public int ConnectorCount { get; }

        /// <summary>Total number of connection points exposed by inspected shapes.</summary>
        public int TotalConnectionPoints => Usages.Sum(usage => usage.ConnectionPointCount);

        /// <summary>Number of inspected shapes that expose at least one connection point.</summary>
        public int ConnectionPointShapeCount => Usages.Sum(usage => usage.ConnectionPointShapeCount);

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

        /// <summary>Number of inspected shapes that carry OfficeIMO stencil identity metadata.</summary>
        public int StencilBackedShapeCount => Usages
            .Where(usage => !string.IsNullOrWhiteSpace(usage.StencilId))
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

        /// <summary>Distinct stencil catalog names represented by inspected shapes.</summary>
        public IReadOnlyList<string> StencilCatalogs { get; }

        /// <summary>Distinct stencil categories represented by inspected shapes.</summary>
        public IReadOnlyList<string> StencilCategories { get; }

        /// <summary>Distinct source package paths represented by inspected shapes.</summary>
        public IReadOnlyList<string> StencilSourcePackagePaths { get; }

        /// <summary>Distinct stencil tags represented by inspected shapes.</summary>
        public IReadOnlyList<string> StencilTags { get; }

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
                CollectSemanticKinds(shapes).AsReadOnly(),
                CollectUsageValues(usages, usage => usage.StencilCatalogName).AsReadOnly(),
                CollectUsageValues(usages, usage => usage.StencilCategory).AsReadOnly(),
                CollectUsageValues(usages, usage => usage.StencilSourcePackagePath).AsReadOnly(),
                CollectUsageListValues(usages, usage => usage.StencilTags).AsReadOnly());
        }

        /// <summary>
        /// Writes a stable line-oriented representation suitable for profile snapshots and review diffs.
        /// </summary>
        public string ToText() {
            StringBuilder builder = new();
            AppendLine(builder, "profile.totalShapes", TotalShapes);
            AppendLine(builder, "profile.connectorCount", ConnectorCount);
            AppendLine(builder, "profile.totalConnectionPoints", TotalConnectionPoints);
            AppendLine(builder, "profile.connectionPointShapeCount", ConnectionPointShapeCount);
            AppendLine(builder, "profile.masterBackedShapeCount", MasterBackedShapeCount);
            AppendLine(builder, "profile.packageBackedShapeCount", PackageBackedShapeCount);
            AppendLine(builder, "profile.generatedMasterBackedShapeCount", GeneratedMasterBackedShapeCount);
            AppendLine(builder, "profile.basicGeometryShapeCount", BasicGeometryShapeCount);
            AppendLine(builder, "profile.stencilBackedShapeCount", StencilBackedShapeCount);
            AppendLine(builder, "profile.semanticOnlyShapeCount", SemanticOnlyShapeCount);
            AppendLine(builder, "profile.shapeDataKeys", string.Join(",", ShapeDataKeys));
            AppendLine(builder, "profile.connectorShapeDataKeys", string.Join(",", ConnectorShapeDataKeys));
            AppendLine(builder, "profile.semanticKinds", string.Join(",", SemanticKinds));
            AppendLine(builder, "profile.stencilCatalogs", string.Join(",", StencilCatalogs));
            AppendLine(builder, "profile.stencilCategories", string.Join(",", StencilCategories));
            AppendLine(builder, "profile.stencilSourcePackagePaths", string.Join(",", StencilSourcePackagePaths));
            AppendLine(builder, "profile.stencilTags", string.Join(",", StencilTags));
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
                string? stencilId = GetStencilId(shape, master);
                return new VisioStencilUsageKey(
                    kind,
                    !string.IsNullOrWhiteSpace(stencilId) ? "stencil:" + stencilId : "master:" + master.NameU,
                    master.Id,
                    master.NameU,
                    master.ShapeNameU,
                    semanticKind,
                    stencilId,
                    GetStencilName(shape, master),
                    GetStencilCategory(shape, master),
                    GetStencilCatalogName(shape, master),
                    GetStencilSourcePackagePath(shape, master),
                    GetStencilKeywords(shape, master),
                    GetStencilAliases(shape, master),
                    GetStencilTags(shape, master));
            }

            if (!string.IsNullOrWhiteSpace(shape.NameU)) {
                string? stencilId = GetStencilId(shape, null);
                return new VisioStencilUsageKey(
                    VisioStencilProfileUsageKind.BasicGeometry,
                    !string.IsNullOrWhiteSpace(stencilId) ? "stencil:" + stencilId : "geometry:" + shape.NameU,
                    null,
                    null,
                    shape.NameU,
                    semanticKind,
                    stencilId,
                    GetStencilName(shape, null),
                    GetStencilCategory(shape, null),
                    GetStencilCatalogName(shape, null),
                    GetStencilSourcePackagePath(shape, null),
                    GetStencilKeywords(shape, null),
                    GetStencilAliases(shape, null),
                    GetStencilTags(shape, null));
            }

            if (!string.IsNullOrWhiteSpace(semanticKind)) {
                return new VisioStencilUsageKey(
                    VisioStencilProfileUsageKind.SemanticOnly,
                    "semantic:" + semanticKind,
                null,
                null,
                null,
                semanticKind,
                null,
                null,
                null,
                null,
                null,
                Array.Empty<string>(),
                Array.Empty<string>(),
                Array.Empty<string>());
        }

            return new VisioStencilUsageKey(
                VisioStencilProfileUsageKind.BasicGeometry,
                "geometry:unknown",
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                Array.Empty<string>(),
                Array.Empty<string>(),
                Array.Empty<string>());
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

        private static List<string> CollectUsageValues(IEnumerable<VisioStencilUsageProfile> usages, Func<VisioStencilUsageProfile, string?> selector) {
            return usages
                .Select(selector)
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Select(value => value!)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static List<string> CollectUsageListValues(IEnumerable<VisioStencilUsageProfile> usages, Func<VisioStencilUsageProfile, IReadOnlyList<string>> selector) {
            return usages
                .SelectMany(selector)
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static string? GetStencilId(VisioInspectionShapeSnapshot shape, VisioInspectionMasterSnapshot? master) {
            return VisioStencilMetadata.GetUserCellValue(shape.UserCells, VisioSemanticUserCells.StencilId) ?? master?.StencilId;
        }

        private static string? GetStencilName(VisioInspectionShapeSnapshot shape, VisioInspectionMasterSnapshot? master) {
            return VisioStencilMetadata.GetUserCellValue(shape.UserCells, VisioSemanticUserCells.StencilName) ?? master?.StencilName;
        }

        private static string? GetStencilCategory(VisioInspectionShapeSnapshot shape, VisioInspectionMasterSnapshot? master) {
            return VisioStencilMetadata.GetUserCellValue(shape.UserCells, VisioSemanticUserCells.StencilCategory) ?? master?.StencilCategory;
        }

        private static string? GetStencilCatalogName(VisioInspectionShapeSnapshot shape, VisioInspectionMasterSnapshot? master) {
            return VisioStencilMetadata.GetUserCellValue(shape.UserCells, VisioSemanticUserCells.StencilCatalog) ?? master?.StencilCatalogName;
        }

        private static string? GetStencilSourcePackagePath(VisioInspectionShapeSnapshot shape, VisioInspectionMasterSnapshot? master) {
            return VisioStencilMetadata.GetUserCellValue(shape.UserCells, VisioSemanticUserCells.StencilSourcePackagePath) ?? master?.StencilSourcePackagePath;
        }

        private static IReadOnlyList<string> GetStencilKeywords(VisioInspectionShapeSnapshot shape, VisioInspectionMasterSnapshot? master) {
            IReadOnlyList<string> values = VisioStencilMetadata.GetUserCellList(shape.UserCells, VisioSemanticUserCells.StencilKeywords);
            return values.Count > 0 ? values : master?.StencilKeywords ?? Array.Empty<string>();
        }

        private static IReadOnlyList<string> GetStencilAliases(VisioInspectionShapeSnapshot shape, VisioInspectionMasterSnapshot? master) {
            IReadOnlyList<string> values = VisioStencilMetadata.GetUserCellList(shape.UserCells, VisioSemanticUserCells.StencilAliases);
            return values.Count > 0 ? values : master?.StencilAliases ?? Array.Empty<string>();
        }

        private static IReadOnlyList<string> GetStencilTags(VisioInspectionShapeSnapshot shape, VisioInspectionMasterSnapshot? master) {
            IReadOnlyList<string> values = VisioStencilMetadata.GetUserCellList(shape.UserCells, VisioSemanticUserCells.StencilTags);
            return values.Count > 0 ? values : master?.StencilTags ?? Array.Empty<string>();
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
            string? stencilId,
            string? stencilName,
            string? stencilCategory,
            string? stencilCatalogName,
            string? stencilSourcePackagePath,
            IReadOnlyList<string> stencilKeywords,
            IReadOnlyList<string> stencilAliases,
            IReadOnlyList<string> stencilTags,
            int count,
            int connectionPointCount,
            int connectionPointShapeCount,
            IReadOnlyList<string> shapeIds,
            IReadOnlyList<string> pageNames,
            IReadOnlyList<string> shapeDataKeys) {
            Key = key;
            Kind = kind;
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
            Count = count;
            ConnectionPointCount = connectionPointCount;
            ConnectionPointShapeCount = connectionPointShapeCount;
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

        /// <summary>OfficeIMO stencil identifier, when known.</summary>
        public string? StencilId { get; }

        /// <summary>OfficeIMO stencil display name, when known.</summary>
        public string? StencilName { get; }

        /// <summary>OfficeIMO stencil category, when known.</summary>
        public string? StencilCategory { get; }

        /// <summary>Stencil catalog name, when known.</summary>
        public string? StencilCatalogName { get; }

        /// <summary>Source package path, when known.</summary>
        public string? StencilSourcePackagePath { get; }

        /// <summary>Searchable stencil keywords represented by this usage group.</summary>
        public IReadOnlyList<string> StencilKeywords { get; }

        /// <summary>Stencil aliases represented by this usage group.</summary>
        public IReadOnlyList<string> StencilAliases { get; }

        /// <summary>Semantic stencil tags represented by this usage group.</summary>
        public IReadOnlyList<string> StencilTags { get; }

        /// <summary>Number of shapes in this usage group.</summary>
        public int Count { get; }

        /// <summary>Total number of connection points exposed by shapes in this usage group.</summary>
        public int ConnectionPointCount { get; }

        /// <summary>Number of shapes in this usage group that expose at least one connection point.</summary>
        public int ConnectionPointShapeCount { get; }

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
                key.StencilId,
                key.StencilName,
                key.StencilCategory,
                key.StencilCatalogName,
                key.StencilSourcePackagePath,
                key.StencilKeywords,
                key.StencilAliases,
                key.StencilTags,
                shapeList.Count,
                shapeList.Sum(shape => shape.ConnectionPointCount),
                shapeList.Count(shape => shape.ConnectionPointCount > 0),
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
            VisioStencilProfile.AppendLine(builder, prefix + ".stencilId", StencilId);
            VisioStencilProfile.AppendLine(builder, prefix + ".stencilName", StencilName);
            VisioStencilProfile.AppendLine(builder, prefix + ".stencilCategory", StencilCategory);
            VisioStencilProfile.AppendLine(builder, prefix + ".stencilCatalog", StencilCatalogName);
            VisioStencilProfile.AppendLine(builder, prefix + ".stencilSourcePackagePath", StencilSourcePackagePath);
            VisioStencilProfile.AppendLine(builder, prefix + ".stencilKeywords", string.Join(",", StencilKeywords));
            VisioStencilProfile.AppendLine(builder, prefix + ".stencilAliases", string.Join(",", StencilAliases));
            VisioStencilProfile.AppendLine(builder, prefix + ".stencilTags", string.Join(",", StencilTags));
            VisioStencilProfile.AppendLine(builder, prefix + ".count", Count);
            VisioStencilProfile.AppendLine(builder, prefix + ".connectionPointCount", ConnectionPointCount);
            VisioStencilProfile.AppendLine(builder, prefix + ".connectionPointShapeCount", ConnectionPointShapeCount);
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
            string? semanticKind,
            string? stencilId,
            string? stencilName,
            string? stencilCategory,
            string? stencilCatalogName,
            string? stencilSourcePackagePath,
            IReadOnlyList<string> stencilKeywords,
            IReadOnlyList<string> stencilAliases,
            IReadOnlyList<string> stencilTags) {
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
                       SequenceEqual(x.StencilTags, y.StencilTags);
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
