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
            IReadOnlyList<string> stencilKeywords,
            IReadOnlyList<string> stencilAliases,
            IReadOnlyList<string> stencilTags,
            IReadOnlyList<string> stencilIconNameUs,
            IReadOnlyList<string> stencilDefaultUnits,
            IReadOnlyList<string> stencilPreviewImageContentTypes,
            IReadOnlyList<string> stencilPreviewImageExtensions,
            IReadOnlyList<VisioStencilFamilyProfile> stencilFamilies) {
            TotalShapes = totalShapes;
            ConnectorCount = connectorCount;
            Usages = usages;
            ShapeDataKeys = shapeDataKeys;
            ConnectorShapeDataKeys = connectorShapeDataKeys;
            SemanticKinds = semanticKinds;
            StencilCatalogs = stencilCatalogs;
            StencilCategories = stencilCategories;
            StencilSourcePackagePaths = stencilSourcePackagePaths;
            StencilKeywords = stencilKeywords;
            StencilAliases = stencilAliases;
            StencilTags = stencilTags;
            StencilIconNameUs = stencilIconNameUs;
            StencilDefaultUnits = stencilDefaultUnits;
            StencilPreviewImageContentTypes = stencilPreviewImageContentTypes;
            StencilPreviewImageExtensions = stencilPreviewImageExtensions;
            StencilFamilies = stencilFamilies;
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

        /// <summary>Stencil family rollups grouped by catalog/category metadata.</summary>
        public IReadOnlyList<VisioStencilFamilyProfile> StencilFamilies { get; }

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

        /// <summary>Distinct stencil keywords represented by inspected shapes.</summary>
        public IReadOnlyList<string> StencilKeywords { get; }

        /// <summary>Distinct stencil aliases represented by inspected shapes.</summary>
        public IReadOnlyList<string> StencilAliases { get; }

        /// <summary>Distinct stencil tags represented by inspected shapes.</summary>
        public IReadOnlyList<string> StencilTags { get; }

        /// <summary>Distinct stencil icon master universal names represented by inspected shapes.</summary>
        public IReadOnlyList<string> StencilIconNameUs { get; }

        /// <summary>Distinct source default-size units represented by inspected shapes.</summary>
        public IReadOnlyList<string> StencilDefaultUnits { get; }

        /// <summary>Distinct stencil preview image content types represented by inspected shapes.</summary>
        public IReadOnlyList<string> StencilPreviewImageContentTypes { get; }

        /// <summary>Distinct stencil preview image extensions represented by inspected shapes.</summary>
        public IReadOnlyList<string> StencilPreviewImageExtensions { get; }

        /// <summary>
        /// Creates a stencil profile from an inspection snapshot.
        /// </summary>
        public static VisioStencilProfile FromSnapshot(VisioInspectionSnapshot snapshot) {
            if (snapshot == null) {
                throw new ArgumentNullException(nameof(snapshot));
            }

            Dictionary<string, IReadOnlyList<VisioInspectionMasterSnapshot>> masters = snapshot.Masters
                .GroupBy(master => master.Id, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => (IReadOnlyList<VisioInspectionMasterSnapshot>)group.ToList().AsReadOnly(), StringComparer.OrdinalIgnoreCase);
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
            DisambiguateUsageSnapshotKeys(usages);
            List<VisioStencilFamilyProfile> stencilFamilies = VisioStencilFamilyProfile.FromUsages(usages);

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
                CollectUsageListValues(usages, usage => usage.StencilKeywords).AsReadOnly(),
                CollectUsageListValues(usages, usage => usage.StencilAliases).AsReadOnly(),
                CollectUsageListValues(usages, usage => usage.StencilTags).AsReadOnly(),
                CollectUsageValues(usages, usage => usage.StencilIconNameU).AsReadOnly(),
                CollectUsageValues(usages, usage => usage.StencilDefaultUnit).AsReadOnly(),
                CollectUsageValues(usages, usage => usage.StencilPreviewImageContentType).AsReadOnly(),
                CollectUsageValues(usages, usage => usage.StencilPreviewImageExtension).AsReadOnly(),
                stencilFamilies.AsReadOnly());
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
            AppendLine(builder, "profile.stencilKeywords", string.Join(",", StencilKeywords));
            AppendLine(builder, "profile.stencilAliases", string.Join(",", StencilAliases));
            AppendLine(builder, "profile.stencilTags", string.Join(",", StencilTags));
            AppendLine(builder, "profile.stencilIconNameUs", string.Join(",", StencilIconNameUs));
            AppendLine(builder, "profile.stencilDefaultUnits", string.Join(",", StencilDefaultUnits));
            AppendLine(builder, "profile.stencilPreviewImageContentTypes", string.Join(",", StencilPreviewImageContentTypes));
            AppendLine(builder, "profile.stencilPreviewImageExtensions", string.Join(",", StencilPreviewImageExtensions));
            AppendLine(builder, "profile.stencilFamilyCount", StencilFamilies.Count);
            AppendLine(builder, "profile.usageCount", Usages.Count);

            foreach (VisioStencilFamilyProfile family in StencilFamilies) {
                family.AppendText(builder);
            }

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
            IReadOnlyDictionary<string, IReadOnlyList<VisioInspectionMasterSnapshot>> masters) {
            string? semanticKind = GetSemanticKind(shape);
            VisioInspectionMasterSnapshot? master = ResolveMaster(shape, masters);

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
                    GetStencilTags(shape, master),
                    GetStencilIconNameU(shape, master),
                    GetStencilDefaultWidth(shape, master),
                    GetStencilDefaultHeight(shape, master),
                    GetStencilDefaultUnit(shape, master),
                    GetStencilPreviewImageContentType(shape, master),
                    GetStencilPreviewImageExtension(shape, master));
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
                    GetStencilTags(shape, null),
                    GetStencilIconNameU(shape, null),
                    GetStencilDefaultWidth(shape, null),
                    GetStencilDefaultHeight(shape, null),
                    GetStencilDefaultUnit(shape, null),
                    GetStencilPreviewImageContentType(shape, null),
                    GetStencilPreviewImageExtension(shape, null));
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
                Array.Empty<string>(),
                null,
                null,
                null,
                null,
                null,
                null);
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
                Array.Empty<string>(),
                null,
                null,
                null,
                null,
                null,
                null);
        }

        private static VisioInspectionMasterSnapshot? ResolveMaster(
            VisioInspectionShapeSnapshot shape,
            IReadOnlyDictionary<string, IReadOnlyList<VisioInspectionMasterSnapshot>> masters) {
            if (string.IsNullOrWhiteSpace(shape.MasterId) ||
                !masters.TryGetValue(shape.MasterId!, out IReadOnlyList<VisioInspectionMasterSnapshot>? candidates) ||
                candidates.Count == 0) {
                return null;
            }

            if (candidates.Count == 1) {
                return candidates[0];
            }

            string? sourcePackagePath = GetStencilSourcePackagePath(shape, null);
            VisioInspectionMasterSnapshot? match = candidates.FirstOrDefault(master =>
                string.Equals(master.NameU, shape.MasterNameU, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(master.StencilSourcePackagePath, sourcePackagePath, StringComparison.OrdinalIgnoreCase));
            if (match != null) {
                return match;
            }

            match = candidates.FirstOrDefault(master =>
                string.Equals(master.NameU, shape.MasterNameU, StringComparison.OrdinalIgnoreCase));
            return match ?? candidates[0];
        }

        internal static string? GetSemanticKind(VisioInspectionShapeSnapshot shape) {
            return shape.UserCells
                .FirstOrDefault(cell => string.Equals(cell.Name, VisioSemanticUserCells.Kind, StringComparison.OrdinalIgnoreCase))
                ?.Value;
        }

        private static void DisambiguateUsageSnapshotKeys(IReadOnlyList<VisioStencilUsageProfile> usages) {
            foreach (IGrouping<string, VisioStencilUsageProfile> group in usages.GroupBy(usage => usage.Key, StringComparer.OrdinalIgnoreCase)) {
                if (group.Count() == 1) {
                    continue;
                }

                foreach (VisioStencilUsageProfile usage in group) {
                    usage.SnapshotKey = usage.Kind + ":" + (usage.MasterId ?? usage.MasterNameU ?? usage.ShapeNameU ?? usage.SemanticKind ?? "none") + ":" + usage.Key;
                }
            }
        }

        private static bool IsMasterArtworkChild(VisioInspectionShapeSnapshot shape) {
            return !string.IsNullOrWhiteSpace(shape.ParentId) &&
                   !string.IsNullOrWhiteSpace(shape.MasterShapeId);
        }

        internal static void AppendLine(StringBuilder builder, string key, object? value) {
            builder.Append(key);
            builder.Append('=');
            builder.Append(VisioInspectionSnapshot.FormatLineValue(value));
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

        private static string? GetStencilIconNameU(VisioInspectionShapeSnapshot shape, VisioInspectionMasterSnapshot? master) {
            return VisioStencilMetadata.GetUserCellValue(shape.UserCells, VisioSemanticUserCells.StencilIconNameU) ?? master?.StencilIconNameU;
        }

        private static double? GetStencilDefaultWidth(VisioInspectionShapeSnapshot shape, VisioInspectionMasterSnapshot? master) {
            return GetUserCellDouble(shape, VisioSemanticUserCells.StencilDefaultWidth) ?? master?.StencilDefaultWidth;
        }

        private static double? GetStencilDefaultHeight(VisioInspectionShapeSnapshot shape, VisioInspectionMasterSnapshot? master) {
            return GetUserCellDouble(shape, VisioSemanticUserCells.StencilDefaultHeight) ?? master?.StencilDefaultHeight;
        }

        private static string? GetStencilDefaultUnit(VisioInspectionShapeSnapshot shape, VisioInspectionMasterSnapshot? master) {
            return VisioStencilMetadata.GetUserCellValue(shape.UserCells, VisioSemanticUserCells.StencilDefaultUnit) ?? master?.StencilDefaultUnit;
        }

        private static string? GetStencilPreviewImageContentType(VisioInspectionShapeSnapshot shape, VisioInspectionMasterSnapshot? master) {
            return VisioStencilMetadata.GetUserCellValue(shape.UserCells, VisioSemanticUserCells.StencilPreviewImageContentType) ?? master?.StencilPreviewImageContentType;
        }

        private static string? GetStencilPreviewImageExtension(VisioInspectionShapeSnapshot shape, VisioInspectionMasterSnapshot? master) {
            return VisioStencilMetadata.GetUserCellValue(shape.UserCells, VisioSemanticUserCells.StencilPreviewImageExtension) ?? master?.StencilPreviewImageExtension;
        }

        private static double? GetUserCellDouble(VisioInspectionShapeSnapshot shape, string name) {
            string? value = VisioStencilMetadata.GetUserCellValue(shape.UserCells, name);
            return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
                ? parsed
                : null;
        }
    }
}
