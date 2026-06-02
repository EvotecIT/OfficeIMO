using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Aggregated stencil family profile grouped by catalog/category metadata.
    /// </summary>
    public sealed class VisioStencilFamilyProfile {
        private VisioStencilFamilyProfile(
            string key,
            string? stencilCatalogName,
            string? stencilCategory,
            IReadOnlyList<string> stencilSourcePackagePaths,
            IReadOnlyList<string> stencilKeywords,
            IReadOnlyList<string> stencilAliases,
            IReadOnlyList<string> stencilTags,
            IReadOnlyList<string> stencilIconNameUs,
            IReadOnlyList<string> stencilDefaultUnits,
            IReadOnlyList<string> stencilPreviewImageContentTypes,
            IReadOnlyList<string> stencilPreviewImageExtensions,
            IReadOnlyList<string> stencilIds,
            IReadOnlyList<string> usageKeys,
            int shapeCount,
            int stencilBackedShapeCount,
            int masterBackedShapeCount,
            int packageBackedShapeCount,
            int generatedMasterBackedShapeCount,
            int basicGeometryShapeCount,
            int connectionPointCount,
            int connectionPointShapeCount,
            double placedWidthMinimum,
            double placedWidthMaximum,
            double placedHeightMinimum,
            double placedHeightMaximum,
            double? sourceDefaultWidthMinimum,
            double? sourceDefaultWidthMaximum,
            double? sourceDefaultHeightMinimum,
            double? sourceDefaultHeightMaximum) {
            Key = key;
            StencilCatalogName = stencilCatalogName;
            StencilCategory = stencilCategory;
            StencilSourcePackagePaths = stencilSourcePackagePaths;
            StencilKeywords = stencilKeywords;
            StencilAliases = stencilAliases;
            StencilTags = stencilTags;
            StencilIconNameUs = stencilIconNameUs;
            StencilDefaultUnits = stencilDefaultUnits;
            StencilPreviewImageContentTypes = stencilPreviewImageContentTypes;
            StencilPreviewImageExtensions = stencilPreviewImageExtensions;
            StencilIds = stencilIds;
            UsageKeys = usageKeys;
            ShapeCount = shapeCount;
            StencilBackedShapeCount = stencilBackedShapeCount;
            MasterBackedShapeCount = masterBackedShapeCount;
            PackageBackedShapeCount = packageBackedShapeCount;
            GeneratedMasterBackedShapeCount = generatedMasterBackedShapeCount;
            BasicGeometryShapeCount = basicGeometryShapeCount;
            ConnectionPointCount = connectionPointCount;
            ConnectionPointShapeCount = connectionPointShapeCount;
            PlacedWidthMinimum = placedWidthMinimum;
            PlacedWidthMaximum = placedWidthMaximum;
            PlacedHeightMinimum = placedHeightMinimum;
            PlacedHeightMaximum = placedHeightMaximum;
            SourceDefaultWidthMinimum = sourceDefaultWidthMinimum;
            SourceDefaultWidthMaximum = sourceDefaultWidthMaximum;
            SourceDefaultHeightMinimum = sourceDefaultHeightMinimum;
            SourceDefaultHeightMaximum = sourceDefaultHeightMaximum;
        }

        /// <summary>Stable family key used in profile snapshots.</summary>
        public string Key { get; }

        /// <summary>Catalog name represented by this family, when available.</summary>
        public string? StencilCatalogName { get; }

        /// <summary>Category represented by this family, when available.</summary>
        public string? StencilCategory { get; }

        /// <summary>Distinct source package paths represented by this family.</summary>
        public IReadOnlyList<string> StencilSourcePackagePaths { get; }

        /// <summary>Distinct stencil keywords represented by this family.</summary>
        public IReadOnlyList<string> StencilKeywords { get; }

        /// <summary>Distinct stencil aliases represented by this family.</summary>
        public IReadOnlyList<string> StencilAliases { get; }

        /// <summary>Distinct stencil tags represented by this family.</summary>
        public IReadOnlyList<string> StencilTags { get; }

        /// <summary>Distinct stencil icon master universal names represented by this family.</summary>
        public IReadOnlyList<string> StencilIconNameUs { get; }

        /// <summary>Distinct source default-size units represented by this family.</summary>
        public IReadOnlyList<string> StencilDefaultUnits { get; }

        /// <summary>Distinct preview image content types represented by this family.</summary>
        public IReadOnlyList<string> StencilPreviewImageContentTypes { get; }

        /// <summary>Distinct preview image extensions represented by this family.</summary>
        public IReadOnlyList<string> StencilPreviewImageExtensions { get; }

        /// <summary>Distinct stencil identifiers represented by this family.</summary>
        public IReadOnlyList<string> StencilIds { get; }

        /// <summary>Usage keys included in this family.</summary>
        public IReadOnlyList<string> UsageKeys { get; }

        /// <summary>Total shapes represented by this family.</summary>
        public int ShapeCount { get; }

        /// <summary>Shapes in this family that carry OfficeIMO stencil identity metadata.</summary>
        public int StencilBackedShapeCount { get; }

        /// <summary>Shapes in this family backed by any registered master.</summary>
        public int MasterBackedShapeCount { get; }

        /// <summary>Shapes in this family backed by imported stencil-package masters.</summary>
        public int PackageBackedShapeCount { get; }

        /// <summary>Shapes in this family backed by generated OfficeIMO masters.</summary>
        public int GeneratedMasterBackedShapeCount { get; }

        /// <summary>Shapes in this family represented by direct geometry.</summary>
        public int BasicGeometryShapeCount { get; }

        /// <summary>Total connection points exposed by shapes in this family.</summary>
        public int ConnectionPointCount { get; }

        /// <summary>Number of shapes in this family that expose at least one connection point.</summary>
        public int ConnectionPointShapeCount { get; }

        /// <summary>Minimum placed width for shapes in this family.</summary>
        public double PlacedWidthMinimum { get; }

        /// <summary>Maximum placed width for shapes in this family.</summary>
        public double PlacedWidthMaximum { get; }

        /// <summary>Minimum placed height for shapes in this family.</summary>
        public double PlacedHeightMinimum { get; }

        /// <summary>Maximum placed height for shapes in this family.</summary>
        public double PlacedHeightMaximum { get; }

        /// <summary>Minimum source default width represented by this family, when known.</summary>
        public double? SourceDefaultWidthMinimum { get; }

        /// <summary>Maximum source default width represented by this family, when known.</summary>
        public double? SourceDefaultWidthMaximum { get; }

        /// <summary>Minimum source default height represented by this family, when known.</summary>
        public double? SourceDefaultHeightMinimum { get; }

        /// <summary>Maximum source default height represented by this family, when known.</summary>
        public double? SourceDefaultHeightMaximum { get; }

        internal static List<VisioStencilFamilyProfile> FromUsages(IEnumerable<VisioStencilUsageProfile> usages) {
            return usages
                .Where(IsStencilFamilyUsage)
                .GroupBy(CreateFamilyKey, StringComparer.OrdinalIgnoreCase)
                .Select(group => FromUsageGroup(group.Key, group))
                .OrderBy(family => family.Key, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        internal void AppendText(StringBuilder builder) {
            string prefix = "family[" + VisioInspectionSnapshot.EscapeKey(Key) + "]";
            VisioStencilProfile.AppendLine(builder, prefix + ".stencilCatalog", StencilCatalogName);
            VisioStencilProfile.AppendLine(builder, prefix + ".stencilCategory", StencilCategory);
            VisioStencilProfile.AppendLine(builder, prefix + ".stencilSourcePackagePaths", string.Join(",", StencilSourcePackagePaths));
            VisioStencilProfile.AppendLine(builder, prefix + ".stencilKeywords", string.Join(",", StencilKeywords));
            VisioStencilProfile.AppendLine(builder, prefix + ".stencilAliases", string.Join(",", StencilAliases));
            VisioStencilProfile.AppendLine(builder, prefix + ".stencilTags", string.Join(",", StencilTags));
            VisioStencilProfile.AppendLine(builder, prefix + ".stencilIconNameUs", string.Join(",", StencilIconNameUs));
            VisioStencilProfile.AppendLine(builder, prefix + ".stencilDefaultUnits", string.Join(",", StencilDefaultUnits));
            VisioStencilProfile.AppendLine(builder, prefix + ".stencilPreviewImageContentTypes", string.Join(",", StencilPreviewImageContentTypes));
            VisioStencilProfile.AppendLine(builder, prefix + ".stencilPreviewImageExtensions", string.Join(",", StencilPreviewImageExtensions));
            VisioStencilProfile.AppendLine(builder, prefix + ".stencilIds", string.Join(",", StencilIds));
            VisioStencilProfile.AppendLine(builder, prefix + ".usageKeys", string.Join(",", UsageKeys));
            VisioStencilProfile.AppendLine(builder, prefix + ".shapeCount", ShapeCount);
            VisioStencilProfile.AppendLine(builder, prefix + ".stencilBackedShapeCount", StencilBackedShapeCount);
            VisioStencilProfile.AppendLine(builder, prefix + ".masterBackedShapeCount", MasterBackedShapeCount);
            VisioStencilProfile.AppendLine(builder, prefix + ".packageBackedShapeCount", PackageBackedShapeCount);
            VisioStencilProfile.AppendLine(builder, prefix + ".generatedMasterBackedShapeCount", GeneratedMasterBackedShapeCount);
            VisioStencilProfile.AppendLine(builder, prefix + ".basicGeometryShapeCount", BasicGeometryShapeCount);
            VisioStencilProfile.AppendLine(builder, prefix + ".connectionPointCount", ConnectionPointCount);
            VisioStencilProfile.AppendLine(builder, prefix + ".connectionPointShapeCount", ConnectionPointShapeCount);
            VisioStencilProfile.AppendLine(builder, prefix + ".placedWidthMinimum", PlacedWidthMinimum);
            VisioStencilProfile.AppendLine(builder, prefix + ".placedWidthMaximum", PlacedWidthMaximum);
            VisioStencilProfile.AppendLine(builder, prefix + ".placedHeightMinimum", PlacedHeightMinimum);
            VisioStencilProfile.AppendLine(builder, prefix + ".placedHeightMaximum", PlacedHeightMaximum);
            VisioStencilProfile.AppendLine(builder, prefix + ".sourceDefaultWidthMinimum", SourceDefaultWidthMinimum);
            VisioStencilProfile.AppendLine(builder, prefix + ".sourceDefaultWidthMaximum", SourceDefaultWidthMaximum);
            VisioStencilProfile.AppendLine(builder, prefix + ".sourceDefaultHeightMinimum", SourceDefaultHeightMinimum);
            VisioStencilProfile.AppendLine(builder, prefix + ".sourceDefaultHeightMaximum", SourceDefaultHeightMaximum);
        }

        private static VisioStencilFamilyProfile FromUsageGroup(string key, IEnumerable<VisioStencilUsageProfile> usages) {
            List<VisioStencilUsageProfile> usageList = usages.ToList();
            return new VisioStencilFamilyProfile(
                key,
                FirstDistinctValue(usageList, usage => usage.StencilCatalogName),
                FirstDistinctValue(usageList, usage => usage.StencilCategory),
                CollectDistinctValues(usageList, usage => usage.StencilSourcePackagePath),
                CollectDistinctListValues(usageList, usage => usage.StencilKeywords),
                CollectDistinctListValues(usageList, usage => usage.StencilAliases),
                CollectDistinctListValues(usageList, usage => usage.StencilTags),
                CollectDistinctValues(usageList, usage => usage.StencilIconNameU),
                CollectDistinctValues(usageList, usage => usage.StencilDefaultUnit),
                CollectDistinctValues(usageList, usage => usage.StencilPreviewImageContentType),
                CollectDistinctValues(usageList, usage => usage.StencilPreviewImageExtension),
                CollectDistinctValues(usageList, usage => usage.StencilId),
                usageList.Select(usage => usage.Key).OrderBy(value => value, StringComparer.OrdinalIgnoreCase).ToList().AsReadOnly(),
                usageList.Sum(usage => usage.Count),
                usageList.Where(usage => !string.IsNullOrWhiteSpace(usage.StencilId)).Sum(usage => usage.Count),
                usageList.Where(IsMasterBacked).Sum(usage => usage.Count),
                usageList.Where(usage => usage.Kind == VisioStencilProfileUsageKind.PackageBackedMaster).Sum(usage => usage.Count),
                usageList.Where(usage => usage.Kind == VisioStencilProfileUsageKind.GeneratedMaster).Sum(usage => usage.Count),
                usageList.Where(usage => usage.Kind == VisioStencilProfileUsageKind.BasicGeometry).Sum(usage => usage.Count),
                usageList.Sum(usage => usage.ConnectionPointCount),
                usageList.Sum(usage => usage.ConnectionPointShapeCount),
                usageList.Min(usage => usage.PlacedWidthMinimum),
                usageList.Max(usage => usage.PlacedWidthMaximum),
                usageList.Min(usage => usage.PlacedHeightMinimum),
                usageList.Max(usage => usage.PlacedHeightMaximum),
                MinNullable(usageList, usage => usage.SourceDefaultWidth),
                MaxNullable(usageList, usage => usage.SourceDefaultWidth),
                MinNullable(usageList, usage => usage.SourceDefaultHeight),
                MaxNullable(usageList, usage => usage.SourceDefaultHeight));
        }

        private static bool IsStencilFamilyUsage(VisioStencilUsageProfile usage) {
            return !string.IsNullOrWhiteSpace(usage.StencilId) ||
                   !string.IsNullOrWhiteSpace(usage.StencilCatalogName) ||
                   !string.IsNullOrWhiteSpace(usage.StencilCategory) ||
                   !string.IsNullOrWhiteSpace(usage.StencilSourcePackagePath);
        }

        private static string CreateFamilyKey(VisioStencilUsageProfile usage) {
            if (!string.IsNullOrWhiteSpace(usage.StencilCatalogName) && !string.IsNullOrWhiteSpace(usage.StencilCategory)) {
                return "stencil-family:" + usage.StencilCatalogName + "/" + usage.StencilCategory;
            }

            if (!string.IsNullOrWhiteSpace(usage.StencilCategory)) {
                return "stencil-family:" + usage.StencilCategory;
            }

            if (!string.IsNullOrWhiteSpace(usage.StencilCatalogName)) {
                return "stencil-family:" + usage.StencilCatalogName;
            }

            if (!string.IsNullOrWhiteSpace(usage.StencilSourcePackagePath)) {
                return "stencil-family:" + usage.StencilSourcePackagePath;
            }

            return "stencil-family:" + usage.StencilId;
        }

        private static bool IsMasterBacked(VisioStencilUsageProfile usage) {
            return usage.Kind == VisioStencilProfileUsageKind.PackageBackedMaster ||
                   usage.Kind == VisioStencilProfileUsageKind.GeneratedMaster;
        }

        private static string? FirstDistinctValue(IEnumerable<VisioStencilUsageProfile> usages, Func<VisioStencilUsageProfile, string?> selector) {
            List<string> values = CollectDistinctValues(usages, selector).ToList();
            return values.Count == 1 ? values[0] : null;
        }

        private static IReadOnlyList<string> CollectDistinctValues(IEnumerable<VisioStencilUsageProfile> usages, Func<VisioStencilUsageProfile, string?> selector) {
            return usages
                .Select(selector)
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Select(value => value!)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
                .ToList()
                .AsReadOnly();
        }

        private static IReadOnlyList<string> CollectDistinctListValues(IEnumerable<VisioStencilUsageProfile> usages, Func<VisioStencilUsageProfile, IReadOnlyList<string>> selector) {
            return usages
                .SelectMany(selector)
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
                .ToList()
                .AsReadOnly();
        }

        private static double? MinNullable(IEnumerable<VisioStencilUsageProfile> usages, Func<VisioStencilUsageProfile, double?> selector) {
            List<double> values = usages
                .Select(selector)
                .Where(value => value.HasValue)
                .Select(value => value!.Value)
                .ToList();
            return values.Count == 0 ? null : values.Min();
        }

        private static double? MaxNullable(IEnumerable<VisioStencilUsageProfile> usages, Func<VisioStencilUsageProfile, double?> selector) {
            List<double> values = usages
                .Select(selector)
                .Where(value => value.HasValue)
                .Select(value => value!.Value)
                .ToList();
            return values.Count == 0 ? null : values.Max();
        }
    }
}
