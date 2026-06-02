using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeIMO.Visio {
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
            string? stencilIconNameU,
            double? sourceDefaultWidth,
            double? sourceDefaultHeight,
            string? stencilDefaultUnit,
            string? stencilPreviewImageContentType,
            string? stencilPreviewImageExtension,
            int count,
            int connectionPointCount,
            int connectionPointShapeCount,
            IReadOnlyList<string> shapeIds,
            IReadOnlyList<string> pageNames,
            IReadOnlyList<string> shapeDataKeys,
            double placedWidthMinimum,
            double placedWidthMaximum,
            double placedHeightMinimum,
            double placedHeightMaximum) {
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
            StencilIconNameU = stencilIconNameU;
            SourceDefaultWidth = sourceDefaultWidth;
            SourceDefaultHeight = sourceDefaultHeight;
            StencilDefaultUnit = stencilDefaultUnit;
            StencilPreviewImageContentType = stencilPreviewImageContentType;
            StencilPreviewImageExtension = stencilPreviewImageExtension;
            Count = count;
            ConnectionPointCount = connectionPointCount;
            ConnectionPointShapeCount = connectionPointShapeCount;
            ShapeIds = shapeIds;
            PageNames = pageNames;
            ShapeDataKeys = shapeDataKeys;
            PlacedWidthMinimum = placedWidthMinimum;
            PlacedWidthMaximum = placedWidthMaximum;
            PlacedHeightMinimum = placedHeightMinimum;
            PlacedHeightMaximum = placedHeightMaximum;
            SnapshotKey = key;
        }

        /// <summary>Stable usage key.</summary>
        public string Key { get; }

        internal string SnapshotKey { get; set; }

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

        /// <summary>Stencil preview icon master universal name represented by this usage group.</summary>
        public string? StencilIconNameU { get; }

        /// <summary>Source stencil default width before placement scaling, when known.</summary>
        public double? SourceDefaultWidth { get; }

        /// <summary>Source stencil default height before placement scaling, when known.</summary>
        public double? SourceDefaultHeight { get; }

        /// <summary>Source stencil default-size unit, when known.</summary>
        public string? StencilDefaultUnit { get; }

        /// <summary>Preview image content type represented by this usage group, when known.</summary>
        public string? StencilPreviewImageContentType { get; }

        /// <summary>Preview image extension represented by this usage group, when known.</summary>
        public string? StencilPreviewImageExtension { get; }

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

        /// <summary>Minimum placed width for shapes in this usage group.</summary>
        public double PlacedWidthMinimum { get; }

        /// <summary>Maximum placed width for shapes in this usage group.</summary>
        public double PlacedWidthMaximum { get; }

        /// <summary>Minimum placed height for shapes in this usage group.</summary>
        public double PlacedHeightMinimum { get; }

        /// <summary>Maximum placed height for shapes in this usage group.</summary>
        public double PlacedHeightMaximum { get; }

        internal static VisioStencilUsageProfile FromShapes(
            VisioStencilUsageKey key,
            IEnumerable<VisioInspectionShapeSnapshot> shapes,
            IReadOnlyList<VisioInspectionPageSnapshot> pages) {
            List<VisioInspectionShapeSnapshot> shapeList = shapes.ToList();
            Dictionary<VisioInspectionShapeSnapshot, string> pageByShape = pages
                .SelectMany(page => page.Shapes.Select(shape => new { Shape = shape, Page = page.Name }))
                .ToDictionary(item => item.Shape, item => item.Page);

            IReadOnlyList<string> shapeIds = shapeList
                .Select(shape => shape.Id)
                .OrderBy(id => id, StringComparer.OrdinalIgnoreCase)
                .ToList()
                .AsReadOnly();
            IReadOnlyList<string> pageNames = shapeList
                .Select(shape => pageByShape.TryGetValue(shape, out string? pageName) ? pageName : string.Empty)
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
                key.StencilIconNameU,
                key.SourceDefaultWidth,
                key.SourceDefaultHeight,
                key.StencilDefaultUnit,
                key.StencilPreviewImageContentType,
                key.StencilPreviewImageExtension,
                shapeList.Count,
                shapeList.Sum(shape => shape.ConnectionPointCount),
                shapeList.Count(shape => shape.ConnectionPointCount > 0),
                shapeIds,
                pageNames,
                shapeDataKeys,
                shapeList.Min(shape => shape.Width),
                shapeList.Max(shape => shape.Width),
                shapeList.Min(shape => shape.Height),
                shapeList.Max(shape => shape.Height));
        }

        internal void AppendText(StringBuilder builder) {
            string prefix = "usage[" + VisioInspectionSnapshot.EscapeKey(SnapshotKey) + "]";
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
            VisioStencilProfile.AppendLine(builder, prefix + ".stencilIconNameU", StencilIconNameU);
            VisioStencilProfile.AppendLine(builder, prefix + ".sourceDefaultWidth", SourceDefaultWidth);
            VisioStencilProfile.AppendLine(builder, prefix + ".sourceDefaultHeight", SourceDefaultHeight);
            VisioStencilProfile.AppendLine(builder, prefix + ".stencilDefaultUnit", StencilDefaultUnit);
            VisioStencilProfile.AppendLine(builder, prefix + ".stencilPreviewImageContentType", StencilPreviewImageContentType);
            VisioStencilProfile.AppendLine(builder, prefix + ".stencilPreviewImageExtension", StencilPreviewImageExtension);
            VisioStencilProfile.AppendLine(builder, prefix + ".count", Count);
            VisioStencilProfile.AppendLine(builder, prefix + ".connectionPointCount", ConnectionPointCount);
            VisioStencilProfile.AppendLine(builder, prefix + ".connectionPointShapeCount", ConnectionPointShapeCount);
            VisioStencilProfile.AppendLine(builder, prefix + ".shapeIds", string.Join(",", ShapeIds));
            VisioStencilProfile.AppendLine(builder, prefix + ".pages", string.Join(",", PageNames));
            VisioStencilProfile.AppendLine(builder, prefix + ".shapeDataKeys", string.Join(",", ShapeDataKeys));
            VisioStencilProfile.AppendLine(builder, prefix + ".placedWidthMinimum", PlacedWidthMinimum);
            VisioStencilProfile.AppendLine(builder, prefix + ".placedWidthMaximum", PlacedWidthMaximum);
            VisioStencilProfile.AppendLine(builder, prefix + ".placedHeightMinimum", PlacedHeightMinimum);
            VisioStencilProfile.AppendLine(builder, prefix + ".placedHeightMaximum", PlacedHeightMaximum);
        }
    }
}
