using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using OfficeIMO.Visio.Stencils;

namespace OfficeIMO.Visio {
    internal static class VisioStencilMetadata {
        private const string ListSeparator = ";";

        internal static void Apply(VisioShape shape, VisioStencilShape stencil, string? catalogName) {
            Set(shape, VisioSemanticUserCells.StencilId, stencil.Id);
            Set(shape, VisioSemanticUserCells.StencilName, stencil.Name);
            Set(shape, VisioSemanticUserCells.StencilCategory, stencil.Category);
            Set(shape, VisioSemanticUserCells.StencilCatalog, catalogName);
            Set(shape, VisioSemanticUserCells.StencilSourcePackagePath, NormalizePath(stencil.SourcePackagePath));
            Set(shape, VisioSemanticUserCells.StencilKeywords, Join(stencil.Keywords));
            Set(shape, VisioSemanticUserCells.StencilAliases, Join(stencil.Aliases));
            Set(shape, VisioSemanticUserCells.StencilTags, Join(stencil.Tags));
            Set(shape, VisioSemanticUserCells.StencilIconNameU, stencil.IconNameU);
            Set(shape, VisioSemanticUserCells.StencilDefaultWidth, FormatDouble(stencil.DefaultWidth));
            Set(shape, VisioSemanticUserCells.StencilDefaultHeight, FormatDouble(stencil.DefaultHeight));
            Set(shape, VisioSemanticUserCells.StencilDefaultUnit, stencil.DefaultUnit?.ToString());
            Set(shape, VisioSemanticUserCells.StencilPreviewImageRelationshipId, stencil.PreviewImage?.RelationshipId);
            Set(shape, VisioSemanticUserCells.StencilPreviewImageTarget, stencil.PreviewImage?.Target);
            Set(shape, VisioSemanticUserCells.StencilPreviewImageContentType, stencil.PreviewImage?.ContentType);
            Set(shape, VisioSemanticUserCells.StencilPreviewImageExtension, stencil.PreviewImage?.Extension);
            Set(shape, VisioSemanticUserCells.StencilPreviewImageByteLength, FormatLong(stencil.PreviewImage?.ByteLength));
            ApplyConnectionPoints(shape, stencil);
        }

        internal static void Apply(VisioMaster master, VisioStencilShape stencil, string? catalogName) {
            master.StencilId = stencil.Id;
            master.StencilName = stencil.Name;
            master.StencilCategory = stencil.Category;
            master.StencilCatalogName = string.IsNullOrWhiteSpace(catalogName) ? master.StencilCatalogName : catalogName;
            master.StencilSourcePackagePath = NormalizePath(stencil.SourcePackagePath) ?? master.StencilSourcePackagePath;
            master.StencilKeywords = Normalize(stencil.Keywords);
            master.StencilAliases = Normalize(stencil.Aliases);
            master.StencilTags = Normalize(stencil.Tags);
            master.StencilIconNameU = stencil.IconNameU;
            master.StencilDefaultWidth = stencil.DefaultWidth;
            master.StencilDefaultHeight = stencil.DefaultHeight;
            master.StencilDefaultUnit = stencil.DefaultUnit;
            master.StencilPreviewImageRelationshipId = stencil.PreviewImage?.RelationshipId;
            master.StencilPreviewImageTarget = stencil.PreviewImage?.Target;
            master.StencilPreviewImageContentType = stencil.PreviewImage?.ContentType;
            master.StencilPreviewImageExtension = stencil.PreviewImage?.Extension;
            master.StencilPreviewImageByteLength = stencil.PreviewImage?.ByteLength;
        }

        internal static void Apply(VisioMaster master, IEnumerable<VisioUserCell> userCells) {
            Dictionary<string, string?> values = userCells
                .GroupBy(cell => cell.Name, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First().Value, StringComparer.OrdinalIgnoreCase);

            master.StencilId = Get(values, VisioSemanticUserCells.StencilId) ?? master.StencilId;
            master.StencilName = Get(values, VisioSemanticUserCells.StencilName) ?? master.StencilName;
            master.StencilCategory = Get(values, VisioSemanticUserCells.StencilCategory) ?? master.StencilCategory;
            master.StencilCatalogName = Get(values, VisioSemanticUserCells.StencilCatalog) ?? master.StencilCatalogName;
            master.StencilSourcePackagePath = Get(values, VisioSemanticUserCells.StencilSourcePackagePath) ?? master.StencilSourcePackagePath;
            master.StencilKeywords = Coalesce(Split(Get(values, VisioSemanticUserCells.StencilKeywords)), master.StencilKeywords);
            master.StencilAliases = Coalesce(Split(Get(values, VisioSemanticUserCells.StencilAliases)), master.StencilAliases);
            master.StencilTags = Coalesce(Split(Get(values, VisioSemanticUserCells.StencilTags)), master.StencilTags);
            master.StencilIconNameU = Get(values, VisioSemanticUserCells.StencilIconNameU) ?? master.StencilIconNameU;
            master.StencilDefaultWidth = GetDouble(values, VisioSemanticUserCells.StencilDefaultWidth) ?? master.StencilDefaultWidth;
            master.StencilDefaultHeight = GetDouble(values, VisioSemanticUserCells.StencilDefaultHeight) ?? master.StencilDefaultHeight;
            master.StencilDefaultUnit = GetUnit(values, VisioSemanticUserCells.StencilDefaultUnit) ?? master.StencilDefaultUnit;
            master.StencilPreviewImageRelationshipId = Get(values, VisioSemanticUserCells.StencilPreviewImageRelationshipId) ?? master.StencilPreviewImageRelationshipId;
            master.StencilPreviewImageTarget = Get(values, VisioSemanticUserCells.StencilPreviewImageTarget) ?? master.StencilPreviewImageTarget;
            master.StencilPreviewImageContentType = Get(values, VisioSemanticUserCells.StencilPreviewImageContentType) ?? master.StencilPreviewImageContentType;
            master.StencilPreviewImageExtension = Get(values, VisioSemanticUserCells.StencilPreviewImageExtension) ?? master.StencilPreviewImageExtension;
            master.StencilPreviewImageByteLength = GetLong(values, VisioSemanticUserCells.StencilPreviewImageByteLength) ?? master.StencilPreviewImageByteLength;
        }

        internal static IReadOnlyList<VisioUserCell> CreateMasterUserCells(VisioMaster master) {
            List<VisioUserCell> cells = new();
            Add(cells, VisioSemanticUserCells.StencilId, master.StencilId);
            Add(cells, VisioSemanticUserCells.StencilName, master.StencilName);
            Add(cells, VisioSemanticUserCells.StencilCategory, master.StencilCategory);
            Add(cells, VisioSemanticUserCells.StencilCatalog, master.StencilCatalogName);
            Add(cells, VisioSemanticUserCells.StencilSourcePackagePath, master.StencilSourcePackagePath);
            Add(cells, VisioSemanticUserCells.StencilKeywords, Join(master.StencilKeywords));
            Add(cells, VisioSemanticUserCells.StencilAliases, Join(master.StencilAliases));
            Add(cells, VisioSemanticUserCells.StencilTags, Join(master.StencilTags));
            Add(cells, VisioSemanticUserCells.StencilIconNameU, master.StencilIconNameU);
            Add(cells, VisioSemanticUserCells.StencilDefaultWidth, FormatDouble(master.StencilDefaultWidth));
            Add(cells, VisioSemanticUserCells.StencilDefaultHeight, FormatDouble(master.StencilDefaultHeight));
            Add(cells, VisioSemanticUserCells.StencilDefaultUnit, master.StencilDefaultUnit?.ToString());
            Add(cells, VisioSemanticUserCells.StencilPreviewImageRelationshipId, master.StencilPreviewImageRelationshipId);
            Add(cells, VisioSemanticUserCells.StencilPreviewImageTarget, master.StencilPreviewImageTarget);
            Add(cells, VisioSemanticUserCells.StencilPreviewImageContentType, master.StencilPreviewImageContentType);
            Add(cells, VisioSemanticUserCells.StencilPreviewImageExtension, master.StencilPreviewImageExtension);
            Add(cells, VisioSemanticUserCells.StencilPreviewImageByteLength, FormatLong(master.StencilPreviewImageByteLength));
            return cells.AsReadOnly();
        }

        internal static bool HasStencilMetadata(VisioShape shape) {
            if (shape == null) {
                return false;
            }

            return shape.UserCells.Any(cell =>
                string.Equals(cell.Name, VisioSemanticUserCells.StencilId, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(cell.Name, VisioSemanticUserCells.StencilName, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(cell.Name, VisioSemanticUserCells.StencilCatalog, StringComparison.OrdinalIgnoreCase));
        }

        internal static void Clear(VisioShape shape) {
            if (shape == null) {
                return;
            }

            for (int i = shape.UserCells.Count - 1; i >= 0; i--) {
                string name = shape.UserCells[i].Name;
                if (string.Equals(name, VisioSemanticUserCells.StencilId, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, VisioSemanticUserCells.StencilName, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, VisioSemanticUserCells.StencilCategory, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, VisioSemanticUserCells.StencilCatalog, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, VisioSemanticUserCells.StencilSourcePackagePath, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, VisioSemanticUserCells.StencilKeywords, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, VisioSemanticUserCells.StencilAliases, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, VisioSemanticUserCells.StencilTags, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, VisioSemanticUserCells.StencilIconNameU, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, VisioSemanticUserCells.StencilDefaultWidth, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, VisioSemanticUserCells.StencilDefaultHeight, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, VisioSemanticUserCells.StencilDefaultUnit, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, VisioSemanticUserCells.StencilPreviewImageRelationshipId, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, VisioSemanticUserCells.StencilPreviewImageTarget, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, VisioSemanticUserCells.StencilPreviewImageContentType, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, VisioSemanticUserCells.StencilPreviewImageExtension, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, VisioSemanticUserCells.StencilPreviewImageByteLength, StringComparison.OrdinalIgnoreCase)) {
                    shape.UserCells.RemoveAt(i);
                }
            }
        }

        internal static string? GetUserCellValue(IEnumerable<VisioInspectionUserCellSnapshot> userCells, string name) {
            return userCells
                .FirstOrDefault(cell => string.Equals(cell.Name, name, StringComparison.OrdinalIgnoreCase))
                ?.Value;
        }

        internal static IReadOnlyList<string> GetUserCellList(IEnumerable<VisioInspectionUserCellSnapshot> userCells, string name) {
            return Split(GetUserCellValue(userCells, name));
        }

        internal static IReadOnlyList<string> Normalize(IEnumerable<string>? values) {
            return (values ?? Enumerable.Empty<string>())
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Select(value => value.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
                .ToList()
                .AsReadOnly();
        }

        internal static string Join(IEnumerable<string>? values) {
            return string.Join(ListSeparator, Normalize(values));
        }

        internal static IReadOnlyList<string> Split(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return Array.Empty<string>();
            }

            return Normalize(value!.Split(new[] { ListSeparator }, StringSplitOptions.RemoveEmptyEntries));
        }

        internal static string? NormalizePath(string? path) {
            if (string.IsNullOrWhiteSpace(path)) {
                return null;
            }

            try {
                return Path.GetFullPath(path!);
            } catch (Exception) {
                return path!.Trim();
            }
        }

        private static void Set(VisioShape shape, string name, string? value) {
            if (!string.IsNullOrWhiteSpace(value)) {
                shape.SetUserCell(name, value, "STR", prompt: "OfficeIMO stencil metadata");
            }
        }

        private static void Add(ICollection<VisioUserCell> cells, string name, string? value) {
            if (!string.IsNullOrWhiteSpace(value)) {
                cells.Add(new VisioUserCell(name, value) {
                    Unit = "STR",
                    Prompt = "OfficeIMO stencil metadata"
                });
            }
        }

        private static string? Get(IReadOnlyDictionary<string, string?> values, string key) {
            return values.TryGetValue(key, out string? value) && !string.IsNullOrWhiteSpace(value)
                ? value
                : null;
        }

        private static double? GetDouble(IReadOnlyDictionary<string, string?> values, string key) {
            string? value = Get(values, key);
            return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
                ? parsed
                : null;
        }

        private static VisioMeasurementUnit? GetUnit(IReadOnlyDictionary<string, string?> values, string key) {
            string? value = Get(values, key);
            return Enum.TryParse(value, ignoreCase: true, out VisioMeasurementUnit unit)
                ? unit
                : null;
        }

        private static long? GetLong(IReadOnlyDictionary<string, string?> values, string key) {
            string? value = Get(values, key);
            return long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long parsed)
                ? parsed
                : null;
        }

        private static string? FormatDouble(double? value) {
            return value.HasValue
                ? value.Value.ToString("0.######", CultureInfo.InvariantCulture)
                : null;
        }

        private static string? FormatLong(long? value) {
            return value.HasValue
                ? value.Value.ToString(CultureInfo.InvariantCulture)
                : null;
        }

        private static void ApplyConnectionPoints(VisioShape shape, VisioStencilShape stencil) {
            if (shape.ConnectionPoints.Count > 0 ||
                stencil.SourceConnectionPoints.Count == 0) {
                return;
            }

            double baseWidth = GetDefaultSizeInInches(stencil.DefaultWidth, stencil.DefaultUnit);
            double baseHeight = GetDefaultSizeInInches(stencil.DefaultHeight, stencil.DefaultUnit);
            double scaleX = baseWidth > 0 ? shape.Width / baseWidth : 1D;
            double scaleY = baseHeight > 0 ? shape.Height / baseHeight : 1D;

            foreach (VisioStencilConnectionPoint point in stencil.SourceConnectionPoints) {
                shape.ConnectionPoints.Add(new VisioConnectionPoint(point.X * scaleX, point.Y * scaleY, point.DirX, point.DirY) {
                    SectionIndex = point.SectionIndex
                });
            }
        }

        private static double GetDefaultSizeInInches(double value, VisioMeasurementUnit? unit) {
            return unit.HasValue
                ? value.ToInches(unit.Value)
                : value;
        }

        private static IReadOnlyList<string> Coalesce(IReadOnlyList<string> candidate, IReadOnlyList<string> fallback) {
            return candidate.Count > 0 ? candidate : fallback;
        }
    }
}
