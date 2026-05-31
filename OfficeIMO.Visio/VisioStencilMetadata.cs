using System;
using System.Collections.Generic;
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
            return cells.AsReadOnly();
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

        private static IReadOnlyList<string> Coalesce(IReadOnlyList<string> candidate, IReadOnlyList<string> fallback) {
            return candidate.Count > 0 ? candidate : fallback;
        }
    }
}
