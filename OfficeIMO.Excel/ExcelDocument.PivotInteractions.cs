using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Adds slicer cache metadata bound to a validated pivot table field.
        /// </summary>
        /// <remarks>
        /// This authors and validates cache metadata. Excel-compatible software is still required to materialize the slicer UI shape.
        /// </remarks>
        public ExtendedPart AddPivotSlicerCache(string pivotTableName, string sourceField, string? cacheName = null) {
            ExcelPivotTableInfo pivot = ValidatePivotInteractionBinding(pivotTableName, sourceField);
            string name = EnsureUniquePivotInteractionCacheName(
                string.IsNullOrWhiteSpace(cacheName) ? CreatePivotInteractionCacheName("Slicer", sourceField) : cacheName!,
                GetWorkbookSlicerCaches());
            return AddWorkbookSlicerCache(new ExcelSlicerCacheOptions {
                Name = name,
                SourceName = sourceField.Trim(),
                PivotTableName = pivot.Name
            });
        }

        /// <summary>
        /// Adds timeline cache metadata bound to a validated pivot table field.
        /// </summary>
        /// <remarks>
        /// This authors and validates cache metadata. Excel-compatible software is still required to materialize the timeline UI shape.
        /// </remarks>
        public ExtendedPart AddPivotTimelineCache(string pivotTableName, string sourceField, string? cacheName = null) {
            ExcelPivotTableInfo pivot = ValidatePivotInteractionBinding(pivotTableName, sourceField);
            string name = EnsureUniquePivotInteractionCacheName(
                string.IsNullOrWhiteSpace(cacheName) ? CreatePivotInteractionCacheName("Timeline", sourceField) : cacheName!,
                GetWorkbookTimelineCaches());
            return AddWorkbookTimelineCache(new ExcelTimelineCacheOptions {
                Name = name,
                SourceName = sourceField.Trim(),
                PivotTableName = pivot.Name
            });
        }

        /// <summary>Returns workbook-level slicer cache metadata parts.</summary>
        public IReadOnlyList<ExcelPivotInteractionCacheInfo> GetWorkbookSlicerCaches() {
            return GetWorkbookPivotInteractionCaches(ExcelPivotInteractionCacheKind.Slicer);
        }

        /// <summary>Returns workbook-level timeline cache metadata parts.</summary>
        public IReadOnlyList<ExcelPivotInteractionCacheInfo> GetWorkbookTimelineCaches() {
            return GetWorkbookPivotInteractionCaches(ExcelPivotInteractionCacheKind.Timeline);
        }

        private ExcelPivotTableInfo ValidatePivotInteractionBinding(string pivotTableName, string sourceField) {
            if (string.IsNullOrWhiteSpace(pivotTableName)) {
                throw new ArgumentNullException(nameof(pivotTableName));
            }
            if (string.IsNullOrWhiteSpace(sourceField)) {
                throw new ArgumentNullException(nameof(sourceField));
            }

            List<ExcelPivotTableInfo> matches = GetPivotTables()
                .Where(pivot => string.Equals(pivot.Name, pivotTableName.Trim(), StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (matches.Count == 0) {
                throw new ArgumentException($"Pivot table '{pivotTableName}' was not found.", nameof(pivotTableName));
            }
            if (matches.Count > 1) {
                throw new InvalidOperationException($"Pivot table name '{pivotTableName}' is ambiguous in this workbook.");
            }

            ExcelPivotTableInfo pivot = matches[0];
            string field = sourceField.Trim();
            bool found = pivot.Fields.Any(item => string.Equals(item.FieldName, field, StringComparison.OrdinalIgnoreCase))
                || pivot.RowFields.Any(item => string.Equals(item, field, StringComparison.OrdinalIgnoreCase))
                || pivot.ColumnFields.Any(item => string.Equals(item, field, StringComparison.OrdinalIgnoreCase))
                || pivot.PageFields.Any(item => string.Equals(item, field, StringComparison.OrdinalIgnoreCase))
                || pivot.DataFields.Any(item => string.Equals(item.FieldName, field, StringComparison.OrdinalIgnoreCase));
            if (!found) {
                throw new ArgumentException($"Field '{sourceField}' was not found in pivot table '{pivot.Name}'.", nameof(sourceField));
            }

            return pivot;
        }

        private IReadOnlyList<ExcelPivotInteractionCacheInfo> GetWorkbookPivotInteractionCaches(ExcelPivotInteractionCacheKind kind) {
            string contentType = kind == ExcelPivotInteractionCacheKind.Slicer
                ? WorkbookSlicerCacheContentType
                : WorkbookTimelineCacheContentType;
            string relationshipType = kind == ExcelPivotInteractionCacheKind.Slicer
                ? WorkbookSlicerCacheRelationshipType
                : WorkbookTimelineCacheRelationshipType;
            var caches = new List<ExcelPivotInteractionCacheInfo>();
            foreach (IdPartPair pair in WorkbookPartRoot.Parts) {
                OpenXmlPart part = pair.OpenXmlPart;
                if (!string.Equals(part.ContentType, contentType, StringComparison.OrdinalIgnoreCase)
                    || !string.Equals(part.RelationshipType, relationshipType, StringComparison.Ordinal)) {
                    continue;
                }

                try {
                    XDocument xml = XDocument.Parse(ReadPivotInteractionPartText(part));
                    XElement? root = xml.Root;
                    if (root == null) {
                        continue;
                    }

                    string name = (string?)root.Attribute("name") ?? string.Empty;
                    string? sourceName = (string?)root.Attribute("sourceName");
                    string? pivotTableName = root.Attributes()
                            .FirstOrDefault(attribute => attribute.Name.LocalName == "pivotTableName")?.Value
                        ?? root.Descendants().FirstOrDefault(element => element.Name.LocalName == "pivotTable")?.Attribute("name")?.Value;
                    caches.Add(new ExcelPivotInteractionCacheInfo(kind, name, sourceName, pivotTableName, pair.RelationshipId));
                } catch (System.Xml.XmlException) {
                    caches.Add(new ExcelPivotInteractionCacheInfo(kind, string.Empty, null, null, pair.RelationshipId));
                }
            }

            return caches
                .OrderBy(cache => cache.Name, StringComparer.OrdinalIgnoreCase)
                .ThenBy(cache => cache.RelationshipId, StringComparer.Ordinal)
                .ToList();
        }

        private static string ReadPivotInteractionPartText(OpenXmlPart part) {
            using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
            using var reader = new StreamReader(stream, Encoding.UTF8);
            return reader.ReadToEnd();
        }

        private static string EnsureUniquePivotInteractionCacheName(
            string candidate,
            IReadOnlyList<ExcelPivotInteractionCacheInfo> existingCaches) {
            string normalized = candidate.Trim();
            if (normalized.Length == 0) {
                throw new ArgumentNullException(nameof(candidate));
            }
            if (normalized.Length > 255) {
                throw new ArgumentException("Pivot interaction cache names cannot exceed 255 characters.", nameof(candidate));
            }
            if (existingCaches.Any(cache => string.Equals(cache.Name, normalized, StringComparison.OrdinalIgnoreCase))) {
                throw new InvalidOperationException($"Pivot interaction cache '{normalized}' already exists.");
            }

            return normalized;
        }

        private static string CreatePivotInteractionCacheName(string prefix, string sourceField) {
            var builder = new StringBuilder(prefix.Length + sourceField.Length + 1);
            builder.Append(prefix);
            builder.Append('_');
            foreach (char character in sourceField.Trim()) {
                builder.Append(char.IsLetterOrDigit(character) || character == '_' ? character : '_');
            }

            return builder.Length > 255 ? builder.ToString(0, 255) : builder.ToString();
        }
    }
}
