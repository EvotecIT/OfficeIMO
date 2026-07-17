using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

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
            IReadOnlyList<ExcelPivotInteractionCacheInfo> existingCaches = GetWorkbookSlicerCaches();
            string name = string.IsNullOrWhiteSpace(cacheName)
                ? CreateUniquePivotInteractionCacheName("Slicer", sourceField, existingCaches)
                : EnsureUniquePivotInteractionCacheName(cacheName!, existingCaches);
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
            if (!IsDateOnlyPivotSourceField(pivot, sourceField.Trim())) {
                throw new ArgumentException(
                    $"Field '{sourceField}' is not a date-only source field and cannot be used for a timeline binding.",
                    nameof(sourceField));
            }
            IReadOnlyList<ExcelPivotInteractionCacheInfo> existingCaches = GetWorkbookTimelineCaches();
            string name = string.IsNullOrWhiteSpace(cacheName)
                ? CreateUniquePivotInteractionCacheName("Timeline", sourceField, existingCaches)
                : EnsureUniquePivotInteractionCacheName(cacheName!, existingCaches);
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

        private bool IsDateOnlyPivotSourceField(ExcelPivotTableInfo pivot, string sourceField) {
            PivotTablePart? pivotPart = FindPivotTablePart(pivot);
            bool? currentSourceResult = TryIsDateOnlyPivotSourceFieldFromCurrentSource(pivot, sourceField, pivotPart);
            if (currentSourceResult.HasValue) {
                return currentSourceResult.Value;
            }

            CacheField? cacheField = pivotPart?
                .PivotTableCacheDefinitionPart?
                .PivotCacheDefinition?
                .CacheFields?
                .Elements<CacheField>()
                .FirstOrDefault(field => string.Equals(field.Name?.Value, sourceField, StringComparison.OrdinalIgnoreCase));
            SharedItems? sharedItems = cacheField?.SharedItems;
            if (sharedItems != null) {
                bool containsDate = sharedItems.ContainsDate?.Value == true
                    || sharedItems.Elements<DateTimeItem>().Any();
                bool containsNonDate = sharedItems.ContainsString?.Value == true
                    || sharedItems.ContainsNumber?.Value == true
                    || sharedItems.ChildElements.Any(item => !(item is DateTimeItem) && !(item is MissingItem));
                if (containsDate || containsNonDate) {
                    return containsDate && !containsNonDate;
                }
            }

            return false;
        }

        private PivotTablePart? FindPivotTablePart(ExcelPivotTableInfo pivot) {
            return WorkbookPartRoot.WorksheetParts
                .SelectMany(part => part.PivotTableParts)
                .FirstOrDefault(part =>
                    part.PivotTableDefinition?.CacheId?.Value == pivot.CacheId
                    && string.Equals(part.PivotTableDefinition?.Name?.Value, pivot.Name, StringComparison.OrdinalIgnoreCase));
        }

        private bool? TryIsDateOnlyPivotSourceFieldFromCurrentSource(
            ExcelPivotTableInfo pivot,
            string sourceField,
            PivotTablePart? pivotPart) {
            if (string.IsNullOrWhiteSpace(pivot.SourceSheet)
                || string.IsNullOrWhiteSpace(pivot.SourceRange)
                || !A1.TryParseRange(pivot.SourceRange!.Replace("$", string.Empty), out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                return null;
            }

            ExcelSheet? sourceSheet = Sheets.FirstOrDefault(sheet =>
                string.Equals(sheet.Name, pivot.SourceSheet, StringComparison.OrdinalIgnoreCase));
            if (sourceSheet == null) {
                return null;
            }

            List<string> normalizedHeaders = sourceSheet.BuildPivotHeaders(firstRow, firstColumn, lastColumn);
            int sourceFieldIndex = normalizedHeaders.FindIndex(header =>
                string.Equals(header, sourceField, StringComparison.OrdinalIgnoreCase));
            if (sourceFieldIndex < 0) {
                List<CacheField> databaseFields = pivotPart?
                    .PivotTableCacheDefinitionPart?
                    .PivotCacheDefinition?
                    .CacheFields?
                    .Elements<CacheField>()
                    .Where(field => field.DatabaseField?.Value != false)
                    .ToList() ?? new List<CacheField>();
                sourceFieldIndex = databaseFields.FindIndex(field =>
                    string.Equals(field.Name?.Value, sourceField, StringComparison.OrdinalIgnoreCase));
                if (sourceFieldIndex < 0 || sourceFieldIndex >= normalizedHeaders.Count) {
                    return false;
                }
            }

            int sourceColumn = firstColumn + sourceFieldIndex;

            bool foundDate = false;
            for (int row = firstRow + 1; row <= lastRow; row++) {
                if (!sourceSheet.TryGetCellValueSnapshot(row, sourceColumn, out ExcelCellValueSnapshot? value)) {
                    continue;
                }
                if (value!.Kind != ExcelCellValueKind.DateTime
                    && !sourceSheet.IsPivotDateSourceValue(row, sourceColumn)) {
                    return false;
                }

                foundDate = true;
            }

            return foundDate;
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

        private static string CreateUniquePivotInteractionCacheName(
            string prefix,
            string sourceField,
            IReadOnlyList<ExcelPivotInteractionCacheInfo> existingCaches) {
            string baseName = CreatePivotInteractionCacheName(prefix, sourceField);
            if (!existingCaches.Any(cache => string.Equals(cache.Name, baseName, StringComparison.OrdinalIgnoreCase))) {
                return baseName;
            }

            for (int suffix = 2; suffix < int.MaxValue; suffix++) {
                string suffixText = "_" + suffix.ToString(System.Globalization.CultureInfo.InvariantCulture);
                int baseLength = Math.Min(baseName.Length, 255 - suffixText.Length);
                string candidate = baseName.Substring(0, baseLength) + suffixText;
                if (!existingCaches.Any(cache => string.Equals(cache.Name, candidate, StringComparison.OrdinalIgnoreCase))) {
                    return candidate;
                }
            }

            throw new InvalidOperationException("Unable to generate a unique pivot interaction cache name.");
        }
    }
}
