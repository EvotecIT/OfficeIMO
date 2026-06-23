using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;
using System.Text;

namespace OfficeIMO.Excel.LegacyXls {
    /// <summary>
    /// Compact import summary intended for corpus baselines and preflight checks.
    /// </summary>
    public sealed class LegacyXlsImportReport {
        internal LegacyXlsImportReport(LegacyXlsWorkbook workbook) {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));

            WorksheetCount = workbook.Worksheets.Count;
            UnsupportedSheetCount = workbook.UnsupportedSheets.Count;
            CellCount = workbook.Worksheets.Sum(sheet => sheet.Cells.Count);
            FormulaCellCount = workbook.Worksheets.Sum(sheet => sheet.Cells.Count(cell => cell.IsFormula));
            CommentCount = workbook.Worksheets.Sum(sheet => sheet.Comments.Count);
            HyperlinkCount = workbook.Worksheets.Sum(sheet => sheet.Hyperlinks.Count);
            DataValidationCount = workbook.Worksheets.Sum(sheet => sheet.DataValidations.Count);
            ConditionalFormattingCount = workbook.Worksheets.Sum(sheet => sheet.ConditionalFormattings.Count);
            AutoFilterCriteriaCount = workbook.Worksheets.Sum(sheet => sheet.AutoFilterCriteria.Count);
            DefinedNameCount = workbook.DefinedNames.Count;
            ExternalReferenceCount = workbook.ExternalReferences.Count;
            ExternalCellCacheCount = workbook.ExternalReferences.Sum(reference => reference.CachedCellCaches.Count);
            ExternalCachedCellCount = workbook.ExternalReferences.Sum(reference => reference.CachedCellCaches.Sum(cache => cache.Cells.Count));
            UnsupportedFeatureCount = workbook.UnsupportedFeatures.Count;
            PreservedFeatureRecordCount = workbook.PreservedFeatureRecords.Count;
            ErrorCount = workbook.Diagnostics.Count(diagnostic => diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error);
            WarningCount = workbook.Diagnostics.Count(diagnostic => diagnostic.Severity == LegacyXlsDiagnosticSeverity.Warning);
            DiagnosticsByCode = CountByCode(workbook.Diagnostics.Select(diagnostic => diagnostic.Code));
            FormulaTokenBlockers = CountByCode(workbook.Diagnostics
                .Where(diagnostic => string.Equals(diagnostic.Code, "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED", StringComparison.OrdinalIgnoreCase))
                .Select(diagnostic => diagnostic.DetailCode ?? "FormulaUnknown"));
            UnsupportedFeaturesByCode = CountByCode(workbook.UnsupportedFeatures.Select(feature => feature.Code));
            UnsupportedFeaturesByKind = CountByKind(workbook.UnsupportedFeatures);
            UnsupportedFeaturesByRecordType = CountByCode(workbook.UnsupportedFeatures
                .Where(feature => feature.RecordType.HasValue)
                .Select(feature => $"{feature.Kind}|{feature.Code}|0x{feature.RecordType!.Value:X4}"));
            UnsupportedFeaturesByDetail = CountByCode(workbook.UnsupportedFeatures
                .Where(feature => !string.IsNullOrWhiteSpace(feature.DetailCode))
                .Select(feature => $"{feature.Kind}|{feature.Code}|{feature.DetailCode}"));
            UnsupportedFeaturesByLocation = CountByCode(workbook.UnsupportedFeatures
                .Select(GetFeatureLocationKey));
            PreservedFeatureRecordsByKind = CountPreservedRecordsByKind(workbook.PreservedFeatureRecords);
            PreservedFeatureRecordsByDetail = CountByCode(workbook.PreservedFeatureRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.DetailCode))
                .Select(record => $"{record.Kind}|{record.Code}|{record.DetailCode}"));
        }

        /// <summary>Gets the number of imported worksheet sheets.</summary>
        public int WorksheetCount { get; }

        /// <summary>Gets the number of sheet entries that were preserved as unsupported metadata.</summary>
        public int UnsupportedSheetCount { get; }

        /// <summary>Gets the number of imported cells, including blank style-only cells.</summary>
        public int CellCount { get; }

        /// <summary>Gets the number of imported formula cells.</summary>
        public int FormulaCellCount { get; }

        /// <summary>Gets the number of imported comments.</summary>
        public int CommentCount { get; }

        /// <summary>Gets the number of imported hyperlinks.</summary>
        public int HyperlinkCount { get; }

        /// <summary>Gets the number of imported data validation rules.</summary>
        public int DataValidationCount { get; }

        /// <summary>Gets the number of imported conditional formatting rules.</summary>
        public int ConditionalFormattingCount { get; }

        /// <summary>Gets the number of imported AutoFilter criteria columns.</summary>
        public int AutoFilterCriteriaCount { get; }

        /// <summary>Gets the number of imported defined names.</summary>
        public int DefinedNameCount { get; }

        /// <summary>Gets the number of preserved external-reference records.</summary>
        public int ExternalReferenceCount { get; }

        /// <summary>Gets the number of preserved external cell cache sections.</summary>
        public int ExternalCellCacheCount { get; }

        /// <summary>Gets the number of preserved cached external cell values.</summary>
        public int ExternalCachedCellCount { get; }

        /// <summary>Gets the number of unsupported or preserve-only feature findings.</summary>
        public int UnsupportedFeatureCount { get; }

        /// <summary>Gets the number of preserve-only BIFF feature records with typed metadata.</summary>
        public int PreservedFeatureRecordCount { get; }

        /// <summary>Gets the number of error diagnostics produced during import.</summary>
        public int ErrorCount { get; }

        /// <summary>Gets the number of warning diagnostics produced during import.</summary>
        public int WarningCount { get; }

        /// <summary>Gets diagnostic counts grouped by stable diagnostic code.</summary>
        public IReadOnlyDictionary<string, int> DiagnosticsByCode { get; }

        /// <summary>Gets unsupported formula token blockers grouped by stable detail key.</summary>
        public IReadOnlyDictionary<string, int> FormulaTokenBlockers { get; }

        /// <summary>Gets unsupported/preserve-only feature counts grouped by stable feature code.</summary>
        public IReadOnlyDictionary<string, int> UnsupportedFeaturesByCode { get; }

        /// <summary>Gets unsupported/preserve-only feature counts grouped by feature kind.</summary>
        public IReadOnlyDictionary<LegacyXlsUnsupportedFeatureKind, int> UnsupportedFeaturesByKind { get; }

        /// <summary>Gets unsupported/preserve-only feature counts grouped by kind, code, and BIFF record type.</summary>
        public IReadOnlyDictionary<string, int> UnsupportedFeaturesByRecordType { get; }

        /// <summary>Gets unsupported/preserve-only feature counts grouped by kind, code, and stable feature subtype.</summary>
        public IReadOnlyDictionary<string, int> UnsupportedFeaturesByDetail { get; }

        /// <summary>Gets unsupported/preserve-only feature counts grouped by code and workbook or sheet location.</summary>
        public IReadOnlyDictionary<string, int> UnsupportedFeaturesByLocation { get; }

        /// <summary>Gets preserved feature record counts grouped by feature kind.</summary>
        public IReadOnlyDictionary<LegacyXlsUnsupportedFeatureKind, int> PreservedFeatureRecordsByKind { get; }

        /// <summary>Gets preserved feature record counts grouped by kind, code, and stable feature subtype.</summary>
        public IReadOnlyDictionary<string, int> PreservedFeatureRecordsByDetail { get; }

        /// <summary>Gets whether the import produced error diagnostics.</summary>
        public bool HasImportErrors => ErrorCount > 0;

        /// <summary>Gets whether the import discovered unsupported or preserve-only features.</summary>
        public bool HasUnsupportedFeatures => UnsupportedFeatureCount > 0;

        /// <summary>
        /// Returns a compact Markdown summary suitable for corpus snapshots.
        /// </summary>
        public string ToMarkdown() {
            var builder = new StringBuilder();
            builder.AppendLine("# Legacy XLS Import Report");
            builder.AppendLine();
            builder.AppendLine($"Worksheets: {WorksheetCount}");
            builder.AppendLine($"Unsupported sheets: {UnsupportedSheetCount}");
            builder.AppendLine($"Cells: {CellCount}");
            builder.AppendLine($"Formula cells: {FormulaCellCount}");
            builder.AppendLine($"Comments: {CommentCount}");
            builder.AppendLine($"Hyperlinks: {HyperlinkCount}");
            builder.AppendLine($"Data validations: {DataValidationCount}");
            builder.AppendLine($"Conditional formatting rules: {ConditionalFormattingCount}");
            builder.AppendLine($"AutoFilter criteria columns: {AutoFilterCriteriaCount}");
            builder.AppendLine($"Defined names: {DefinedNameCount}");
            builder.AppendLine($"External references: {ExternalReferenceCount}");
            builder.AppendLine($"External cell caches: {ExternalCellCacheCount}");
            builder.AppendLine($"External cached cells: {ExternalCachedCellCount}");
            builder.AppendLine($"Unsupported features: {UnsupportedFeatureCount}");
            builder.AppendLine($"Preserved feature records: {PreservedFeatureRecordCount}");
            builder.AppendLine($"Errors: {ErrorCount}");
            builder.AppendLine($"Warnings: {WarningCount}");
            AppendDictionary(builder, "Diagnostics By Code", DiagnosticsByCode);
            AppendDictionary(builder, "Formula Token Blockers", FormulaTokenBlockers);
            AppendDictionary(builder, "Unsupported Features By Code", UnsupportedFeaturesByCode);
            AppendDictionary(builder, "Unsupported Features By Kind", UnsupportedFeaturesByKind.ToDictionary(
                entry => entry.Key.ToString(),
                entry => entry.Value,
                StringComparer.OrdinalIgnoreCase));
            AppendDictionary(builder, "Unsupported Feature Record Types", UnsupportedFeaturesByRecordType);
            AppendDictionary(builder, "Unsupported Feature Details", UnsupportedFeaturesByDetail);
            AppendDictionary(builder, "Unsupported Feature Locations", UnsupportedFeaturesByLocation);
            AppendDictionary(builder, "Preserved Feature Records By Kind", PreservedFeatureRecordsByKind.ToDictionary(
                entry => entry.Key.ToString(),
                entry => entry.Value,
                StringComparer.OrdinalIgnoreCase));
            AppendDictionary(builder, "Preserved Feature Record Details", PreservedFeatureRecordsByDetail);
            return builder.ToString();
        }

        private static IReadOnlyDictionary<string, int> CountByCode(IEnumerable<string> codes) {
            return codes
                .Where(code => !string.IsNullOrWhiteSpace(code))
                .GroupBy(code => code, StringComparer.OrdinalIgnoreCase)
                .OrderBy(group => group.Key, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.Count(), StringComparer.OrdinalIgnoreCase);
        }

        private static IReadOnlyDictionary<LegacyXlsUnsupportedFeatureKind, int> CountByKind(IEnumerable<LegacyXlsUnsupportedFeature> features) {
            return features
                .GroupBy(feature => feature.Kind)
                .OrderBy(group => group.Key.ToString(), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.Count());
        }

        private static IReadOnlyDictionary<LegacyXlsUnsupportedFeatureKind, int> CountPreservedRecordsByKind(IEnumerable<LegacyXlsPreservedFeatureRecord> records) {
            return records
                .GroupBy(record => record.Kind)
                .OrderBy(group => group.Key.ToString(), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.Count());
        }

        private static string GetFeatureLocationKey(LegacyXlsUnsupportedFeature feature) {
            string location = string.IsNullOrWhiteSpace(feature.SheetName) ? "(workbook)" : feature.SheetName!;
            return feature.Code + "|" + location;
        }

        private static void AppendDictionary(StringBuilder builder, string title, IReadOnlyDictionary<string, int> values) {
            if (values.Count == 0) {
                return;
            }

            builder.AppendLine();
            builder.AppendLine("## " + title);
            builder.AppendLine();
            builder.AppendLine("| Key | Count |");
            builder.AppendLine("| --- | --- |");
            foreach (KeyValuePair<string, int> entry in values) {
                builder.Append("| ");
                builder.Append(EscapeMarkdownCell(entry.Key));
                builder.Append(" | ");
                builder.Append(entry.Value);
                builder.AppendLine(" |");
            }
        }

        private static string EscapeMarkdownCell(string value) {
            return value.Replace("\\", "\\\\").Replace("|", "\\|").Replace("\r", " ").Replace("\n", " ");
        }
    }
}
