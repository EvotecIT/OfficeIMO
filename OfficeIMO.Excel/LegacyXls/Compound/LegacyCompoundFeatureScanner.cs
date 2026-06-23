using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Compound {
    internal static class LegacyCompoundFeatureScanner {
        private const string VbaProjectCode = "XLS-COMPOUND-FEATURE-VBA-PROJECT-PRESERVED";
        private const string OleObjectCode = "XLS-COMPOUND-FEATURE-OLE-OBJECT-PRESERVED";

        internal static void AddPreserveOnlyFeatures(
            LegacyCompoundFile compoundFile,
            LegacyXlsWorkbook workbook,
            LegacyXlsImportOptions options) {
            if (TryGetVbaProjectEntries(compoundFile, out IReadOnlyList<string> entries, out IReadOnlyDictionary<string, LegacyXlsCompoundFeatureEntryRole> entryRoles, out string description)) {
                workbook.MutableCompoundFeatureRecords.Add(new LegacyXlsCompoundFeatureRecord(
                    LegacyXlsCompoundFeatureRecordKind.VbaProject,
                    entries,
                    entryRoles));
                AddFeature(workbook, options, new LegacyXlsUnsupportedFeature(
                    LegacyXlsUnsupportedFeatureKind.VbaProject,
                    VbaProjectCode,
                    description,
                    detailCode: "Compound:VbaProjectStorage"));
            }

            if (TryGetOleObjectEntries(compoundFile, out entries, out entryRoles, out description)) {
                workbook.MutableCompoundFeatureRecords.Add(new LegacyXlsCompoundFeatureRecord(
                    LegacyXlsCompoundFeatureRecordKind.OleObject,
                    entries,
                    entryRoles));
                AddFeature(workbook, options, new LegacyXlsUnsupportedFeature(
                    LegacyXlsUnsupportedFeatureKind.OleObject,
                    OleObjectCode,
                    description,
                    detailCode: "Compound:OleObjectStorage"));
            }
        }

        private static void AddFeature(
            LegacyXlsWorkbook workbook,
            LegacyXlsImportOptions options,
            LegacyXlsUnsupportedFeature feature) {
            workbook.MutableUnsupportedFeatures.Add(feature);
            if (options.ReportUnsupportedRecords) {
                workbook.MutableDiagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Info,
                    feature.Code,
                    feature.Description,
                    detailCode: feature.DetailCode));
            }
        }

        private static bool TryGetVbaProjectEntries(
            LegacyCompoundFile compoundFile,
            out IReadOnlyList<string> entries,
            out IReadOnlyDictionary<string, LegacyXlsCompoundFeatureEntryRole> entryRoles,
            out string description) {
            Dictionary<string, LegacyXlsCompoundFeatureEntryRole> matchingEntries = compoundFile.Entries
                .Where(IsVbaProjectEntry)
                .Select(entry => new KeyValuePair<string, LegacyXlsCompoundFeatureEntryRole>(
                    string.IsNullOrWhiteSpace(entry.Path) ? entry.Name : entry.Path,
                    ClassifyVbaProjectEntry(entry)))
                .GroupBy(entry => entry.Key, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First().Value, StringComparer.OrdinalIgnoreCase);
            List<string> orderedEntries = matchingEntries.Keys
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(entry => entry, StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (orderedEntries.Count == 0) {
                entries = Array.Empty<string>();
                entryRoles = new Dictionary<string, LegacyXlsCompoundFeatureEntryRole>(StringComparer.OrdinalIgnoreCase);
                description = string.Empty;
                return false;
            }

            description = "The compound XLS container contains VBA project storage. Macro projects are preserve-only; OfficeIMO.Excel does not import, edit, or execute VBA code. Entries: "
                + string.Join("; ", orderedEntries.Take(8))
                + (orderedEntries.Count > 8 ? $"; +{orderedEntries.Count - 8} more" : string.Empty);
            entries = orderedEntries;
            entryRoles = matchingEntries;
            return true;
        }

        private static bool TryGetOleObjectEntries(
            LegacyCompoundFile compoundFile,
            out IReadOnlyList<string> entries,
            out IReadOnlyDictionary<string, LegacyXlsCompoundFeatureEntryRole> entryRoles,
            out string description) {
            Dictionary<string, LegacyXlsCompoundFeatureEntryRole> matchingEntries = compoundFile.Entries
                .Where(IsOleObjectEntry)
                .Select(entry => new KeyValuePair<string, LegacyXlsCompoundFeatureEntryRole>(
                    string.IsNullOrWhiteSpace(entry.Path) ? entry.Name : entry.Path,
                    ClassifyOleObjectEntry(entry)))
                .GroupBy(entry => entry.Key, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First().Value, StringComparer.OrdinalIgnoreCase);
            List<string> orderedEntries = matchingEntries.Keys
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(entry => entry, StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (orderedEntries.Count == 0) {
                entries = Array.Empty<string>();
                entryRoles = new Dictionary<string, LegacyXlsCompoundFeatureEntryRole>(StringComparer.OrdinalIgnoreCase);
                description = string.Empty;
                return false;
            }

            description = "The compound XLS container contains embedded OLE object storage. Embedded objects are preserve-only; OfficeIMO.Excel does not import, edit, or execute embedded OLE content. Entries: "
                + string.Join("; ", orderedEntries.Take(8))
                + (orderedEntries.Count > 8 ? $"; +{orderedEntries.Count - 8} more" : string.Empty);
            entries = orderedEntries;
            entryRoles = matchingEntries;
            return true;
        }

        private static bool IsVbaProjectEntry(LegacyCompoundFileEntry entry) {
            if (entry.Name.Equals("_VBA_PROJECT_CUR", StringComparison.OrdinalIgnoreCase)
                || entry.Name.Equals("_VBA_PROJECT", StringComparison.OrdinalIgnoreCase)) {
                return true;
            }

            return entry.Path.IndexOf("/_VBA_PROJECT_CUR/", StringComparison.OrdinalIgnoreCase) >= 0
                || entry.Path.EndsWith("/_VBA_PROJECT_CUR", StringComparison.OrdinalIgnoreCase)
                || entry.Path.IndexOf("/VBA/", StringComparison.OrdinalIgnoreCase) >= 0
                || entry.Path.EndsWith("/VBA", StringComparison.OrdinalIgnoreCase);
        }

        private static LegacyXlsCompoundFeatureEntryRole ClassifyVbaProjectEntry(LegacyCompoundFileEntry entry) {
            if (entry.Name.Equals("_VBA_PROJECT_CUR", StringComparison.OrdinalIgnoreCase)) {
                return LegacyXlsCompoundFeatureEntryRole.VbaProjectStorage;
            }

            if (entry.Path.EndsWith("/VBA", StringComparison.OrdinalIgnoreCase)
                || entry.Name.Equals("VBA", StringComparison.OrdinalIgnoreCase)) {
                return LegacyXlsCompoundFeatureEntryRole.VbaStorage;
            }

            if (entry.Name.Equals("dir", StringComparison.OrdinalIgnoreCase)
                && entry.Path.IndexOf("/VBA/", StringComparison.OrdinalIgnoreCase) >= 0) {
                return LegacyXlsCompoundFeatureEntryRole.VbaDirStream;
            }

            if (entry.Name.Equals("_VBA_PROJECT", StringComparison.OrdinalIgnoreCase)) {
                return LegacyXlsCompoundFeatureEntryRole.VbaProjectStream;
            }

            if (entry.Path.IndexOf("/VBA/", StringComparison.OrdinalIgnoreCase) >= 0) {
                return LegacyXlsCompoundFeatureEntryRole.VbaModuleStream;
            }

            return LegacyXlsCompoundFeatureEntryRole.Unknown;
        }

        private static bool IsOleObjectEntry(LegacyCompoundFileEntry entry) {
            return entry.Name.Equals("ObjectPool", StringComparison.OrdinalIgnoreCase)
                || entry.Name.Equals("Ole", StringComparison.OrdinalIgnoreCase)
                || entry.Name.Equals("\u0001Ole", StringComparison.OrdinalIgnoreCase)
                || entry.Name.Equals("Ole10Native", StringComparison.OrdinalIgnoreCase)
                || entry.Path.IndexOf("/ObjectPool/", StringComparison.OrdinalIgnoreCase) >= 0
                || entry.Path.EndsWith("/ObjectPool", StringComparison.OrdinalIgnoreCase)
                || entry.Path.IndexOf("/Ole10Native", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static LegacyXlsCompoundFeatureEntryRole ClassifyOleObjectEntry(LegacyCompoundFileEntry entry) {
            if (entry.Name.Equals("ObjectPool", StringComparison.OrdinalIgnoreCase)
                || entry.Path.EndsWith("/ObjectPool", StringComparison.OrdinalIgnoreCase)) {
                return LegacyXlsCompoundFeatureEntryRole.OleObjectPoolStorage;
            }

            if (entry.Name.Equals("Ole10Native", StringComparison.OrdinalIgnoreCase)
                || entry.Path.IndexOf("/Ole10Native", StringComparison.OrdinalIgnoreCase) >= 0) {
                return LegacyXlsCompoundFeatureEntryRole.OleNativeStream;
            }

            if (entry.Name.Equals("Ole", StringComparison.OrdinalIgnoreCase)
                || entry.Name.Equals("\u0001Ole", StringComparison.OrdinalIgnoreCase)) {
                return LegacyXlsCompoundFeatureEntryRole.OleStream;
            }

            if (entry.Path.IndexOf("/ObjectPool/", StringComparison.OrdinalIgnoreCase) >= 0) {
                return LegacyXlsCompoundFeatureEntryRole.OleObjectStorage;
            }

            return LegacyXlsCompoundFeatureEntryRole.Unknown;
        }
    }
}
