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
            if (HasVbaProjectStorage(compoundFile, out string description)) {
                AddFeature(workbook, options, new LegacyXlsUnsupportedFeature(
                    LegacyXlsUnsupportedFeatureKind.VbaProject,
                    VbaProjectCode,
                    description));
            }

            if (HasOleObjectStorage(compoundFile, out description)) {
                AddFeature(workbook, options, new LegacyXlsUnsupportedFeature(
                    LegacyXlsUnsupportedFeatureKind.OleObject,
                    OleObjectCode,
                    description));
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
                    feature.Description));
            }
        }

        private static bool HasVbaProjectStorage(LegacyCompoundFile compoundFile, out string description) {
            List<string> entries = compoundFile.Entries
                .Where(IsVbaProjectEntry)
                .Select(entry => string.IsNullOrWhiteSpace(entry.Path) ? entry.Name : entry.Path)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(entry => entry, StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (entries.Count == 0) {
                description = string.Empty;
                return false;
            }

            description = "The compound XLS container contains VBA project storage. Macro projects are preserve-only; OfficeIMO.Excel does not import, edit, or execute VBA code. Entries: "
                + string.Join("; ", entries.Take(8))
                + (entries.Count > 8 ? $"; +{entries.Count - 8} more" : string.Empty);
            return true;
        }

        private static bool HasOleObjectStorage(LegacyCompoundFile compoundFile, out string description) {
            List<string> entries = compoundFile.Entries
                .Where(IsOleObjectEntry)
                .Select(entry => string.IsNullOrWhiteSpace(entry.Path) ? entry.Name : entry.Path)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(entry => entry, StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (entries.Count == 0) {
                description = string.Empty;
                return false;
            }

            description = "The compound XLS container contains embedded OLE object storage. Embedded objects are preserve-only; OfficeIMO.Excel does not import, edit, or execute embedded OLE content. Entries: "
                + string.Join("; ", entries.Take(8))
                + (entries.Count > 8 ? $"; +{entries.Count - 8} more" : string.Empty);
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

        private static bool IsOleObjectEntry(LegacyCompoundFileEntry entry) {
            return entry.Name.Equals("ObjectPool", StringComparison.OrdinalIgnoreCase)
                || entry.Name.Equals("Ole", StringComparison.OrdinalIgnoreCase)
                || entry.Name.Equals("\u0001Ole", StringComparison.OrdinalIgnoreCase)
                || entry.Name.Equals("Ole10Native", StringComparison.OrdinalIgnoreCase)
                || entry.Path.IndexOf("/ObjectPool/", StringComparison.OrdinalIgnoreCase) >= 0
                || entry.Path.EndsWith("/ObjectPool", StringComparison.OrdinalIgnoreCase)
                || entry.Path.IndexOf("/Ole10Native", StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }
}
