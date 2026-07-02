using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;
using OfficeIMO.Shared;

namespace OfficeIMO.Excel.LegacyXls.Compound {
    internal static class LegacyCompoundFeatureScanner {
        private const string VbaProjectCode = "XLS-COMPOUND-FEATURE-VBA-PROJECT-PRESERVED";
        private const string OleObjectCode = "XLS-COMPOUND-FEATURE-OLE-OBJECT-PRESERVED";
        private const string DigitalSignatureCode = "XLS-COMPOUND-FEATURE-DIGITAL-SIGNATURE-DIAGNOSED";

        internal static void AddPreserveOnlyFeatures(
            OfficeCompoundFile compoundFile,
            LegacyXlsWorkbook workbook,
            LegacyXlsImportOptions options) {
            if (TryGetVbaProjectEntries(
                compoundFile,
                out IReadOnlyList<string> entries,
                out IReadOnlyDictionary<string, LegacyXlsCompoundFeatureEntryRole> entryRoles,
                out IReadOnlyDictionary<string, long> entrySizes,
                out IReadOnlyDictionary<string, LegacyXlsCompoundFeatureEntryObjectType> entryObjectTypes,
                out IReadOnlyDictionary<string, LegacyXlsCompoundFeatureEntryContentKind> entryContentKinds,
                out string description)) {
                workbook.MutableCompoundFeatureRecords.Add(new LegacyXlsCompoundFeatureRecord(
                    LegacyXlsCompoundFeatureRecordKind.VbaProject,
                    entries,
                    entryRoles,
                    entrySizes,
                    entryObjectTypes,
                    entryContentKinds));
                AddFeature(workbook, options, new LegacyXlsUnsupportedFeature(
                    LegacyXlsUnsupportedFeatureKind.VbaProject,
                    VbaProjectCode,
                    description,
                    detailCode: "Compound:VbaProjectStorage"));
            }

            if (TryGetOleObjectEntries(
                compoundFile,
                out entries,
                out entryRoles,
                out entrySizes,
                out entryObjectTypes,
                out entryContentKinds,
                out description)) {
                workbook.MutableCompoundFeatureRecords.Add(new LegacyXlsCompoundFeatureRecord(
                    LegacyXlsCompoundFeatureRecordKind.OleObject,
                    entries,
                    entryRoles,
                    entrySizes,
                    entryObjectTypes,
                    entryContentKinds));
                AddFeature(workbook, options, new LegacyXlsUnsupportedFeature(
                    LegacyXlsUnsupportedFeatureKind.OleObject,
                    OleObjectCode,
                    description,
                    detailCode: "Compound:OleObjectStorage"));
            }

            if (TryGetDigitalSignatureEntries(
                compoundFile,
                out entries,
                out entryRoles,
                out entrySizes,
                out entryObjectTypes,
                out entryContentKinds,
                out description)) {
                workbook.MutableCompoundFeatureRecords.Add(new LegacyXlsCompoundFeatureRecord(
                    LegacyXlsCompoundFeatureRecordKind.DigitalSignature,
                    entries,
                    entryRoles,
                    entrySizes,
                    entryObjectTypes,
                    entryContentKinds));
                AddFeature(workbook, options, new LegacyXlsUnsupportedFeature(
                    LegacyXlsUnsupportedFeatureKind.DigitalSignature,
                    DigitalSignatureCode,
                    description,
                    detailCode: "Compound:DigitalSignature"));
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
            OfficeCompoundFile compoundFile,
            out IReadOnlyList<string> entries,
            out IReadOnlyDictionary<string, LegacyXlsCompoundFeatureEntryRole> entryRoles,
            out IReadOnlyDictionary<string, long> entrySizes,
            out IReadOnlyDictionary<string, LegacyXlsCompoundFeatureEntryObjectType> entryObjectTypes,
            out IReadOnlyDictionary<string, LegacyXlsCompoundFeatureEntryContentKind> entryContentKinds,
            out string description) {
            OfficeCompoundFileEntry[] matchingCompoundEntries = compoundFile.Entries
                .Where(IsVbaProjectEntry)
                .ToArray();
            Dictionary<string, LegacyXlsCompoundFeatureEntryRole> matchingEntries = matchingCompoundEntries
                .Select(entry => new KeyValuePair<string, LegacyXlsCompoundFeatureEntryRole>(
                    GetEntryKey(entry),
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
                entrySizes = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
                entryObjectTypes = new Dictionary<string, LegacyXlsCompoundFeatureEntryObjectType>(StringComparer.OrdinalIgnoreCase);
                entryContentKinds = new Dictionary<string, LegacyXlsCompoundFeatureEntryContentKind>(StringComparer.OrdinalIgnoreCase);
                description = string.Empty;
                return false;
            }

            description = "The compound XLS container contains VBA project storage. Macro projects are preserve-only; OfficeIMO.Excel does not import, edit, or execute VBA code. Entries: "
                + string.Join("; ", orderedEntries.Take(8))
                + (orderedEntries.Count > 8 ? $"; +{orderedEntries.Count - 8} more" : string.Empty);
            entries = orderedEntries;
            entryRoles = matchingEntries;
            entrySizes = BuildEntrySizes(matchingCompoundEntries);
            entryObjectTypes = BuildEntryObjectTypes(matchingCompoundEntries);
            entryContentKinds = BuildEntryContentKinds(compoundFile, matchingCompoundEntries, matchingEntries);
            return true;
        }

        private static bool TryGetOleObjectEntries(
            OfficeCompoundFile compoundFile,
            out IReadOnlyList<string> entries,
            out IReadOnlyDictionary<string, LegacyXlsCompoundFeatureEntryRole> entryRoles,
            out IReadOnlyDictionary<string, long> entrySizes,
            out IReadOnlyDictionary<string, LegacyXlsCompoundFeatureEntryObjectType> entryObjectTypes,
            out IReadOnlyDictionary<string, LegacyXlsCompoundFeatureEntryContentKind> entryContentKinds,
            out string description) {
            OfficeCompoundFileEntry[] matchingCompoundEntries = compoundFile.Entries
                .Where(IsOleObjectEntry)
                .ToArray();
            Dictionary<string, LegacyXlsCompoundFeatureEntryRole> matchingEntries = matchingCompoundEntries
                .Select(entry => new KeyValuePair<string, LegacyXlsCompoundFeatureEntryRole>(
                    GetEntryKey(entry),
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
                entrySizes = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
                entryObjectTypes = new Dictionary<string, LegacyXlsCompoundFeatureEntryObjectType>(StringComparer.OrdinalIgnoreCase);
                entryContentKinds = new Dictionary<string, LegacyXlsCompoundFeatureEntryContentKind>(StringComparer.OrdinalIgnoreCase);
                description = string.Empty;
                return false;
            }

            description = "The compound XLS container contains embedded OLE object storage. Embedded objects are preserve-only; OfficeIMO.Excel does not import, edit, or execute embedded OLE content. Entries: "
                + string.Join("; ", orderedEntries.Take(8))
                + (orderedEntries.Count > 8 ? $"; +{orderedEntries.Count - 8} more" : string.Empty);
            entries = orderedEntries;
            entryRoles = matchingEntries;
            entrySizes = BuildEntrySizes(matchingCompoundEntries);
            entryObjectTypes = BuildEntryObjectTypes(matchingCompoundEntries);
            entryContentKinds = BuildEntryContentKinds(compoundFile, matchingCompoundEntries, matchingEntries);
            return true;
        }

        private static bool TryGetDigitalSignatureEntries(
            OfficeCompoundFile compoundFile,
            out IReadOnlyList<string> entries,
            out IReadOnlyDictionary<string, LegacyXlsCompoundFeatureEntryRole> entryRoles,
            out IReadOnlyDictionary<string, long> entrySizes,
            out IReadOnlyDictionary<string, LegacyXlsCompoundFeatureEntryObjectType> entryObjectTypes,
            out IReadOnlyDictionary<string, LegacyXlsCompoundFeatureEntryContentKind> entryContentKinds,
            out string description) {
            OfficeCompoundFileEntry[] matchingCompoundEntries = compoundFile.Entries
                .Where(IsDigitalSignatureEntry)
                .ToArray();
            Dictionary<string, LegacyXlsCompoundFeatureEntryRole> matchingEntries = matchingCompoundEntries
                .Select(entry => new KeyValuePair<string, LegacyXlsCompoundFeatureEntryRole>(
                    GetEntryKey(entry),
                    ClassifyDigitalSignatureEntry(entry)))
                .GroupBy(entry => entry.Key, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First().Value, StringComparer.OrdinalIgnoreCase);
            List<string> orderedEntries = matchingEntries.Keys
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(entry => entry, StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (orderedEntries.Count == 0) {
                entries = Array.Empty<string>();
                entryRoles = new Dictionary<string, LegacyXlsCompoundFeatureEntryRole>(StringComparer.OrdinalIgnoreCase);
                entrySizes = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
                entryObjectTypes = new Dictionary<string, LegacyXlsCompoundFeatureEntryObjectType>(StringComparer.OrdinalIgnoreCase);
                entryContentKinds = new Dictionary<string, LegacyXlsCompoundFeatureEntryContentKind>(StringComparer.OrdinalIgnoreCase);
                description = string.Empty;
                return false;
            }

            description = "The compound XLS container contains digital signature storage or streams. Legacy XLS signatures are diagnosed before conversion; OfficeIMO.Excel does not validate or preserve XLS digital signatures in converted .xlsx output. Entries: "
                + string.Join("; ", orderedEntries.Take(8))
                + (orderedEntries.Count > 8 ? $"; +{orderedEntries.Count - 8} more" : string.Empty);
            entries = orderedEntries;
            entryRoles = matchingEntries;
            entrySizes = BuildEntrySizes(matchingCompoundEntries);
            entryObjectTypes = BuildEntryObjectTypes(matchingCompoundEntries);
            entryContentKinds = BuildEntryContentKinds(compoundFile, matchingCompoundEntries, matchingEntries);
            return true;
        }

        private static IReadOnlyDictionary<string, long> BuildEntrySizes(IEnumerable<OfficeCompoundFileEntry> entries) {
            return entries
                .GroupBy(GetEntryKey, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First().Size, StringComparer.OrdinalIgnoreCase);
        }

        private static IReadOnlyDictionary<string, LegacyXlsCompoundFeatureEntryObjectType> BuildEntryObjectTypes(IEnumerable<OfficeCompoundFileEntry> entries) {
            return entries
                .GroupBy(GetEntryKey, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => ToModelObjectType(group.First().ObjectType), StringComparer.OrdinalIgnoreCase);
        }

        private static IReadOnlyDictionary<string, LegacyXlsCompoundFeatureEntryContentKind> BuildEntryContentKinds(
            OfficeCompoundFile compoundFile,
            IEnumerable<OfficeCompoundFileEntry> entries,
            IReadOnlyDictionary<string, LegacyXlsCompoundFeatureEntryRole> entryRoles) {
            return entries
                .GroupBy(GetEntryKey, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(
                    group => group.Key,
                    group => {
                        OfficeCompoundFileEntry entry = group.First();
                        LegacyXlsCompoundFeatureEntryRole role = entryRoles.TryGetValue(group.Key, out LegacyXlsCompoundFeatureEntryRole value)
                            ? value
                            : LegacyXlsCompoundFeatureEntryRole.Unknown;
                        return ClassifyEntryContentKind(compoundFile, entry, role);
                    },
                    StringComparer.OrdinalIgnoreCase);
        }

        private static string GetEntryKey(OfficeCompoundFileEntry entry) {
            return string.IsNullOrWhiteSpace(entry.Path) ? entry.Name : entry.Path;
        }

        private static LegacyXlsCompoundFeatureEntryObjectType ToModelObjectType(byte objectType) {
            return objectType switch {
                1 => LegacyXlsCompoundFeatureEntryObjectType.Storage,
                2 => LegacyXlsCompoundFeatureEntryObjectType.Stream,
                5 => LegacyXlsCompoundFeatureEntryObjectType.RootStorage,
                _ => LegacyXlsCompoundFeatureEntryObjectType.Unknown
            };
        }

        private static LegacyXlsCompoundFeatureEntryContentKind ClassifyEntryContentKind(
            OfficeCompoundFile compoundFile,
            OfficeCompoundFileEntry entry,
            LegacyXlsCompoundFeatureEntryRole role) {
            if (entry.IsStorage) {
                return LegacyXlsCompoundFeatureEntryContentKind.Storage;
            }

            if (role == LegacyXlsCompoundFeatureEntryRole.VbaProjectStream) {
                return LegacyXlsCompoundFeatureEntryContentKind.VbaProjectMetadataStream;
            }

            if (role == LegacyXlsCompoundFeatureEntryRole.OleNativeStream
                || role == LegacyXlsCompoundFeatureEntryRole.OleStream) {
                return LegacyXlsCompoundFeatureEntryContentKind.OlePayloadStream;
            }

            if (role == LegacyXlsCompoundFeatureEntryRole.DigitalSignatureStream
                || role == LegacyXlsCompoundFeatureEntryRole.XmlDigitalSignatureStream) {
                return LegacyXlsCompoundFeatureEntryContentKind.DigitalSignatureStream;
            }

            if (!entry.IsStream || !TryGetStreamBytes(compoundFile, entry, out byte[] bytes)) {
                return LegacyXlsCompoundFeatureEntryContentKind.Unknown;
            }

            if (bytes.Length == 0) {
                return LegacyXlsCompoundFeatureEntryContentKind.EmptyStream;
            }

            if (role == LegacyXlsCompoundFeatureEntryRole.VbaDirStream
                || role == LegacyXlsCompoundFeatureEntryRole.VbaModuleStream) {
                return bytes[0] == 0x01
                    ? LegacyXlsCompoundFeatureEntryContentKind.VbaCompressedContainer
                    : LegacyXlsCompoundFeatureEntryContentKind.BinaryStream;
            }

            return LegacyXlsCompoundFeatureEntryContentKind.BinaryStream;
        }

        private static bool TryGetStreamBytes(OfficeCompoundFile compoundFile, OfficeCompoundFileEntry entry, out byte[] bytes) {
            if (compoundFile.Streams.TryGetValue(entry.Path, out byte[]? streamBytes) && streamBytes != null) {
                bytes = streamBytes;
                return true;
            }

            if (compoundFile.Streams.TryGetValue(entry.Name, out streamBytes) && streamBytes != null) {
                bytes = streamBytes;
                return true;
            }

            string key = GetEntryKey(entry);
            if (compoundFile.Streams.TryGetValue(key, out streamBytes) && streamBytes != null) {
                bytes = streamBytes;
                return true;
            }

            bytes = Array.Empty<byte>();
            return false;
        }

        private static bool IsVbaProjectEntry(OfficeCompoundFileEntry entry) {
            if (entry.Name.Equals("_VBA_PROJECT_CUR", StringComparison.OrdinalIgnoreCase)
                || entry.Name.Equals("_VBA_PROJECT", StringComparison.OrdinalIgnoreCase)) {
                return true;
            }

            return entry.Path.IndexOf("/_VBA_PROJECT_CUR/", StringComparison.OrdinalIgnoreCase) >= 0
                || entry.Path.EndsWith("/_VBA_PROJECT_CUR", StringComparison.OrdinalIgnoreCase)
                || entry.Path.IndexOf("/VBA/", StringComparison.OrdinalIgnoreCase) >= 0
                || entry.Path.EndsWith("/VBA", StringComparison.OrdinalIgnoreCase);
        }

        private static LegacyXlsCompoundFeatureEntryRole ClassifyVbaProjectEntry(OfficeCompoundFileEntry entry) {
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

        private static bool IsOleObjectEntry(OfficeCompoundFileEntry entry) {
            return entry.Name.Equals("ObjectPool", StringComparison.OrdinalIgnoreCase)
                || entry.Name.Equals("Ole", StringComparison.OrdinalIgnoreCase)
                || entry.Name.Equals("\u0001Ole", StringComparison.OrdinalIgnoreCase)
                || entry.Name.Equals("Ole10Native", StringComparison.OrdinalIgnoreCase)
                || entry.Path.IndexOf("/ObjectPool/", StringComparison.OrdinalIgnoreCase) >= 0
                || entry.Path.EndsWith("/ObjectPool", StringComparison.OrdinalIgnoreCase)
                || entry.Path.IndexOf("/Ole10Native", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static LegacyXlsCompoundFeatureEntryRole ClassifyOleObjectEntry(OfficeCompoundFileEntry entry) {
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

        private static bool IsDigitalSignatureEntry(OfficeCompoundFileEntry entry) {
            return entry.Name.Equals("_signatures", StringComparison.OrdinalIgnoreCase)
                || entry.Name.Equals("_xmlsignatures", StringComparison.OrdinalIgnoreCase)
                || entry.Path.IndexOf("/_xmlsignatures/", StringComparison.OrdinalIgnoreCase) >= 0
                || entry.Path.EndsWith("/_xmlsignatures", StringComparison.OrdinalIgnoreCase);
        }

        private static LegacyXlsCompoundFeatureEntryRole ClassifyDigitalSignatureEntry(OfficeCompoundFileEntry entry) {
            if (entry.Name.Equals("_signatures", StringComparison.OrdinalIgnoreCase)) {
                return entry.IsStorage
                    ? LegacyXlsCompoundFeatureEntryRole.DigitalSignatureStorage
                    : LegacyXlsCompoundFeatureEntryRole.DigitalSignatureStream;
            }

            if (entry.Name.Equals("_xmlsignatures", StringComparison.OrdinalIgnoreCase)
                || entry.Path.EndsWith("/_xmlsignatures", StringComparison.OrdinalIgnoreCase)) {
                return LegacyXlsCompoundFeatureEntryRole.XmlDigitalSignatureStorage;
            }

            if (entry.Path.IndexOf("/_xmlsignatures/", StringComparison.OrdinalIgnoreCase) >= 0) {
                return LegacyXlsCompoundFeatureEntryRole.XmlDigitalSignatureStream;
            }

            return LegacyXlsCompoundFeatureEntryRole.Unknown;
        }
    }
}
