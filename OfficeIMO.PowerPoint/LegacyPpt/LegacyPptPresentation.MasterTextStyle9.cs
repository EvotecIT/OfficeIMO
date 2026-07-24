using OfficeIMO.PowerPoint.LegacyPpt.Diagnostics;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;

namespace OfficeIMO.PowerPoint.LegacyPpt {
    public sealed partial class LegacyPptPresentation {
        private const ushort RecordTextMasterStyle9Atom = 0x0FAD;

        private IReadOnlyDictionary<ushort, LegacyPptRecord>
            ReadMasterTextStyle9Records(LegacyPptRecord masterRecord,
                LegacyPptImportOptions options) {
            var result = new Dictionary<ushort, LegacyPptRecord>();
            LegacyPptRecord[] tags = masterRecord.Children
                .Where(record => record.Type == RecordProgTags)
                .SelectMany(record => record.Children)
                .Where(IsPpt9BinaryTag).ToArray();
            if (tags.Length > 1) {
                AddDiagnostic("PPT-TEXT-MASTER-STYLE9-DUPLICATE",
                    LegacyPptDiagnosticSeverity.Warning,
                    "A master has multiple PPT9 extended-style tags; they remain preserve-only.",
                    tags[1].Offset);
                return result;
            }
            if (tags.Length == 0) return result;
            LegacyPptRecord[] blobs = tags[0].Children.Where(record =>
                record.Type == RecordBinaryTagDataBlob).ToArray();
            if (blobs.Length != 1 || blobs[0].Version != 0
                || blobs[0].Instance != 0) {
                AddDiagnostic("PPT-TEXT-MASTER-STYLE9-DATA",
                    LegacyPptDiagnosticSeverity.Warning,
                    "A master PPT9 tag has no unique valid data blob; its text defaults remain preserve-only.",
                    tags[0].Offset);
                return result;
            }
            try {
                LegacyPptRecord blob = blobs[0];
                foreach (KeyValuePair<ushort, LegacyPptRecord> item in
                         CollectUniqueMasterTextStyle9Records(
                             LegacyPptRecordReader
                                 .ReadSequence(blob.CopyRecordBytes(), 8,
                                     blob.PayloadLength, options, _recordBudget)
                                 .Where(record => record.Type
                                     == RecordTextMasterStyle9Atom),
                             record => {
                        AddDiagnostic(
                            "PPT-TEXT-MASTER-STYLE9-TYPE-DUPLICATE",
                            LegacyPptDiagnosticSeverity.Warning,
                            $"A master has multiple TextMasterStyle9Atom records for instance {record.Instance}; that extended style remains preserve-only.",
                            record.Offset);
                             })) {
                    result.Add(item.Key, item.Value);
                }
            } catch (Exception exception) when (exception
                is InvalidDataException or OverflowException
                    or ArgumentOutOfRangeException) {
                AddDiagnostic("PPT-TEXT-MASTER-STYLE9-TRUNCATED",
                    LegacyPptDiagnosticSeverity.Warning,
                    "A master PPT9 tag is malformed or truncated; its extended text defaults remain preserve-only.",
                    blobs[0].Offset);
                result.Clear();
            }
            return result;
        }

        internal static IReadOnlyDictionary<ushort, LegacyPptRecord>
            CollectUniqueMasterTextStyle9Records(
                IEnumerable<LegacyPptRecord> records,
                Action<LegacyPptRecord>? duplicate = null) {
            var result = new Dictionary<ushort, LegacyPptRecord>();
            var ambiguousInstances = new HashSet<ushort>();
            foreach (LegacyPptRecord record in records) {
                if (ambiguousInstances.Contains(record.Instance)) {
                    continue;
                }
                if (result.ContainsKey(record.Instance)) {
                    duplicate?.Invoke(record);
                    result.Remove(record.Instance);
                    ambiguousInstances.Add(record.Instance);
                } else {
                    result.Add(record.Instance, record);
                }
            }
            return result;
        }
    }
}
