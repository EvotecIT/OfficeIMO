using OfficeIMO.PowerPoint.LegacyPpt.Diagnostics;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;

namespace OfficeIMO.PowerPoint.LegacyPpt {
    public sealed partial class LegacyPptPresentation {
        private const ushort RecordStyleTextProp9Atom = 0x0FAC;

        private LegacyPptRecord? ReadShapeStyle9(
            LegacyPptRecord shapeContainer,
            LegacyPptImportOptions options,
            out bool malformed) {
            malformed = false;
            LegacyPptRecord[] ppt9Tags = shapeContainer.Children
                .Where(record => record.Type == OfficeArtClientData)
                .SelectMany(record => record.Children)
                .Where(record => record.Type == RecordProgTags)
                .SelectMany(record => record.Children)
                .Where(IsPpt9BinaryTag)
                .ToArray();
            if (ppt9Tags.Length == 0) return null;
            if (ppt9Tags.Length != 1) {
                malformed = true;
                AddDiagnostic("PPT-TEXT-STYLE9-DUPLICATE",
                    LegacyPptDiagnosticSeverity.Warning,
                    "A shape has multiple PPT9 extended-text tags; they remain preserve-only.",
                    ppt9Tags[1].Offset);
                return null;
            }
            LegacyPptRecord tag = ppt9Tags[0];
            LegacyPptRecord[] dataBlobs = tag.Children.Where(record =>
                record.Type == RecordBinaryTagDataBlob).ToArray();
            if (dataBlobs.Length != 1 || dataBlobs[0].Version != 0
                || dataBlobs[0].Instance != 0) {
                malformed = true;
                AddDiagnostic("PPT-TEXT-STYLE9-DATA",
                    LegacyPptDiagnosticSeverity.Warning,
                    "A PPT9 extended-text tag has no unique valid data blob; it remains preserve-only.",
                    tag.Offset);
                return null;
            }
            try {
                LegacyPptRecord blob = dataBlobs[0];
                IReadOnlyList<LegacyPptRecord> records =
                    LegacyPptRecordReader.ReadSequence(
                        blob.CopyRecordBytes(), 8, blob.PayloadLength,
                        options, _recordBudget);
                if (records.Count != 1
                    || records[0].Type != RecordStyleTextProp9Atom) {
                    malformed = true;
                    AddDiagnostic("PPT-TEXT-STYLE9-CONTENT",
                        LegacyPptDiagnosticSeverity.Warning,
                        "A PPT9 extended-text data blob does not contain exactly one StyleTextProp9Atom; it remains preserve-only.",
                        blob.Offset);
                    return null;
                }
                return records[0];
            } catch (Exception exception) when (exception
                is InvalidDataException or OverflowException
                    or ArgumentOutOfRangeException) {
                malformed = true;
                AddDiagnostic("PPT-TEXT-STYLE9-TRUNCATED",
                    LegacyPptDiagnosticSeverity.Warning,
                    "A PPT9 extended-text tag is malformed or truncated; it remains preserve-only.",
                    dataBlobs[0].Offset);
                return null;
            }
        }
    }
}
