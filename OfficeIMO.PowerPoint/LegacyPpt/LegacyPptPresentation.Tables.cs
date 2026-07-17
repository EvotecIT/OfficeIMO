using OfficeIMO.PowerPoint.LegacyPpt.Diagnostics;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt {
    public sealed partial class LegacyPptPresentation {
        private byte? ReadOfficeImoTableStyleFlags(
            LegacyPptRecord descriptor,
            LegacyPptImportOptions options) {
            LegacyPptRecord[] tags = descriptor.Children
                .Where(record => record.Type == OfficeArtClientData)
                .SelectMany(record => record.Children)
                .Where(record => record.Type == RecordProgTags)
                .SelectMany(record => record.Children)
                .Where(IsOfficeImoTableStyleTag)
                .ToArray();
            if (tags.Length == 0) return null;
            if (tags.Length != 1) {
                if (options.ReportUnsupportedContent) {
                    AddDiagnostic("PPT-TABLE-STYLE-TAG-DUPLICATE",
                        LegacyPptDiagnosticSeverity.Warning,
                        "A native table has duplicate OfficeIMO style metadata; its semantic style flags were ignored.",
                        tags[1].Offset);
                }
                return null;
            }

            LegacyPptRecord[] data = tags[0].Children.Where(record =>
                record.Type == RecordBinaryTagDataBlob).ToArray();
            if (data.Length != 1 || data[0].Version != 0
                || data[0].Instance != 0 || data[0].PayloadLength != 2
                || data[0].ReadByte(0) != LegacyPptTableMetadata.Version) {
                if (options.ReportUnsupportedContent) {
                    AddDiagnostic("PPT-TABLE-STYLE-TAG-DATA",
                        LegacyPptDiagnosticSeverity.Warning,
                        "A native table has malformed OfficeIMO style metadata; its semantic style flags were ignored.",
                        tags[0].Offset);
                }
                return null;
            }
            return unchecked((byte)(data[0].ReadByte(1)
                & LegacyPptTableMetadata.KnownStyleFlagsMask));
        }

        private static bool IsOfficeImoTableStyleTag(
            LegacyPptRecord record) {
            if (record.Type != RecordProgBinaryTag) return false;
            LegacyPptRecord? name = record.Children.FirstOrDefault(child =>
                child.Type == RecordCString && child.Instance == 0);
            return name != null && TryReadUnicodeString(name,
                out string value) && string.Equals(value,
                LegacyPptTableMetadata.TagName,
                StringComparison.Ordinal);
        }
    }
}
