using OfficeIMO.PowerPoint.LegacyPpt.Diagnostics;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt {
    public sealed partial class LegacyPptPresentation {
        private const ushort RecordHeadersFooters = 0x0FD9;
        private const ushort RecordHeadersFootersAtom = 0x0FDA;
        private const ushort RecordCString = 0x0FBA;

        /// <summary>Gets the document defaults for headers and footers on presentation slides.</summary>
        public LegacyPptHeaderFooterSettings? SlideHeaderFooterDefaults { get; private set; }

        /// <summary>Gets the document defaults for headers and footers on notes and handout pages.</summary>
        public LegacyPptHeaderFooterSettings? NotesHeaderFooterDefaults { get; private set; }

        private void ParseDocumentHeaderFooterSettings(LegacyPptRecord document,
            LegacyPptImportOptions options) {
            SlideHeaderFooterDefaults = ReadHeaderFooterSettings(document, instance: 3,
                "slide defaults", allowHeader: false, options);
            NotesHeaderFooterDefaults = ReadHeaderFooterSettings(document, instance: 4,
                "notes defaults", allowHeader: true, options);
        }

        private LegacyPptHeaderFooterSettings? ReadHeaderFooterSettings(
            LegacyPptRecord owner, ushort instance, string ownerDescription,
            bool allowHeader, LegacyPptImportOptions options) {
            LegacyPptRecord? container = owner.Children.FirstOrDefault(record =>
                record.Type == RecordHeadersFooters && record.Instance == instance);
            if (container == null) return null;

            LegacyPptRecord? atom = container.Children.FirstOrDefault(record =>
                record.Type == RecordHeadersFootersAtom);
            if (atom == null || atom.PayloadLength < 4) {
                AddDiagnostic("PPT-HEADER-FOOTER-ATOM",
                    LegacyPptDiagnosticSeverity.Warning,
                    $"The {ownerDescription} header/footer container has no complete atom and remains preserve-only.",
                    atom?.Offset ?? container.Offset);
                return null;
            }

            short formatId = atom.ReadInt16(0);
            ushort flags = atom.ReadUInt16(2);
            if (formatId < 0 || formatId > 13) {
                AddDiagnostic("PPT-HEADER-FOOTER-FORMAT",
                    LegacyPptDiagnosticSeverity.Warning,
                    $"The {ownerDescription} date/time format id {formatId} is outside the defined range.",
                    atom.Offset);
            }
            if (options.ReportUnsupportedContent && (flags & 0xFFC0) != 0) {
                AddDiagnostic("PPT-HEADER-FOOTER-RESERVED",
                    LegacyPptDiagnosticSeverity.Warning,
                    $"The {ownerDescription} header/footer atom has nonzero reserved flags that remain preserve-only.",
                    atom.Offset);
            }
            if (!allowHeader && options.ReportUnsupportedContent && (flags & 0x0010) != 0) {
                AddDiagnostic("PPT-SLIDE-HEADER-IGNORED",
                    LegacyPptDiagnosticSeverity.Information,
                    "The slide header flag is ignored by the binary PowerPoint format.", atom.Offset);
            }

            string userDate = ReadHeaderFooterText(container, instance: 0);
            string header = ReadHeaderFooterText(container, instance: 1);
            string footer = ReadHeaderFooterText(container, instance: 2);
            return new LegacyPptHeaderFooterSettings(formatId, flags,
                userDate, header, footer);
        }

        private static string ReadHeaderFooterText(LegacyPptRecord container,
            ushort instance) {
            LegacyPptRecord? atom = container.Children.FirstOrDefault(record =>
                record.Type == RecordCString && record.Instance == instance);
            if (atom == null || (atom.PayloadLength & 1) != 0) return string.Empty;
            return atom.ReadUtf16Text().TrimEnd('\0');
        }
    }
}
