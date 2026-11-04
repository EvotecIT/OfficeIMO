using OfficeIMO.PowerPoint.LegacyPpt.Diagnostics;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt {
    internal sealed partial class LegacyPptDocument {
        private void ParseDocumentSettings(LegacyPptRecord? atom) {
            if (atom == null) return;
            if (atom.PayloadLength >= 8) {
                int width = atom.ReadInt32(0);
                int height = atom.ReadInt32(4);
                if (width > 0 && height > 0) {
                    SlideWidth = width;
                    SlideHeight = height;
                }
            }
            if (atom.PayloadLength != LegacyPptDocumentAtomReader.PayloadLength) {
                AddDiagnostic("ppt-document-atom-length", LegacyPptDiagnosticSeverity.Warning,
                    $"DocumentAtom has {atom.PayloadLength} payload bytes instead of 40; complete document settings remain preserve-only.",
                    atom.Offset);
            }
            DocumentSettings = LegacyPptDocumentAtomReader.Read(atom);
            if (DocumentSettings == null) return;

            LegacyPptDocumentSettings settings = DocumentSettings;
            if (settings.SlideWidth < 576 || settings.SlideHeight < 576
                || settings.NotesWidth < 576 || settings.NotesHeight < 576
                || settings.ServerZoomNumerator <= 0 || settings.ServerZoomDenominator <= 0
                || settings.FirstSlideNumber >= 10000 || !settings.SlideSizeType.HasValue) {
                AddDiagnostic("ppt-document-atom-value", LegacyPptDiagnosticSeverity.Warning,
                    "DocumentAtom contains an out-of-range page size, zoom ratio, slide number, or slide-size type.",
                    atom.Offset);
            }
            for (int index = 36; index < 40; index++) {
                if (atom.ReadByte(index) <= 1) continue;
                AddDiagnostic("ppt-document-atom-boolean", LegacyPptDiagnosticSeverity.Warning,
                    "DocumentAtom contains a Boolean field other than zero or one.", atom.Offset);
                break;
            }
        }
    }
}
