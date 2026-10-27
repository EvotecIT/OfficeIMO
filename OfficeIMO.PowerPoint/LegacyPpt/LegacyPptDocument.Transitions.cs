using OfficeIMO.PowerPoint.LegacyPpt.Diagnostics;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt {
    public sealed partial class LegacyPptDocument {
        private void ParseSlideShowInfo(LegacyPptRecord slideRecord,
            LegacyPptSlide slide, LegacyPptImportOptions options) {
            LegacyPptRecord? atom = slideRecord.Children.FirstOrDefault(record =>
                record.Type == RecordSlideShowSlideInfoAtom);
            if (atom == null) return;
            if (atom.PayloadLength < 16) {
                AddDiagnostic("ppt-transition-truncated",
                    LegacyPptDiagnosticSeverity.Warning,
                    $"Slide {slide.SlideId} has a truncated slide-show information atom.",
                    atom.Offset);
                return;
            }

            var transition = new LegacyPptTransition(atom.ReadInt32(0),
                atom.ReadUInt32(4), atom.ReadByte(8), atom.ReadByte(9),
                atom.ReadUInt16(10), atom.ReadByte(12));
            slide.Transition = transition;
            slide.Hidden = transition.Hidden;
            if (transition.SlideTimeMilliseconds < 0
                || transition.SlideTimeMilliseconds > 86399000) {
                AddDiagnostic("ppt-transition-time",
                    LegacyPptDiagnosticSeverity.Warning,
                    $"Slide {slide.SlideId} has an invalid automatic-advance delay.",
                    atom.Offset);
            }
            if (transition.Speed > 2) {
                AddDiagnostic("ppt-transition-speed",
                    LegacyPptDiagnosticSeverity.Warning,
                    $"Slide {slide.SlideId} has undefined transition speed {transition.Speed}.",
                    atom.Offset);
            }
            if (options.ReportUnsupportedContent
                && (transition.RawFlags & 0xEAAA) != 0) {
                AddDiagnostic("ppt-transition-reserved",
                    LegacyPptDiagnosticSeverity.Warning,
                    $"Slide {slide.SlideId} has nonzero reserved transition flags.",
                    atom.Offset);
            }
            if (options.ReportUnsupportedContent && transition.PlaySound
                && transition.StopSound) {
                AddDiagnostic("ppt-transition-sound-flags",
                    LegacyPptDiagnosticSeverity.Warning,
                    $"Slide {slide.SlideId} both stops the current sound and starts a new sound; Open XML exposes the start action while the combined binary flags remain preserved until the transition is edited.",
                    atom.Offset);
            }
        }
    }
}
