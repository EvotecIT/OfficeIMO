using OfficeIMO.PowerPoint.LegacyPpt.Diagnostics;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt {
    public sealed partial class LegacyPptPresentation {
        private void ParseSlideAtom(LegacyPptRecord slideRecord, LegacyPptSlide slide,
            LegacyPptImportOptions options) {
            LegacyPptRecord? record = slideRecord.Children.FirstOrDefault(child =>
                child.Type == RecordSlideAtom);
            if (!LegacyPptLayoutReader.TryReadSlideAtom(record, out LegacyPptSlideAtomData data)) {
                AddDiagnostic("PPT-SLIDE-ATOM-TRUNCATED", LegacyPptDiagnosticSeverity.Warning,
                    "A required SlideAtom is missing or truncated; layout and master inheritance remain preserve-only.",
                    record?.Offset ?? slideRecord.Offset);
                return;
            }

            ApplySlideAtom(slide, data);
            ReportSlideAtomIssues(data, record!.Offset, "slide", options);
        }

        private void ParseMasterSlideAtom(LegacyPptRecord masterRecord, LegacyPptMaster master,
            LegacyPptRecord? record, LegacyPptImportOptions options) {
            if (!LegacyPptLayoutReader.TryReadSlideAtom(record, out LegacyPptSlideAtomData data)) {
                AddDiagnostic("PPT-MASTER-SLIDE-ATOM-TRUNCATED",
                    LegacyPptDiagnosticSeverity.Warning,
                    "A required master SlideAtom is missing or truncated; its layout metadata remains preserve-only.",
                    record?.Offset ?? masterRecord.Offset);
                return;
            }

            master.LayoutType = data.RawLayoutType;
            master.Layout = data.Layout;
            master.SetLayoutPlaceholderTypes(data.PlaceholderTypes);
            master.FollowsMasterObjects = data.FollowsMasterObjects;
            master.FollowsMasterColorScheme = data.FollowsMasterColorScheme;
            master.FollowsMasterBackground = data.FollowsMasterBackground;
            ReportSlideAtomIssues(data, record!.Offset, "master", options);
        }

        private static void ApplySlideAtom(LegacyPptSlide slide, LegacyPptSlideAtomData data) {
            slide.LayoutType = data.RawLayoutType;
            slide.Layout = data.Layout;
            slide.SetLayoutPlaceholderTypes(data.PlaceholderTypes);
            slide.MasterId = data.MasterId;
            slide.NotesId = data.NotesId;
            slide.FollowsMasterObjects = data.FollowsMasterObjects;
            slide.FollowsMasterColorScheme = data.FollowsMasterColorScheme;
            slide.FollowsMasterBackground = data.FollowsMasterBackground;
        }

        private void ReportSlideAtomIssues(LegacyPptSlideAtomData data, long offset,
            string owner, LegacyPptImportOptions options) {
            if (data.HasInvalidLength) {
                AddDiagnostic("PPT-SLIDE-ATOM-LENGTH", LegacyPptDiagnosticSeverity.Warning,
                    $"A {owner} SlideAtom has a nonstandard length; its first 24 bytes were decoded.",
                    offset);
            }
            if (!data.Layout.HasValue) {
                AddDiagnostic("PPT-SLIDE-LAYOUT-TYPE", LegacyPptDiagnosticSeverity.Warning,
                    $"A {owner} uses undefined layout type 0x{data.RawLayoutType:X8}; the raw value remains available.",
                    offset);
            }
            if (data.HasInvalidPlaceholderType) {
                AddDiagnostic("PPT-SLIDE-LAYOUT-PLACEHOLDER",
                    LegacyPptDiagnosticSeverity.Warning,
                    $"A {owner} layout signature contains an undefined placeholder type; the slot remains preserve-only.",
                    offset);
            }
            if (data.HasReservedFlags && options.ReportUnsupportedContent) {
                AddDiagnostic("PPT-SLIDE-FLAGS-RESERVED", LegacyPptDiagnosticSeverity.Warning,
                    $"A {owner} SlideFlags value uses reserved bits; those bits remain preserved only.",
                    offset);
            }
        }

        private LegacyPptPlaceholder? ReadPlaceholder(LegacyPptRecord shapeContainer,
            LegacyPptImportOptions options) {
            LegacyPptRecord? record = shapeContainer.DescendantsAndSelf()
                .FirstOrDefault(child => child.Type == RecordPlaceholder);
            LegacyPptPlaceholder? placeholder = LegacyPptLayoutReader.ReadPlaceholder(record,
                out LegacyPptPlaceholderReadStatus status);
            if (status == LegacyPptPlaceholderReadStatus.Invalid
                && options.ReportUnsupportedContent) {
                AddDiagnostic("PPT-PLACEHOLDER-INVALID", LegacyPptDiagnosticSeverity.Warning,
                    "A PlaceholderAtom has an invalid length, identifier, kind, or size and remains preserve-only.",
                    record?.Offset ?? shapeContainer.Offset);
            }
            return placeholder;
        }
    }
}
