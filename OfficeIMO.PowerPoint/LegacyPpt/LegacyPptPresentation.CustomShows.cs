using OfficeIMO.PowerPoint.LegacyPpt.Diagnostics;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt {
    public sealed partial class LegacyPptPresentation {
        private const ushort RecordNamedShows = 0x0410;
        private const ushort RecordNamedShow = 0x0411;
        private const ushort RecordNamedShowSlides = 0x0412;

        private readonly List<LegacyPptCustomShow> _customShows = new();
        private bool _customShowsAreEditable = true;

        /// <summary>Gets named custom slide shows in document order.</summary>
        public IReadOnlyList<LegacyPptCustomShow> CustomShows => _customShows;

        internal bool CustomShowsAreEditable => _customShowsAreEditable;

        private void ParseNamedShows(LegacyPptRecord document,
            LegacyPptImportOptions options) {
            LegacyPptRecord[] containers = document.Children.Where(record =>
                record.Type == RecordNamedShows).ToArray();
            if (containers.Length == 0) return;
            bool ambiguousContainer = containers.Length != 1;
            if (ambiguousContainer) {
                _customShowsAreEditable = false;
                AddDiagnostic("PPT-CUSTOM-SHOW-LIST", LegacyPptDiagnosticSeverity.Warning,
                    "The document has multiple named-show lists; only the first can be projected and all remain preserve-only.",
                    containers[0].Offset);
            }

            LegacyPptRecord container = containers[0];
            if (container.Version != 0x0F || container.Instance != 0) {
                _customShowsAreEditable = false;
                AddDiagnostic("PPT-CUSTOM-SHOW-LIST", LegacyPptDiagnosticSeverity.Warning,
                    "The named-show list has an invalid record header and remains preserve-only.",
                    container.Offset);
                return;
            }
            foreach (LegacyPptRecord child in container.Children) {
                if (child.Type != RecordNamedShow) {
                    _customShowsAreEditable = false;
                    if (options.ReportUnsupportedContent) {
                        AddDiagnostic("PPT-CUSTOM-SHOW-CHILD",
                            LegacyPptDiagnosticSeverity.Warning,
                            $"The named-show list contains unexpected record 0x{child.Type:X4}; it remains preserve-only.",
                            child.Offset);
                    }
                    continue;
                }
                LegacyPptCustomShow? show = TryReadCustomShow(child,
                    ambiguousContainer, options);
                if (show == null) {
                    _customShowsAreEditable = false;
                } else {
                    _customShows.Add(show);
                    if (!show.IsEditable) _customShowsAreEditable = false;
                }
            }

            foreach (IGrouping<string, LegacyPptCustomShow> duplicate in _customShows
                         .GroupBy(show => show.Name, StringComparer.Ordinal)
                         .Where(group => group.Count() > 1)) {
                foreach (LegacyPptCustomShow show in duplicate) {
                    show.MarkStructurallyLossy();
                }
                _customShowsAreEditable = false;
                AddDiagnostic("PPT-CUSTOM-SHOW-NAME",
                    LegacyPptDiagnosticSeverity.Warning,
                    $"Named show '{duplicate.Key}' occurs more than once; its actions are ambiguous and remain preserve-only.",
                    duplicate.First().RecordOffset);
            }
        }

        private LegacyPptCustomShow? TryReadCustomShow(LegacyPptRecord container,
            bool inheritedStructuralLoss, LegacyPptImportOptions options) {
            if (container.Version != 0x0F || container.Instance != 0) {
                AddDiagnostic("PPT-CUSTOM-SHOW", LegacyPptDiagnosticSeverity.Warning,
                    "A named show has an invalid record header and remains preserve-only.",
                    container.Offset);
                return null;
            }
            LegacyPptRecord[] names = container.Children.Where(record =>
                record.Type == RecordCString && record.Instance == 0).ToArray();
            LegacyPptRecord[] slideAtoms = container.Children.Where(record =>
                record.Type == RecordNamedShowSlides).ToArray();
            bool structurallyLossy = inheritedStructuralLoss
                || names.Length != 1 || slideAtoms.Length > 1
                || container.Children.Any(record =>
                    record.Type == RecordCString
                        ? record.Instance != 0
                        : record.Type != RecordNamedShowSlides);
            if (names.Length != 1 || names[0].Version != 0
                || (names[0].PayloadLength & 1) != 0) {
                AddDiagnostic("PPT-CUSTOM-SHOW-NAME",
                    LegacyPptDiagnosticSeverity.Warning,
                    "A named show has no unique valid Unicode name and remains preserve-only.",
                    container.Offset);
                return null;
            }
            string name = names[0].ReadUtf16Text().TrimEnd('\0');
            if (name.Length == 0 || name.IndexOf('\0') >= 0) {
                AddDiagnostic("PPT-CUSTOM-SHOW-NAME",
                    LegacyPptDiagnosticSeverity.Warning,
                    "A named show has an empty or invalid name and remains preserve-only.",
                    names[0].Offset);
                structurallyLossy = true;
            }

            var slideIds = new List<uint>();
            if (slideAtoms.Length == 1) {
                LegacyPptRecord slides = slideAtoms[0];
                if (slides.Version != 0 || slides.Instance != 0
                    || (slides.PayloadLength & 3) != 0) {
                    AddDiagnostic("PPT-CUSTOM-SHOW-SLIDES",
                        LegacyPptDiagnosticSeverity.Warning,
                        $"Named show '{name}' has a malformed slide-id array and remains preserve-only.",
                        slides.Offset);
                    structurallyLossy = true;
                } else {
                    for (int offset = 0; offset < slides.PayloadLength; offset += 4) {
                        slideIds.Add(slides.ReadUInt32(offset));
                    }
                }
            } else if (slideAtoms.Length > 1) {
                AddDiagnostic("PPT-CUSTOM-SHOW-SLIDES",
                    LegacyPptDiagnosticSeverity.Warning,
                    $"Named show '{name}' has multiple slide-id arrays and remains preserve-only.",
                    container.Offset);
            }
            if (structurallyLossy && options.ReportUnsupportedContent) {
                AddDiagnostic("PPT-CUSTOM-SHOW-PRESERVE-ONLY",
                    LegacyPptDiagnosticSeverity.Warning,
                    $"Named show '{name}' contains unsupported record structure and remains preserve-only.",
                    container.Offset);
            }
            return new LegacyPptCustomShow(name, slideIds, container.Offset,
                structurallyLossy);
        }

        private void ValidateCustomShowSlideReferences() {
            var slideIds = new HashSet<uint>(_slides.Select(slide => slide.SlideId));
            foreach (LegacyPptCustomShow show in _customShows) {
                uint[] missing = show.SlideIds.Where(id => id == 0
                    || !slideIds.Contains(id)).Distinct().ToArray();
                if (missing.Length == 0) continue;
                show.MarkUnresolvedSlides();
                _customShowsAreEditable = false;
                AddDiagnostic("PPT-CUSTOM-SHOW-SLIDE-MISSING",
                    LegacyPptDiagnosticSeverity.Warning,
                    $"Named show '{show.Name}' references {missing.Length} missing slide identifier(s); those entries are ignored during projection.",
                    show.RecordOffset);
            }
        }
    }
}
