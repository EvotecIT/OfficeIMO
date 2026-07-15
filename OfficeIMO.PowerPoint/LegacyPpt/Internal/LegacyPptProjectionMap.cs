using System.Collections.ObjectModel;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>Links projected Open XML slides and shapes back to their original binary persist records.</summary>
    internal sealed class LegacyPptProjectionMap {
        private readonly IReadOnlyDictionary<string, LegacyPptSlideProjection> _slidesByPartUri;
        private readonly IReadOnlyDictionary<string, uint> _masterIdsByLayoutPartUri;

        private LegacyPptProjectionMap(IReadOnlyList<LegacyPptSlideProjection> slides,
            IReadOnlyDictionary<string, uint> masterIdsByLayoutPartUri) {
            Slides = new ReadOnlyCollection<LegacyPptSlideProjection>(slides.ToArray());
            _slidesByPartUri = new ReadOnlyDictionary<string, LegacyPptSlideProjection>(slides.ToDictionary(
                slide => slide.SlidePartUri, StringComparer.Ordinal));
            _masterIdsByLayoutPartUri = new ReadOnlyDictionary<string, uint>(
                masterIdsByLayoutPartUri.ToDictionary(pair => pair.Key, pair => pair.Value,
                    StringComparer.Ordinal));
        }

        internal IReadOnlyList<LegacyPptSlideProjection> Slides { get; }

        internal bool TryGetSlide(PowerPointSlide slide, out LegacyPptSlideProjection? projection) {
            if (slide == null) throw new ArgumentNullException(nameof(slide));
            return _slidesByPartUri.TryGetValue(slide.SlidePart.Uri.ToString(), out projection);
        }

        internal bool TryGetMasterId(PowerPointSlide slide, out uint masterId) {
            if (slide == null) throw new ArgumentNullException(nameof(slide));
            masterId = 0;
            SlideLayoutPart? layoutPart = slide.SlidePart.SlideLayoutPart;
            return layoutPart != null
                && _masterIdsByLayoutPartUri.TryGetValue(layoutPart.Uri.ToString(), out masterId);
        }

        internal bool IsProjectedLayoutPart(string partUri) =>
            partUri != null && _masterIdsByLayoutPartUri.ContainsKey(partUri);

        internal static LegacyPptProjectionMap Create(PowerPointPresentation presentation,
            LegacyPptPresentation legacy) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            if (legacy == null) throw new ArgumentNullException(nameof(legacy));
            if (presentation.Slides.Count != legacy.Slides.Count) {
                throw new InvalidDataException("The projected slide count does not match the binary source slide count.");
            }

            var slides = new List<LegacyPptSlideProjection>(legacy.Slides.Count);
            for (int slideIndex = 0; slideIndex < legacy.Slides.Count; slideIndex++) {
                PowerPointSlide projectedSlide = presentation.Slides[slideIndex];
                LegacyPptSlide sourceSlide = legacy.Slides[slideIndex];
                PowerPointShape[] projectedShapes = projectedSlide.Shapes.ToArray();
                LegacyPptShape[] sourceShapes = sourceSlide.Shapes
                    .Where(shape => shape.Kind != LegacyPptShapeKind.Unsupported)
                    .ToArray();
                if (projectedShapes.Length != sourceShapes.Length) {
                    throw new InvalidDataException(
                        $"Projected slide {slideIndex + 1} has {projectedShapes.Length} editable shapes, "
                        + $"but the binary source exposed {sourceShapes.Length}.");
                }

                var shapes = new List<LegacyPptShapeProjection>(projectedShapes.Length);
                for (int shapeIndex = 0; shapeIndex < projectedShapes.Length; shapeIndex++) {
                    uint? openXmlShapeId = projectedShapes[shapeIndex].Id;
                    if (!openXmlShapeId.HasValue) {
                        throw new InvalidDataException(
                            $"Projected slide {slideIndex + 1}, shape {shapeIndex + 1} has no Open XML shape id.");
                    }
                    LegacyPptShape sourceShape = sourceShapes[shapeIndex];
                    string? textFormattingFingerprint = sourceShape.TextBody.HasStyleRecord
                        && projectedShapes[shapeIndex].Element is DocumentFormat.OpenXml.Presentation.Shape projectedTextShape
                        ? LegacyPptTextProjection.CreateFormattingFingerprint(projectedTextShape.TextBody)
                        : null;
                    shapes.Add(new LegacyPptShapeProjection(openXmlShapeId.Value, sourceShape.ShapeId,
                        sourceShape.RecordOffset, sourceShape.Kind, sourceShape.Bounds, sourceShape.Text,
                        textFormattingFingerprint));
                }
                slides.Add(new LegacyPptSlideProjection(projectedSlide.SlidePart.Uri.ToString(),
                    sourceSlide.PersistId, sourceSlide.SlideId, sourceSlide.MasterId,
                    sourceSlide.Hidden, sourceSlide.HeaderFooter, shapes,
                    sourceSlide.NotesPage == null
                        ? null
                        : new LegacyPptNotesProjection(sourceSlide.NotesPage.PersistId,
                            sourceSlide.NotesPage.NotesId, sourceSlide.NotesPage.Text)));
            }
            return new LegacyPptProjectionMap(slides, CreateLayoutMasterMap(presentation, legacy));
        }

        private static IReadOnlyDictionary<string, uint> CreateLayoutMasterMap(
            PowerPointPresentation presentation, LegacyPptPresentation legacy) {
            SlideMasterPart[] masterParts = presentation.OpenXmlDocument.PresentationPart?
                .SlideMasterParts.ToArray() ?? Array.Empty<SlideMasterPart>();
            var result = new Dictionary<string, uint>(StringComparer.Ordinal);
            var masterIdsByName = legacy.Masters.ToDictionary(master =>
                $"Binary {(master.IsMainMaster ? "Main" : "Title")} Master {master.MasterId:X8}",
                master => master.MasterId, StringComparer.Ordinal);
            foreach (SlideMasterPart masterPart in masterParts) {
                foreach (SlideLayoutPart layoutPart in masterPart.SlideLayoutParts) {
                    string? name = layoutPart.SlideLayout?.CommonSlideData?.Name?.Value;
                    if (name == null) continue;
                    foreach (KeyValuePair<string, uint> candidate in masterIdsByName) {
                        if (string.Equals(name, candidate.Key, StringComparison.Ordinal)
                            || name.StartsWith(candidate.Key + " / ", StringComparison.Ordinal)) {
                            result[layoutPart.Uri.ToString()] = candidate.Value;
                            break;
                        }
                    }
                }
            }
            return result;
        }
    }

    /// <summary>Maps one projected slide to its binary persist object.</summary>
    internal sealed class LegacyPptSlideProjection {
        private readonly IReadOnlyDictionary<uint, LegacyPptShapeProjection> _shapesByOpenXmlId;

        internal LegacyPptSlideProjection(string slidePartUri, uint persistId, uint slideId, uint masterId,
            bool hidden, LegacyPptHeaderFooterSettings? headerFooter,
            IReadOnlyList<LegacyPptShapeProjection> shapes, LegacyPptNotesProjection? notes) {
            SlidePartUri = slidePartUri ?? throw new ArgumentNullException(nameof(slidePartUri));
            PersistId = persistId;
            SlideId = slideId;
            MasterId = masterId;
            Hidden = hidden;
            HeaderFooter = headerFooter;
            Notes = notes;
            Shapes = new ReadOnlyCollection<LegacyPptShapeProjection>(shapes.ToArray());
            _shapesByOpenXmlId = new ReadOnlyDictionary<uint, LegacyPptShapeProjection>(shapes.ToDictionary(
                shape => shape.OpenXmlShapeId));
        }

        internal string SlidePartUri { get; }

        internal uint PersistId { get; }

        internal uint SlideId { get; }

        internal uint MasterId { get; }

        internal bool Hidden { get; }

        internal LegacyPptHeaderFooterSettings? HeaderFooter { get; }

        internal LegacyPptNotesProjection? Notes { get; }

        internal IReadOnlyList<LegacyPptShapeProjection> Shapes { get; }

        internal bool TryGetShape(uint openXmlShapeId, out LegacyPptShapeProjection? projection) =>
            _shapesByOpenXmlId.TryGetValue(openXmlShapeId, out projection);
    }

    /// <summary>Maps projected speaker-note text to its binary NotesContainer.</summary>
    internal sealed class LegacyPptNotesProjection {
        internal LegacyPptNotesProjection(uint persistId, uint notesId, string text) {
            PersistId = persistId;
            NotesId = notesId;
            Text = text ?? string.Empty;
        }

        internal uint PersistId { get; }

        internal uint NotesId { get; }

        internal string Text { get; }
    }

    /// <summary>Maps one projected Open XML shape to its OfficeArt shape container.</summary>
    internal sealed class LegacyPptShapeProjection {
        internal LegacyPptShapeProjection(uint openXmlShapeId, uint officeArtShapeId, long recordOffset,
            LegacyPptShapeKind kind, LegacyPptBounds bounds, string text,
            string? textFormattingFingerprint) {
            OpenXmlShapeId = openXmlShapeId;
            OfficeArtShapeId = officeArtShapeId;
            RecordOffset = recordOffset;
            Kind = kind;
            Bounds = bounds;
            Text = text ?? string.Empty;
            TextFormattingFingerprint = textFormattingFingerprint;
        }

        internal uint OpenXmlShapeId { get; }

        internal uint OfficeArtShapeId { get; }

        internal long RecordOffset { get; }

        internal LegacyPptShapeKind Kind { get; }

        internal LegacyPptBounds Bounds { get; }

        internal string Text { get; }

        internal string? TextFormattingFingerprint { get; }
    }
}
