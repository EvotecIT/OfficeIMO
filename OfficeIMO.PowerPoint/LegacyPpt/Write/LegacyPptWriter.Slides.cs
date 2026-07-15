using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        private static byte[] BuildSlideRecord(LegacyPptRecord prototype, PowerPointSlide slide,
            IReadOnlyList<PowerPointShape> shapes, uint drawingId, uint? masterIdRef,
            uint? notesIdRef, IReadOnlyList<LegacyPptWriterComment> comments,
            LegacyPptWriterInteractionCatalog interactionCatalog) {
            var children = new List<byte[]>();
            bool hasSlideShowInfo = false;
            if (!TryReadTransition(slide, interactionCatalog.Sounds,
                    out LegacyPptWriterTransition? transition,
                    out string? transitionReason)) {
                throw new NotSupportedException(transitionReason);
            }
            bool needsSlideShowInfo = slide.Hidden || transition != null;
            bool hasHeaderFooter = prototype.Children.Any(child =>
                child.Type == RecordHeadersFooters && child.Instance == 0);
            LegacyPptWriterHeaderFooter? headerFooter = ReadSlideHeaderFooter(slide);
            bool wroteComments = false;
            foreach (LegacyPptRecord child in prototype.Children) {
                if (child.Type == RecordSlideAtom) {
                    byte[] atom = child.CopyRecordBytes();
                    LegacyPptSlideLayoutType layoutType = MapSlideLayout(slide, shapes);
                    WriteUInt32(atom, 8, (uint)layoutType);
                    byte[] placeholderTypes = BuildLayoutPlaceholderTypes(slide, shapes);
                    for (int index = 0; index < placeholderTypes.Length; index++) {
                        atom[12 + index] = placeholderTypes[index];
                    }
                    if (masterIdRef.HasValue) WriteUInt32(atom, 20, masterIdRef.Value);
                    WriteUInt32(atom, 24, notesIdRef.GetValueOrDefault());
                    children.Add(atom);
                    if (headerFooter != null && !hasHeaderFooter
                        && !prototype.Children.Any(candidate =>
                            candidate.Type == RecordSlideShowSlideInfoAtom)) {
                        children.Add(BuildHeaderFooterRecord(headerFooter,
                            instance: 0, allowHeader: false));
                    }
                } else if (child.Type == RecordDrawing) {
                    children.Add(BuildDrawingRecord(prototype, shapes, drawingId,
                        interactionCatalog));
                } else if (child.Type == RecordSlideShowSlideInfoAtom) {
                    if (needsSlideShowInfo) {
                        children.Add(PatchSlideShowInfo(child.CopyRecordBytes(), slide,
                            interactionCatalog.Sounds));
                    }
                    hasSlideShowInfo = true;
                    if (headerFooter != null && !hasHeaderFooter) {
                        children.Add(BuildHeaderFooterRecord(headerFooter,
                            instance: 0, allowHeader: false));
                    }
                } else if (child.Type == RecordHeadersFooters
                           && child.Instance == 0) {
                    children.Add(BuildHeaderFooterRecord(
                        headerFooter ?? LegacyPptWriterHeaderFooter.Empty,
                        instance: 0, allowHeader: false));
                } else if (child.Type == RecordProgTags) {
                    if (comments.Count > 0) {
                        children.Add(BuildCommentProgrammableTagsRecord(comments));
                        wroteComments = true;
                    }
                } else {
                    children.Add(child.CopyRecordBytes());
                }
            }
            if (comments.Count > 0 && !wroteComments) {
                children.Add(BuildCommentProgrammableTagsRecord(comments));
            }
            if (needsSlideShowInfo && !hasSlideShowInfo) {
                int slideAtomIndex = prototype.Children.TakeWhile(child =>
                    child.Type != RecordSlideAtom).Count();
                children.Insert(Math.Min(children.Count, slideAtomIndex + 1),
                    BuildSlideShowInfoRecord(slide, interactionCatalog.Sounds));
            }
            return BuildContainer(RecordSlide, instance: 0, children);
        }
    }
}
