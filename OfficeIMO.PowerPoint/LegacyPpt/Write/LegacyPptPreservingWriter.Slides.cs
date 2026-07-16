using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptPreservingWriter {
        private static bool TryRewriteSlide(PowerPointSlide slide,
            LegacyPptRecord slideRecord,
            IReadOnlyDictionary<uint, ProjectedShapeEdit> editsByOfficeArtId,
            bool? hidden, bool? followsMasterObjects,
            uint? layoutType,
            IReadOnlyList<byte>? layoutPlaceholderTypes,
            bool rewriteTransition,
            LegacyPptWriter.LegacyPptWriterTransition? transition,
            LegacyPptWriter.LegacyPptWriterSoundCatalog soundCatalog,
            bool rewriteHeaderFooter,
            LegacyPptWriter.LegacyPptWriterHeaderFooter? headerFooter,
            bool rewriteComments,
            IReadOnlyList<LegacyPptWriter.LegacyPptWriterComment> comments,
            out RecordRewrite result) {
            bool hasSlideShowInfo = slideRecord.Children.Any(child =>
                child.Type == RecordSlideShowSlideInfoAtom);
            bool hasHeaderFooter = slideRecord.Children.Any(child =>
                child.Type == RecordHeadersFooters && child.Instance == 0);
            bool rewriteSlideShowInfo = hidden.HasValue || rewriteTransition;
            bool needsSlideShowInfo = slide.Hidden || transition != null;
            bool patchedSlideShowInfo = !rewriteSlideShowInfo;
            bool patchedHeaderFooter = !rewriteHeaderFooter;
            bool patchedComments = !rewriteComments;
            bool rewritePlaceholderSignature = layoutPlaceholderTypes != null
                || editsByOfficeArtId.Values.Any(edit =>
                    edit.RewritePlaceholder);
            byte[]? placeholderTypes = layoutPlaceholderTypes?.ToArray();
            if (rewritePlaceholderSignature && placeholderTypes == null) {
                IReadOnlyList<PowerPointShape> shapes = LegacyPptWriter
                    .ReadSlideShapesForWrite(slide, out string? layoutReason);
                if (layoutReason != null) {
                    result = new RecordRewrite(slideRecord.CopyRecordBytes(),
                        changed: false, patchedShapeCount: 0);
                    return false;
                }
                placeholderTypes = LegacyPptWriter
                    .BuildLayoutPlaceholderTypes(slide, shapes);
            }
            bool changed = false;
            int patchedShapeCount = 0;
            var children = new List<byte[]>(slideRecord.Children.Count + 1);
            foreach (LegacyPptRecord child in slideRecord.Children) {
                if (rewriteComments && child.Type == LegacyPptWriter.RecordProgTags) {
                    if (!TryRewriteProgrammableTags(child, comments,
                            out byte[]? rewrittenProgTags)) {
                        result = new RecordRewrite(slideRecord.CopyRecordBytes(),
                            changed: false, patchedShapeCount: 0);
                        return false;
                    }
                    if (rewrittenProgTags != null) children.Add(rewrittenProgTags);
                    patchedComments = true;
                    changed = true;
                } else if (rewriteHeaderFooter && child.Type == RecordHeadersFooters
                    && child.Instance == 0) {
                    if (headerFooter != null) {
                        children.Add(LegacyPptWriter.BuildHeaderFooterRecord(
                            headerFooter, instance: 0, allowHeader: false));
                    }
                    patchedHeaderFooter = true;
                    changed = true;
                } else if (rewriteSlideShowInfo
                           && child.Type == RecordSlideShowSlideInfoAtom) {
                    if (rewriteTransition && needsSlideShowInfo) {
                        children.Add(LegacyPptWriter.PatchSlideShowInfo(
                            child.CopyRecordBytes(), slide, soundCatalog));
                    } else if (!rewriteTransition) {
                        children.Add(PatchHiddenState(child.CopyRecordBytes(),
                            slide.Hidden));
                    }
                    patchedSlideShowInfo = true;
                    changed = true;
                } else if ((followsMasterObjects.HasValue
                               || layoutType.HasValue
                               || rewritePlaceholderSignature)
                           && child.Type == RecordSlideAtom
                           && child.PayloadLength >= 22) {
                    byte[] atom = child.CopyRecordBytes();
                    if (layoutType.HasValue) {
                        WriteUInt32(atom, 8, layoutType.Value);
                    }
                    if (placeholderTypes != null) {
                        Buffer.BlockCopy(placeholderTypes, 0, atom, 12,
                            placeholderTypes.Length);
                    }
                    ushort flags = child.ReadUInt16(20);
                    if (followsMasterObjects.HasValue) {
                        flags = followsMasterObjects.Value
                            ? unchecked((ushort)(flags | 0x0001))
                            : unchecked((ushort)(flags & ~0x0001));
                    }
                    WriteUInt16(atom, 28, flags);
                    children.Add(atom);
                    changed = true;
                } else {
                    RecordRewrite childResult = RewriteRecord(child, editsByOfficeArtId);
                    children.Add(childResult.Bytes);
                    changed |= childResult.Changed;
                    patchedShapeCount = checked(patchedShapeCount
                        + childResult.PatchedShapeCount);
                }
                if (rewriteSlideShowInfo && needsSlideShowInfo && !hasSlideShowInfo
                    && child.Type == RecordSlideAtom) {
                    children.Add(rewriteTransition
                        ? LegacyPptWriter.BuildSlideShowInfoRecord(slide,
                            soundCatalog)
                        : BuildSlideShowInfo(slide.Hidden));
                    patchedSlideShowInfo = true;
                    changed = true;
                }
                if (rewriteHeaderFooter && headerFooter != null && !hasHeaderFooter
                    && (child.Type == RecordSlideShowSlideInfoAtom
                        || !hasSlideShowInfo && child.Type == RecordSlideAtom)) {
                    children.Add(LegacyPptWriter.BuildHeaderFooterRecord(headerFooter,
                        instance: 0, allowHeader: false));
                    patchedHeaderFooter = true;
                    changed = true;
                }
            }
            if (rewriteComments && !patchedComments && comments.Count > 0) {
                children.Add(LegacyPptWriter.BuildCommentProgrammableTagsRecord(comments));
                patchedComments = true;
                changed = true;
            }
            if (!patchedSlideShowInfo || !patchedHeaderFooter || !patchedComments) {
                result = new RecordRewrite(slideRecord.CopyRecordBytes(),
                    changed: false, patchedShapeCount: 0);
                return false;
            }
            result = changed
                ? new RecordRewrite(BuildRecord(slideRecord.Version, slideRecord.Instance,
                    slideRecord.Type, Concat(children)), changed: true, patchedShapeCount)
                : new RecordRewrite(slideRecord.CopyRecordBytes(),
                    changed: false, patchedShapeCount: 0);
            return true;
        }

        private static bool CommentsEqual(IReadOnlyList<LegacyPptComment> source,
            IReadOnlyList<LegacyPptWriter.LegacyPptWriterComment> current) {
            if (source.Count != current.Count) return false;
            for (int index = 0; index < source.Count; index++) {
                LegacyPptComment left = source[index];
                LegacyPptWriter.LegacyPptWriterComment right = current[index];
                if (left.Index != right.Index || left.X != right.X || left.Y != right.Y
                    || !string.Equals(left.Author, right.Author, StringComparison.Ordinal)
                    || !string.Equals(left.Initials, right.Initials, StringComparison.Ordinal)
                    || !string.Equals(left.Text, right.Text, StringComparison.Ordinal)
                    || left.CreatedAtUtc != right.CreatedAtUtc) {
                    return false;
                }
            }
            return true;
        }
    }
}
