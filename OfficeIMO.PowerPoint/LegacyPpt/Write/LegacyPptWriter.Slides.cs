using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        private static byte[] BuildSlideRecord(LegacyPptRecord prototype, PowerPointSlide slide,
            IReadOnlyList<PowerPointShape> shapes, uint drawingId, uint? masterIdRef,
            uint? notesIdRef, IReadOnlyList<LegacyPptWriterComment> comments,
            LegacyPptWriterInteractionCatalog interactionCatalog,
            LegacyPptWriterAnimationCatalog animationCatalog,
            LegacyPptWriterMediaCatalog mediaCatalog,
            LegacyPptWriterOleObjectCatalog oleCatalog,
            LegacyPptWriterPictureCatalog pictureCatalog,
            LegacyPptWriterFontCatalog fonts,
            LegacyPptWriterPictureBulletCatalog pictureBullets,
            bool layoutIsIndependentMaster = false) {
            var children = new List<byte[]>();
            ThemeOverridePart? themePart = slide.SlidePart.ThemeOverridePart
                ?? slide.SlidePart.SlideLayoutPart?.ThemeOverridePart;
            P.ColorMapOverride? colorMap = slide.SlidePart.Slide?
                .ColorMapOverride
                ?? slide.SlidePart.SlideLayoutPart?.SlideLayout?
                    .ColorMapOverride;
            IReadOnlyList<byte[]> roundTripThemeRecords =
                BuildRoundTripThemeRecords(themePart?.ThemeOverride, colorMap);
            A.ColorScheme? overrideColors = themePart?.ThemeOverride?
                .ColorScheme;
            LegacyPptWriterColorScheme? classicOverride = overrideColors == null
                ? null
                : ReadColorScheme(overrideColors);
            if (!TryReadBackground(slide, out LegacyPptWriterBackground? background,
                    out string? backgroundReason)) {
                throw new NotSupportedException(backgroundReason);
            }
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
            bool wroteClassicOverride = false;
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
                    ushort slideFlags = ReadUInt16(atom, 28);
                    slideFlags = background == null
                        ? unchecked((ushort)(slideFlags | 0x0004))
                        : unchecked((ushort)(slideFlags & ~0x0004));
                    slideFlags = classicOverride == null
                        ? unchecked((ushort)(slideFlags | 0x0002))
                        : unchecked((ushort)(slideFlags & ~0x0002));
                    slideFlags = FollowsMasterObjects(slide,
                            layoutIsIndependentMaster)
                        ? unchecked((ushort)(slideFlags | 0x0001))
                        : unchecked((ushort)(slideFlags & ~0x0001));
                    WriteUInt16(atom, 28, slideFlags);
                    children.Add(atom);
                    if (headerFooter != null && !hasHeaderFooter
                        && !prototype.Children.Any(candidate =>
                            candidate.Type == RecordSlideShowSlideInfoAtom)) {
                        children.Add(BuildHeaderFooterRecord(headerFooter,
                            instance: 0, allowHeader: false));
                    }
                } else if (child.Type == RecordDrawing) {
                    children.Add(BuildDrawingRecord(prototype, shapes, drawingId,
                        interactionCatalog, animationCatalog, mediaCatalog,
                        oleCatalog, pictureCatalog, fonts, pictureBullets,
                        background));
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
                } else if (classicOverride != null
                           && child.Type == RecordColorSchemeAtom
                           && child.Instance == 1) {
                    children.Add(BuildColorSchemeAtom(classicOverride));
                    wroteClassicOverride = true;
                } else if (child.Type == RecordProgTags) {
                    if (comments.Count > 0) {
                        children.Add(BuildCommentProgrammableTagsRecord(comments));
                        wroteComments = true;
                    }
                } else if (!IsRoundTripThemeRecord(child.Type)) {
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
            if (classicOverride != null && !wroteClassicOverride) {
                children.Add(BuildColorSchemeAtom(classicOverride));
            }
            children.AddRange(roundTripThemeRecords);
            return BuildContainer(RecordSlide, instance: 0, children);
        }
    }
}
