using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Capabilities;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using static OfficeIMO.PowerPoint.LegacyPpt.Write.LegacyPptWriterXml;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWritePreflight {
        internal static LegacyPptWritePreflightReport Analyze(PowerPointDocument presentation) {
            ArgumentNullException.ThrowIfNull(presentation);
            using LegacyPptPackageView view = LegacyPptPackageView.Capture(presentation);
            return Analyze(view, presentation);
        }

        internal static LegacyPptWritePreflightReport Analyze(LegacyPptPackageView view,
            PowerPointDocument presentation) {
            ArgumentNullException.ThrowIfNull(view);
            ArgumentNullException.ThrowIfNull(presentation);
            var findings = new List<LegacyPptWriteFinding>();
            AddPackagePartFindings(view, presentation, findings);
            int packageFindingCount = findings.Count;
            bool canPreserve = presentation.CanPreserveOriginalLegacyPackage
                || LegacyPptPreservingWriter.CanWriteDocument(
                    presentation);
            if (canPreserve) {
                AddPreservedActiveContentFindings(presentation, findings);
            }
            if (canPreserve && packageFindingCount == 0) {
                return new LegacyPptWritePreflightReport(findings);
            }
            if (presentation.LegacyPptPackage != null) {
                findings.Add(new LegacyPptWriteFinding(LegacyPptFeature.UnknownRecordsAndStreams,
                    "ppt-write-import-loss",
                    "The imported binary presentation contains edits that cannot be encoded without losing preserved content."));
            }
            if (presentation.GetSections().Count > 0) {
                findings.Add(new LegacyPptWriteFinding(LegacyPptFeature.Sections, "ppt-write-sections",
                    "Presentation sections are not encoded by the native binary writer."));
            }
            bool canWriteVba = LegacyPptWriter.TryReadVbaProject(presentation,
                out byte[]? vbaProjectBytes, out string? vbaReason);
            if (!canWriteVba) {
                findings.Add(new LegacyPptWriteFinding(LegacyPptFeature.VbaProjects, "ppt-write-vba",
                    vbaReason ?? "The VBA project cannot be encoded by the native binary writer."));
            }
            if (!LegacyPptPropertySetCodec.TryCreateFreshStreams(presentation,
                    out _, out string? propertyReason)) {
                findings.Add(new LegacyPptWriteFinding(
                    LegacyPptFeature.CustomProperties,
                    "ppt-write-document-property",
                    propertyReason
                    ?? "A document property cannot be encoded by the native binary writer."));
            }
            IReadOnlyList<LegacyPptPartView> masterParts = view.SlideMasterParts;
            int masterCount = masterParts.Count;
            LegacyPptWriter.LegacyPptWriterPictureBulletCatalog
                pictureBullets;
            if (!LegacyPptWriter.TryReadPictureBulletCatalog(presentation,
                    out pictureBullets,
                    out string? pictureBulletReason)) {
                findings.Add(new LegacyPptWriteFinding(
                    LegacyPptFeature.BulletsAndNumbering,
                    "ppt-write-picture-bullet",
                    pictureBulletReason
                    ?? "A picture bullet cannot be encoded by the native binary writer."));
                pictureBullets = LegacyPptWriter
                    .LegacyPptWriterPictureBulletCatalog.Empty;
            }
            if (masterCount == 0 || masterCount > LegacyPptWriter.MaxNativeMasterCount) {
                findings.Add(new LegacyPptWriteFinding(LegacyPptFeature.Masters,
                    "ppt-write-master-count",
                    masterCount == 0
                        ? "The presentation has no slide master to encode."
                        : $"The native binary writer currently supports at most {LegacyPptWriter.MaxNativeMasterCount} slide masters; the presentation contains {masterCount}."));
            }
            if (!LegacyPptWriter.CanWriteMasterTextStyles(presentation,
                    out string? masterTextStyleReason)) {
                findings.Add(new LegacyPptWriteFinding(
                    LegacyPptFeature.ParagraphFormatting,
                    "ppt-write-master-text-style",
                    masterTextStyleReason
                    ?? "A slide-master text style cannot be encoded by the native binary writer."));
            }
            LegacyPptWriter.LegacyPptWriterFontCatalog shapeTextFonts =
                LegacyPptWriter.CreateFontCatalogForWrite();
            for (int masterIndex = 0; masterIndex < masterParts.Count;
                 masterIndex++) {
                LegacyPptPartView masterPart = masterParts[masterIndex];
                if (!LegacyPptWriter.TryReadBackground(presentation,
                        masterPart, out _,
                        out string? masterBackgroundReason)) {
                    findings.Add(new LegacyPptWriteFinding(
                        LegacyPptFeature.Backgrounds, "ppt-write-background",
                        masterBackgroundReason
                        ?? "A slide-master background cannot be encoded by the native binary writer."));
                }
                IReadOnlyList<PowerPointShape> masterShapes = LegacyPptWriter
                    .ReadMasterShapesForWrite(presentation, masterPart,
                        out string? masterShapeReason);
                if (masterShapeReason != null) {
                    findings.Add(new LegacyPptWriteFinding(LegacyPptFeature.Masters,
                        "ppt-write-master-shape",
                        $"Slide master {masterIndex}: {masterShapeReason}"));
                }
                AddMasterShapeFindings(findings, masterShapes,
                    $"Slide master {masterIndex}",
                    LegacyPptWriter.LegacyPptWriterShapeContext.MainMaster,
                    shapeTextFonts, pictureBullets);
                if (HasDescendant(masterPart.Root, LegacyPptXml.P, "timing")
                    || HasDescendant(masterPart.Root, LegacyPptXml.P,
                        "transition")) {
                    findings.Add(new LegacyPptWriteFinding(
                        LegacyPptFeature.Animations,
                        "ppt-write-master-timing",
                        $"Slide master {masterIndex} contains transition or timing data that is not encoded by the native binary writer."));
                }
            }
            LegacyPptPartView notesMasterPart = view.NotesMaster;
            if (!notesMasterPart.IsNone
                && !LegacyPptWriter.TryReadBackground(presentation,
                    notesMasterPart, out _,
                    out string? notesBackgroundReason)) {
                findings.Add(new LegacyPptWriteFinding(
                    LegacyPptFeature.Backgrounds, "ppt-write-background",
                    notesBackgroundReason
                    ?? "The notes-master background cannot be encoded by the native binary writer."));
            }
            if (!notesMasterPart.IsNone) {
                IReadOnlyList<PowerPointShape> notesMasterShapes =
                    LegacyPptWriter.ReadMasterShapesForWrite(presentation,
                        notesMasterPart, out string? notesMasterShapeReason);
                if (notesMasterShapeReason != null) {
                    findings.Add(new LegacyPptWriteFinding(
                        LegacyPptFeature.Masters,
                        "ppt-write-notes-master-shape",
                        notesMasterShapeReason));
                }
                AddMasterShapeFindings(findings, notesMasterShapes,
                    "Notes master",
                    LegacyPptWriter.LegacyPptWriterShapeContext.NotesMaster,
                    shapeTextFonts, pictureBullets);
            }
            LegacyPptPartView handoutMasterPart = view.HandoutMaster;
            if (!handoutMasterPart.IsNone
                && !LegacyPptWriter.TryReadBackground(presentation,
                    handoutMasterPart, out _,
                    out string? handoutBackgroundReason)) {
                findings.Add(new LegacyPptWriteFinding(
                    LegacyPptFeature.Backgrounds, "ppt-write-background",
                    handoutBackgroundReason
                    ?? "The handout-master background cannot be encoded by the native binary writer."));
            }
            if (!handoutMasterPart.IsNone) {
                IReadOnlyList<PowerPointShape> handoutMasterShapes =
                    LegacyPptWriter.ReadMasterShapesForWrite(presentation,
                        handoutMasterPart, out string? handoutMasterShapeReason);
                if (handoutMasterShapeReason != null) {
                    findings.Add(new LegacyPptWriteFinding(
                        LegacyPptFeature.Masters,
                        "ppt-write-handout-master-shape",
                        handoutMasterShapeReason));
                }
                AddMasterShapeFindings(findings, handoutMasterShapes,
                    "Handout master",
                    LegacyPptWriter.LegacyPptWriterShapeContext.HandoutMaster,
                    shapeTextFonts, pictureBullets);
            }
            LegacyPptWriter.LegacyPptWriterTopology? topology = null;
            if (masterCount > 0
                && masterCount <= LegacyPptWriter.MaxNativeMasterCount) {
                int notesCount = presentation.Slides.Count(slide =>
                    LegacyPptWriter.ShouldWriteNotesPage(slide, out _));
                try {
                    topology = new LegacyPptWriter.LegacyPptWriterTopology(
                        masterCount, presentation.Slides.Count, notesCount,
                        !handoutMasterPart.IsNone);
                } catch (NotSupportedException exception) {
                    findings.Add(new LegacyPptWriteFinding(
                        LegacyPptFeature.UnknownRecordsAndStreams,
                        "ppt-write-persist-capacity", exception.Message));
                }
            }
            if (LegacyPptWriter.HasModernComments(presentation)) {
                findings.Add(new LegacyPptWriteFinding(LegacyPptFeature.ModernComments,
                    "ppt-write-modern-comments",
                    "Modern threaded comments, replies, status, and shape anchors have no native PowerPoint 97-2003 representation."));
            }
            if (!LegacyPptWriter.TryReadAllClassicComments(presentation, out _,
                    out string? commentReason)) {
                findings.Add(new LegacyPptWriteFinding(LegacyPptFeature.Comments,
                    "ppt-write-comments",
                    commentReason ?? "Classic comments cannot be encoded by the native binary writer."));
            }
            if (!LegacyPptWriter.TryReadCustomShows(presentation, out _,
                    out string? customShowReason)) {
                findings.Add(new LegacyPptWriteFinding(LegacyPptFeature.CustomShows,
                    "ppt-write-custom-show",
                    customShowReason
                    ?? "A custom show cannot be encoded by the native binary writer."));
            }
            var externalObjectSounds = new LegacyPptWriter
                .LegacyPptWriterSoundCatalog();
            uint firstMediaObjectId = 1;
            if (!LegacyPptWriter.TryReadInteractions(view, presentation.Slides,
                    externalObjectSounds,
                    out LegacyPptWriter.LegacyPptWriterInteractionCatalog
                        interactionCatalog,
                    out string? interactionReason)) {
                findings.Add(new LegacyPptWriteFinding(LegacyPptFeature.Hyperlinks,
                    "ppt-write-interaction",
                    interactionReason ?? "A hyperlink or action cannot be encoded by the native binary writer."));
            } else {
                firstMediaObjectId = checked((uint)
                    interactionCatalog.Hyperlinks.Count + 1U);
            }
            uint firstOleObjectId = firstMediaObjectId;
            if (!LegacyPptWriter.TryReadMedia(view, presentation.Slides,
                    firstMediaObjectId, externalObjectSounds,
                    out LegacyPptWriter.LegacyPptWriterMediaCatalog
                        mediaCatalog,
                    out string? mediaReason)) {
                findings.Add(new LegacyPptWriteFinding(
                    presentation.Slides.SelectMany(slide => slide.Media)
                        .Any(media => media.Kind == PowerPointMediaKind.Video)
                        ? LegacyPptFeature.EmbeddedVideo
                        : LegacyPptFeature.Media,
                    "ppt-write-media",
                    mediaReason
                    ?? "Embedded media cannot be encoded by the native binary writer."));
            } else {
                firstOleObjectId = mediaCatalog.NextObjectId;
            }
            if (!LegacyPptWriter.TryReadOleObjects(view, presentation.Slides,
                    firstOleObjectId,
                    out LegacyPptWriter.LegacyPptWriterOleObjectCatalog
                        oleCatalog,
                    out string? oleReason)) {
                findings.Add(new LegacyPptWriteFinding(
                    LegacyPptFeature.EmbeddedOle, "ppt-write-ole",
                    oleReason
                    ?? "An embedded OLE object cannot be encoded by the native binary writer."));
            } else if (topology != null) {
                try {
                    topology.EnsurePersistObjectCapacity(checked(
                        oleCatalog.Objects.Count
                        + (canWriteVba && vbaProjectBytes != null ? 1 : 0)));
                } catch (NotSupportedException exception) {
                    findings.Add(new LegacyPptWriteFinding(
                        LegacyPptFeature.UnknownRecordsAndStreams,
                        "ppt-write-persist-capacity", exception.Message));
                }
            }
            if (!LegacyPptWriter.TryReadPictureCatalog(view, presentation,
                    shapeTextFonts, pictureBullets,
                    convertUnsupportedTables: true,
                    out _, out LegacyPptFeature pictureFailureFeature,
                    out string? pictureReason)) {
                findings.Add(new LegacyPptWriteFinding(
                    pictureFailureFeature,
                    pictureFailureFeature == LegacyPptFeature.Charts
                        ? "ppt-write-chart"
                        : pictureFailureFeature == LegacyPptFeature.Tables
                            ? "ppt-write-table"
                        : pictureFailureFeature == LegacyPptFeature.SmartArt
                            ? "ppt-write-smartart"
                        : pictureFailureFeature == LegacyPptFeature.Backgrounds
                            ? "ppt-write-background"
                        : pictureFailureFeature == LegacyPptFeature.Layouts
                            ? "ppt-write-layout-picture"
                        : "ppt-write-picture",
                    pictureReason
                    ?? "A visual cannot be encoded by the native binary writer."));
            }
            for (int slideIndex = 0; slideIndex < presentation.Slides.Count; slideIndex++) {
                PowerPointSlide slide = presentation.Slides[slideIndex];
                IReadOnlyList<PowerPointShape> shapes = LegacyPptWriter
                    .ReadSlideShapesForWrite(slide, out string? layoutShapeReason);
                if (layoutShapeReason != null) {
                    findings.Add(new LegacyPptWriteFinding(LegacyPptFeature.Layouts,
                        "ppt-write-layout-shape", layoutShapeReason, slideIndex));
                }
                shapes = LegacyPptWriter.FlattenShapeTreeForWrite(shapes,
                    out string? groupShapeReason);
                if (groupShapeReason != null) {
                    findings.Add(new LegacyPptWriteFinding(
                        LegacyPptFeature.Groups,
                        "ppt-write-group-shape", groupShapeReason,
                        slideIndex));
                }
                if (HasUnsupportedRichNotes(view, slide)) {
                    findings.Add(new LegacyPptWriteFinding(LegacyPptFeature.RichNotes,
                        "ppt-write-rich-notes",
                        "Notes-page drawings beyond the speaker-notes body are not encoded by the native binary writer.",
                        slideIndex));
                }
                OfficeXmlNode slideRoot = view.Part(slide.PartRef.Name).Root;
                if (!LegacyPptWriter.TryReadTransition(view, slide, out _,
                        out string? transitionReason)) {
                    findings.Add(new LegacyPptWriteFinding(LegacyPptFeature.Transitions, "ppt-write-transition",
                        transitionReason ?? "The slide transition cannot be encoded by the native binary writer.",
                        slideIndex));
                }
                if (!slideRoot.Child(LegacyPptXml.P, "timing").IsNone) {
                    var animationSounds = new LegacyPptWriter.LegacyPptWriterSoundCatalog();
                    if (!LegacyPptWriter.TryReadClassicAnimations(view, new[] { slide },
                            animationSounds, out _, out string? animationReason)) {
                        findings.Add(new LegacyPptWriteFinding(LegacyPptFeature.Animations,
                            "ppt-write-timing",
                            animationReason ?? "The timing tree cannot be encoded as classic binary PowerPoint animations.",
                            slideIndex));
                    }
                }
                if (!LegacyPptWriter.TryReadBackground(slide, out _,
                        out string? backgroundReason)) {
                    findings.Add(new LegacyPptWriteFinding(LegacyPptFeature.Backgrounds, "ppt-write-background",
                        backgroundReason
                        ?? "The slide background cannot be encoded by the native binary writer.",
                        slideIndex));
                }
                for (int shapeIndex = 0; shapeIndex < shapes.Count; shapeIndex++) {
                    PowerPointShape shape = shapes[shapeIndex];
                    if (!LegacyPptWriter.TryReadPlaceholderForWrite(shape,
                            LegacyPptWriter.LegacyPptWriterShapeContext.Slide,
                            out _, out string? placeholderReason)) {
                        findings.Add(new LegacyPptWriteFinding(
                            LegacyPptFeature.Placeholders,
                            "ppt-write-placeholder",
                            placeholderReason
                            ?? "A placeholder cannot be encoded by the native binary writer.",
                            slideIndex, shapeIndex));
                    }
                    if (!IsSupportedShape(shape,
                            includeOleObjects: true,
                            includeMedia: true,
                            includePictures: true,
                            includeCharts: true,
                            includeSmartArt: true)) {
                        findings.Add(new LegacyPptWriteFinding(MapShapeFeature(shape), "ppt-write-shape",
                            $"{shape.ShapeContentType} content is outside the native writer's supported shape subset.",
                            slideIndex, shapeIndex));
                        continue;
                    }
                    if (shape is PowerPointTable table
                        && !LegacyPptWriter.TryReadTableForWrite(table,
                            shapeTextFonts, pictureBullets,
                            out string? tableReason)) {
                        findings.Add(new LegacyPptWriteFinding(
                            LegacyPptFeature.Tables,
                            "ppt-write-table",
                            tableReason
                            ?? "The DrawingML table cannot be encoded as a native binary PowerPoint table.",
                            slideIndex, shapeIndex));
                    } else if (shape is PowerPointChart) {
                        findings.Add(new LegacyPptWriteFinding(
                            LegacyPptFeature.Charts,
                            "ppt-write-chart-converted",
                            "The editable DrawingML chart will be converted to a static PNG picture in binary PowerPoint. Chart data and editability will not survive the conversion.",
                            slideIndex, shapeIndex));
                    } else if (shape is PowerPointSmartArt) {
                        findings.Add(new LegacyPptWriteFinding(
                            LegacyPptFeature.SmartArt,
                            "ppt-write-smartart-converted",
                            "The editable SmartArt diagram will be converted to a static PNG picture in binary PowerPoint. Diagram structure and editability will not survive the conversion.",
                            slideIndex, shapeIndex));
                    }
                    if (HasUnsupportedVisualStyle(shape)) {
                        findings.Add(new LegacyPptWriteFinding(LegacyPptFeature.ShapeStyles, "ppt-write-shape-style",
                            "Fill, outline, transform, effects, hyperlink, or visibility styling is not encoded.",
                            slideIndex, shapeIndex));
                    }
                    if (!LegacyPptWriter.TryReadShapeMetadataForWrite(shape,
                            out _, out string? metadataReason)) {
                        findings.Add(new LegacyPptWriteFinding(
                            LegacyPptFeature.AccessibilityMetadata,
                            "ppt-write-accessibility-metadata",
                            metadataReason
                            ?? "The shape accessibility metadata has no classic binary representation.",
                            slideIndex, shapeIndex));
                    }
                    if (LegacyPptWriter.IsLayoutShape(shape)
                        && HasUnsupportedMasterInteraction(shape)) {
                        findings.Add(new LegacyPptWriteFinding(
                            LegacyPptFeature.Layouts,
                            "ppt-write-layout-interaction",
                            "Interactions on a materialized ordinary-layout shape are not encoded by the native binary writer.",
                            slideIndex, shapeIndex));
                    }
                    if (shape is PowerPointTextBox textBox
                        && !LegacyPptWriter.TryReadTextBoxForWrite(textBox,
                            shapeTextFonts, pictureBullets,
                            out string? textReason)) {
                        findings.Add(new LegacyPptWriteFinding(LegacyPptFeature.RichText, "ppt-write-rich-text",
                            textReason
                            ?? "The text formatting cannot be encoded by the native binary writer.",
                            slideIndex, shapeIndex));
                    }
                }
            }

            foreach (var diagnostic in presentation.LegacyPptImportDiagnostics) {
                if (diagnostic.Severity == Diagnostics.LegacyPptDiagnosticSeverity.Warning
                    || diagnostic.Severity == Diagnostics.LegacyPptDiagnosticSeverity.Error) {
                    findings.Add(new LegacyPptWriteFinding(LegacyPptFeature.UnknownRecordsAndStreams,
                        "ppt-write-import-loss",
                        $"Imported legacy content was not fully projected: {diagnostic.Code}."));
                }
            }
            return new LegacyPptWritePreflightReport(findings);
        }

        private static void AddPreservedActiveContentFindings(
            PowerPointDocument presentation,
            ICollection<LegacyPptWriteFinding> findings) {
            if (presentation.LegacyPptWillPreserveVbaContent) {
                findings.Add(new LegacyPptWriteFinding(
                    LegacyPptFeature.VbaProjects,
                    "ppt-write-preserved-vba",
                    "The edited binary presentation will preserve an imported VBA project. Set LossPolicy to Allow only when retaining that active content is intentional."));
            }
            if (presentation.LegacyPptHasEmbeddedOleContent) {
                findings.Add(new LegacyPptWriteFinding(
                    LegacyPptFeature.EmbeddedOle,
                    "ppt-write-preserved-ole",
                    "The edited binary presentation will preserve imported embedded OLE content. Set LossPolicy to Allow only when retaining that active content is intentional."));
            }
            if (presentation.LegacyPptHasLinkedOleContent) {
                findings.Add(new LegacyPptWriteFinding(
                    LegacyPptFeature.LinkedOle,
                    "ppt-write-preserved-linked-ole",
                    "The edited binary presentation will preserve imported linked OLE content. Set LossPolicy to Allow only when retaining that external content is intentional."));
            }
            if (presentation.LegacyPptHasActiveXContent) {
                findings.Add(new LegacyPptWriteFinding(
                    LegacyPptFeature.ActiveX,
                    "ppt-write-preserved-activex",
                    "The edited binary presentation will preserve imported ActiveX content. Set LossPolicy to Allow only when retaining that active content is intentional."));
            }
            if (presentation.LegacyPptWillPreserveRunProgramContent) {
                findings.Add(new LegacyPptWriteFinding(
                    LegacyPptFeature.Actions,
                    "ppt-write-preserved-run-program",
                    "The edited binary presentation will preserve a run-program action. Set LossPolicy to Allow only when retaining that active action is intentional."));
            }
            if (presentation.LegacyPptWillPreserveExternalHyperlinkContent) {
                findings.Add(new LegacyPptWriteFinding(
                    LegacyPptFeature.Hyperlinks,
                    "ppt-write-preserved-external-link",
                    "The edited binary presentation will preserve an external hyperlink. Set LossPolicy to Allow only when retaining that external target is intentional."));
            }
            if (presentation.LegacyPptHasExternalMediaContent) {
                findings.Add(new LegacyPptWriteFinding(
                    LegacyPptFeature.LinkedMedia,
                    "ppt-write-preserved-external-media",
                    "The edited binary presentation will preserve externally linked media. Set LossPolicy to Allow only when retaining that external target is intentional."));
            }
        }

        internal static bool CanWriteSlideLosslessly(PowerPointSlide slide) {
            ArgumentNullException.ThrowIfNull(slide);
            if (LegacyPptWriter.ShouldWriteNotesPage(slide, out _)) return false;
            using LegacyPptPackageView view = LegacyPptPackageView.Capture(
                slide.OwnerDocument ?? throw new InvalidOperationException(
                    "The slide does not belong to a presentation."));
            OfficeXmlNode slideRoot = view.Part(slide.PartRef.Name).Root;
            bool animationsSupported = slideRoot.Child(LegacyPptXml.P, "timing").IsNone
                || LegacyPptWriter.TryReadClassicAnimations(view, new[] { slide },
                    new LegacyPptWriter.LegacyPptWriterSoundCatalog(),
                    out _, out _);
            if (!LegacyPptWriter.TryReadTransition(view, slide, out _, out _)
                || !animationsSupported
                || !LegacyPptWriter.TryReadBackground(slide, out _, out _)) {
                return false;
            }
            if (!LegacyPptWriter.TryReadInteractions(view, new[] { slide },
                    new LegacyPptWriter.LegacyPptWriterSoundCatalog(),
                    out _, out _)) return false;
            IReadOnlyList<PowerPointShape> shapes = LegacyPptWriter
                .ReadSlideShapesForWrite(slide, out string? layoutShapeReason);
            if (layoutShapeReason != null) return false;
            shapes = LegacyPptWriter.FlattenShapeTreeForWrite(shapes,
                out string? groupShapeReason);
            if (groupShapeReason != null) return false;
            LegacyPptWriter.LegacyPptWriterFontCatalog fonts =
                LegacyPptWriter.CreateFontCatalogForWrite();
            if (slide.OwnerDocument == null
                || !LegacyPptWriter.TryReadPictureBulletCatalog(view,
                    slide.OwnerDocument, out LegacyPptWriter
                        .LegacyPptWriterPictureBulletCatalog pictureBullets,
                    out _)) return false;
            foreach (PowerPointShape shape in shapes) {
                if (!IsSupportedShape(shape) || HasUnsupportedVisualStyle(shape)
                    || !LegacyPptWriter.TryReadShapeMetadataForWrite(shape,
                        out _, out _)
                    || LegacyPptWriter.IsLayoutShape(shape)
                    && HasUnsupportedMasterInteraction(shape)
                    || !LegacyPptWriter.TryReadPlaceholderForWrite(shape,
                        LegacyPptWriter.LegacyPptWriterShapeContext.Slide,
                        out _, out _)) return false;
                if (shape is PowerPointTextBox textBox
                    && !LegacyPptWriter.TryReadTextBoxForWrite(textBox,
                        fonts, pictureBullets, out _)) return false;
                if (shape is PowerPointTable table
                    && !LegacyPptWriter.TryReadTableForWrite(table,
                        fonts, pictureBullets, out _)) return false;
            }
            return true;
        }

        private static bool IsSupportedShape(PowerPointShape shape,
            bool includeOleObjects = false,
            bool includeMedia = false,
            bool includePictures = false,
            bool includeCharts = false,
            bool includeSmartArt = false) => shape is PowerPointTextBox
            || shape is PowerPointTable
            || includeMedia && shape is PowerPointMedia
            || includePictures && shape is PowerPointImage
            || includeCharts && shape is PowerPointChart
            || includeSmartArt && shape is PowerPointSmartArt
            || includeOleObjects && shape is PowerPointOleObject
            || shape is PowerPointGroupShape group
            && LegacyPptWriter.TryReadGroupForWrite(group,
                out _, out _)
            || shape is PowerPointConnectionShape connector
            && LegacyPptWriter.TryReadOfficeArtShapeType(connector,
                requireConnector: true, out _, out _)
            || shape is PowerPointAutoShape autoShape
            && LegacyPptWriter.TryReadOfficeArtShapeType(autoShape,
                requireConnector: false, out _, out _);

        private static bool HasUnsupportedRichNotes(LegacyPptPackageView view,
            PowerPointSlide slide) {
            OfficeXmlNode tree = view.Part(slide.PartRef.Name).NotesSlide.Root
                .Child(LegacyPptXml.P, "cSld").Child(LegacyPptXml.P, "spTree");
            if (tree.IsNone) return false;
            int bodyPlaceholderCount = 0;
            foreach (OfficeXmlNode child in tree.Children()) {
                if (child.Is(LegacyPptXml.P, "nvGrpSpPr")
                    || child.Is(LegacyPptXml.P, "grpSpPr")) continue;
                if (!child.Is(LegacyPptXml.P, "sp")) return true;
                OfficeXmlNode placeholder = child.Child(LegacyPptXml.P, "nvSpPr")
                    .Child(LegacyPptXml.P, "nvPr").Child(LegacyPptXml.P, "ph");
                if (placeholder.IsNone || placeholder.Attribute("type") != "body") return true;
                bodyPlaceholderCount++;
            }
            return bodyPlaceholderCount > 1;
        }

        private static void AddMasterShapeFindings(
            ICollection<LegacyPptWriteFinding> findings,
            IReadOnlyList<PowerPointShape> shapes, string ownerName,
            LegacyPptWriter.LegacyPptWriterShapeContext shapeContext,
            LegacyPptWriter.LegacyPptWriterFontCatalog fonts,
            LegacyPptWriter.LegacyPptWriterPictureBulletCatalog
                pictureBullets) {
            shapes = LegacyPptWriter.FlattenShapeTreeForWrite(shapes,
                out string? groupReason);
            if (groupReason != null) {
                findings.Add(new LegacyPptWriteFinding(
                    LegacyPptFeature.Groups,
                    "ppt-write-master-group", $"{ownerName}: {groupReason}"));
            }
            for (int shapeIndex = 0; shapeIndex < shapes.Count; shapeIndex++) {
                PowerPointShape shape = shapes[shapeIndex];
                string location = $"{ownerName}, shape {shapeIndex}";
                if (!LegacyPptWriter.TryReadPlaceholderForWrite(shape,
                        shapeContext, out _, out string? placeholderReason)) {
                    findings.Add(new LegacyPptWriteFinding(
                        LegacyPptFeature.Placeholders,
                        "ppt-write-master-placeholder",
                        $"{location}: {placeholderReason}"));
                }
                if (!LegacyPptWriter.IsSupportedMasterShape(shape)) {
                    findings.Add(new LegacyPptWriteFinding(
                        MapShapeFeature(shape), "ppt-write-master-shape",
                        $"{location}: {shape.ShapeContentType} content is outside the native writer's supported master-shape subset."));
                    continue;
                }
                if (HasUnsupportedVisualStyle(shape)
                    || HasUnsupportedMasterInteraction(shape)) {
                    findings.Add(new LegacyPptWriteFinding(
                        LegacyPptFeature.ShapeStyles,
                        "ppt-write-master-shape-style",
                        $"{location}: visual styling or interactive content is not encoded on binary masters."));
                }
                if (!LegacyPptWriter.TryReadShapeMetadataForWrite(shape,
                        out _, out string? metadataReason)) {
                    findings.Add(new LegacyPptWriteFinding(
                        LegacyPptFeature.AccessibilityMetadata,
                        "ppt-write-master-accessibility-metadata",
                        $"{location}: {metadataReason}"));
                }
                if (shape is PowerPointTextBox textBox
                    && !LegacyPptWriter.TryReadTextBoxForWrite(textBox,
                        fonts, pictureBullets,
                        out string? textReason)) {
                    findings.Add(new LegacyPptWriteFinding(
                        LegacyPptFeature.RichText,
                        "ppt-write-master-rich-text",
                        $"{location}: {textReason}"));
                } else if (shape is PowerPointTable table
                    && !LegacyPptWriter.TryReadTableForWrite(table, fonts,
                        pictureBullets, out string? tableReason)) {
                    findings.Add(new LegacyPptWriteFinding(
                        LegacyPptFeature.Tables,
                        "ppt-write-master-table",
                        $"{location}: {tableReason ?? "The DrawingML table cannot be encoded as a native binary PowerPoint table."}"));
                }
            }
        }

        private static LegacyPptFeature MapShapeFeature(PowerPointShape shape) {
            switch (shape.ShapeContentType) {
                case PowerPointShapeContentType.Picture: return LegacyPptFeature.RasterPictures;
                case PowerPointShapeContentType.Table: return LegacyPptFeature.Tables;
                case PowerPointShapeContentType.Chart: return LegacyPptFeature.Charts;
                case PowerPointShapeContentType.Group: return LegacyPptFeature.Groups;
                case PowerPointShapeContentType.Media:
                    return shape is PowerPointMedia {
                        Kind: PowerPointMediaKind.Video
                    }
                        ? LegacyPptFeature.EmbeddedVideo
                        : LegacyPptFeature.Media;
                case PowerPointShapeContentType.SmartArt: return LegacyPptFeature.SmartArt;
                case PowerPointShapeContentType.OleObject: return LegacyPptFeature.EmbeddedOle;
                case PowerPointShapeContentType.Connector: return LegacyPptFeature.Connectors;
                case PowerPointShapeContentType.AutoShape: return LegacyPptFeature.AutoShapes;
                case PowerPointShapeContentType.TextBox: return LegacyPptFeature.RichText;
                default: return LegacyPptFeature.UnknownRecordsAndStreams;
            }
        }

        private static bool HasUnsupportedVisualStyle(PowerPointShape shape) =>
            !LegacyPptWriter.TryReadShapeTransform(shape, out _, out _)
            || !LegacyPptWriter.TryReadShapeVisualStyle(shape, out _, out _)
            || shape is PowerPointTextBox textBox
                && !LegacyPptWriter.TryReadTextFrameForWrite(textBox,
                    out _, out _)
            || HasDescendant(shape.XmlElement, LegacyPptXml.A, "effectDag");

        private static bool HasUnsupportedMasterInteraction(
            PowerPointShape shape) =>
            HasDescendant(shape.XmlElement, LegacyPptXml.A, "hlinkClick")
            || HasDescendant(shape.XmlElement, LegacyPptXml.A, "hlinkHover")
            || HasDescendant(shape.XmlElement, LegacyPptXml.A, "hlinkMouseOver");

        internal static void ThrowIfBlocked(LegacyPptWritePreflightReport report, PowerPointSaveOptions? options) {
            if (!report.HasConversionLoss || options?.LossPolicy == OfficeConversionLossPolicy.Allow) return;
            string details = string.Join("; ", report.Findings.Take(8));
            throw new NotSupportedException(
                "Native PPT/POT/PPS saving is blocked because known content cannot be encoded without loss. "
                + details + " Set PowerPointSaveOptions.LossPolicy to Allow only when that loss is intentional.");
        }
    }
}
