using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Capabilities;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWritePreflight {
        internal static LegacyPptWritePreflightReport Analyze(PowerPointPresentation presentation) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            var findings = new List<LegacyPptWriteFinding>();
            AddPackagePartFindings(presentation, findings);
            if (findings.Count == 0
                && (presentation.CanPreserveOriginalLegacyPackage
                    || LegacyPptPreservingWriter.CanWritePresentation(presentation))) {
                return new LegacyPptWritePreflightReport(findings);
            }
            if (presentation.LegacyPptPackage != null) {
                findings.Add(new LegacyPptWriteFinding(LegacyPptFeature.UnknownRecordsAndStreams,
                    "PPT-WRITE-IMPORT-LOSS",
                    "The imported binary presentation contains edits that cannot be encoded without losing preserved content."));
            }
            if (presentation.GetSections().Count > 0) {
                findings.Add(new LegacyPptWriteFinding(LegacyPptFeature.Sections, "PPT-WRITE-SECTIONS",
                    "Presentation sections are not encoded by the native binary writer."));
            }
            bool canWriteVba = LegacyPptWriter.TryReadVbaProject(presentation,
                out byte[]? vbaProjectBytes, out string? vbaReason);
            if (!canWriteVba) {
                findings.Add(new LegacyPptWriteFinding(LegacyPptFeature.VbaProjects, "PPT-WRITE-VBA",
                    vbaReason ?? "The VBA project cannot be encoded by the native binary writer."));
            }
            if (!LegacyPptPropertySetCodec.TryCreateFreshStreams(presentation,
                    out _, out string? propertyReason)) {
                findings.Add(new LegacyPptWriteFinding(
                    LegacyPptFeature.CustomProperties,
                    "PPT-WRITE-DOCUMENT-PROPERTY",
                    propertyReason
                    ?? "A document property cannot be encoded by the native binary writer."));
            }
            SlideMasterPart[] masterParts = presentation.OpenXmlDocument
                .PresentationPart?.SlideMasterParts.ToArray()
                ?? Array.Empty<SlideMasterPart>();
            int masterCount = masterParts.Length;
            LegacyPptWriter.LegacyPptWriterPictureBulletCatalog
                pictureBullets;
            if (!LegacyPptWriter.TryReadPictureBulletCatalog(presentation,
                    out pictureBullets,
                    out string? pictureBulletReason)) {
                findings.Add(new LegacyPptWriteFinding(
                    LegacyPptFeature.BulletsAndNumbering,
                    "PPT-WRITE-PICTURE-BULLET",
                    pictureBulletReason
                    ?? "A picture bullet cannot be encoded by the native binary writer."));
                pictureBullets = LegacyPptWriter
                    .LegacyPptWriterPictureBulletCatalog.Empty;
            }
            if (masterCount == 0 || masterCount > LegacyPptWriter.MaxNativeMasterCount) {
                findings.Add(new LegacyPptWriteFinding(LegacyPptFeature.Masters,
                    "PPT-WRITE-MASTER-COUNT",
                    masterCount == 0
                        ? "The presentation has no slide master to encode."
                        : $"The native binary writer currently supports at most {LegacyPptWriter.MaxNativeMasterCount} slide masters; the presentation contains {masterCount}."));
            }
            if (!LegacyPptWriter.CanWriteMasterTextStyles(presentation,
                    out string? masterTextStyleReason)) {
                findings.Add(new LegacyPptWriteFinding(
                    LegacyPptFeature.ParagraphFormatting,
                    "PPT-WRITE-MASTER-TEXT-STYLE",
                    masterTextStyleReason
                    ?? "A slide-master text style cannot be encoded by the native binary writer."));
            }
            LegacyPptWriter.LegacyPptWriterFontCatalog shapeTextFonts =
                LegacyPptWriter.CreateFontCatalogForWrite();
            for (int masterIndex = 0; masterIndex < masterParts.Length;
                 masterIndex++) {
                SlideMasterPart masterPart = masterParts[masterIndex];
                if (!LegacyPptWriter.TryReadBackground(masterPart, out _,
                        out string? masterBackgroundReason)) {
                    findings.Add(new LegacyPptWriteFinding(
                        LegacyPptFeature.Backgrounds, "PPT-WRITE-BACKGROUND",
                        masterBackgroundReason
                        ?? "A slide-master background cannot be encoded by the native binary writer."));
                }
                IReadOnlyList<PowerPointShape> masterShapes = LegacyPptWriter
                    .ReadMasterShapesForWrite(masterPart,
                        out string? masterShapeReason);
                if (masterShapeReason != null) {
                    findings.Add(new LegacyPptWriteFinding(LegacyPptFeature.Masters,
                        "PPT-WRITE-MASTER-SHAPE",
                        $"Slide master {masterIndex}: {masterShapeReason}"));
                }
                AddMasterShapeFindings(findings, masterShapes,
                    $"Slide master {masterIndex}",
                    LegacyPptWriter.LegacyPptWriterShapeContext.MainMaster,
                    shapeTextFonts, pictureBullets);
                if (masterPart.SlideMaster?.Descendants<P.Timing>().Any()
                        == true
                    || masterPart.SlideMaster?.Descendants<P.Transition>().Any()
                        == true) {
                    findings.Add(new LegacyPptWriteFinding(
                        LegacyPptFeature.Animations,
                        "PPT-WRITE-MASTER-TIMING",
                        $"Slide master {masterIndex} contains transition or timing data that is not encoded by the native binary writer."));
                }
            }
            NotesMasterPart? notesMasterPart = presentation.OpenXmlDocument
                .PresentationPart?.NotesMasterPart;
            if (notesMasterPart != null
                && !LegacyPptWriter.TryReadBackground(notesMasterPart, out _,
                    out string? notesBackgroundReason)) {
                findings.Add(new LegacyPptWriteFinding(
                    LegacyPptFeature.Backgrounds, "PPT-WRITE-BACKGROUND",
                    notesBackgroundReason
                    ?? "The notes-master background cannot be encoded by the native binary writer."));
            }
            if (notesMasterPart != null) {
                IReadOnlyList<PowerPointShape> notesMasterShapes =
                    LegacyPptWriter.ReadMasterShapesForWrite(notesMasterPart,
                        out string? notesMasterShapeReason);
                if (notesMasterShapeReason != null) {
                    findings.Add(new LegacyPptWriteFinding(
                        LegacyPptFeature.Masters,
                        "PPT-WRITE-NOTES-MASTER-SHAPE",
                        notesMasterShapeReason));
                }
                AddMasterShapeFindings(findings, notesMasterShapes,
                    "Notes master",
                    LegacyPptWriter.LegacyPptWriterShapeContext.NotesMaster,
                    shapeTextFonts, pictureBullets);
            }
            HandoutMasterPart? handoutMasterPart = presentation.OpenXmlDocument
                .PresentationPart?.HandoutMasterPart;
            if (handoutMasterPart != null
                && !LegacyPptWriter.TryReadBackground(handoutMasterPart, out _,
                    out string? handoutBackgroundReason)) {
                findings.Add(new LegacyPptWriteFinding(
                    LegacyPptFeature.Backgrounds, "PPT-WRITE-BACKGROUND",
                    handoutBackgroundReason
                    ?? "The handout-master background cannot be encoded by the native binary writer."));
            }
            if (handoutMasterPart != null) {
                IReadOnlyList<PowerPointShape> handoutMasterShapes =
                    LegacyPptWriter.ReadMasterShapesForWrite(
                        handoutMasterPart, out string? handoutMasterShapeReason);
                if (handoutMasterShapeReason != null) {
                    findings.Add(new LegacyPptWriteFinding(
                        LegacyPptFeature.Masters,
                        "PPT-WRITE-HANDOUT-MASTER-SHAPE",
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
                        handoutMasterPart != null);
                } catch (NotSupportedException exception) {
                    findings.Add(new LegacyPptWriteFinding(
                        LegacyPptFeature.UnknownRecordsAndStreams,
                        "PPT-WRITE-PERSIST-CAPACITY", exception.Message));
                }
            }
            if (LegacyPptWriter.HasModernComments(presentation)) {
                findings.Add(new LegacyPptWriteFinding(LegacyPptFeature.ModernComments,
                    "PPT-WRITE-MODERN-COMMENTS",
                    "Modern threaded comments, replies, status, and shape anchors have no native PowerPoint 97-2003 representation."));
            }
            if (!LegacyPptWriter.TryReadAllClassicComments(presentation, out _,
                    out string? commentReason)) {
                findings.Add(new LegacyPptWriteFinding(LegacyPptFeature.Comments,
                    "PPT-WRITE-COMMENTS",
                    commentReason ?? "Classic comments cannot be encoded by the native binary writer."));
            }
            if (!LegacyPptWriter.TryReadCustomShows(presentation, out _,
                    out string? customShowReason)) {
                findings.Add(new LegacyPptWriteFinding(LegacyPptFeature.CustomShows,
                    "PPT-WRITE-CUSTOM-SHOW",
                    customShowReason
                    ?? "A custom show cannot be encoded by the native binary writer."));
            }
            var externalObjectSounds = new LegacyPptWriter
                .LegacyPptWriterSoundCatalog();
            uint firstMediaObjectId = 1;
            if (!LegacyPptWriter.TryReadInteractions(presentation.Slides,
                    externalObjectSounds,
                    out LegacyPptWriter.LegacyPptWriterInteractionCatalog
                        interactionCatalog,
                    out string? interactionReason)) {
                findings.Add(new LegacyPptWriteFinding(LegacyPptFeature.Hyperlinks,
                    "PPT-WRITE-INTERACTION",
                    interactionReason ?? "A hyperlink or action cannot be encoded by the native binary writer."));
            } else {
                firstMediaObjectId = checked((uint)
                    interactionCatalog.Hyperlinks.Count + 1U);
            }
            uint firstOleObjectId = firstMediaObjectId;
            if (!LegacyPptWriter.TryReadMedia(presentation.Slides,
                    firstMediaObjectId, externalObjectSounds,
                    out LegacyPptWriter.LegacyPptWriterMediaCatalog
                        mediaCatalog,
                    out string? mediaReason)) {
                findings.Add(new LegacyPptWriteFinding(
                    presentation.Slides.SelectMany(slide => slide.Media)
                        .Any(media => media.Kind == PowerPointMediaKind.Video)
                        ? LegacyPptFeature.EmbeddedVideo
                        : LegacyPptFeature.Media,
                    "PPT-WRITE-MEDIA",
                    mediaReason
                    ?? "Embedded media cannot be encoded by the native binary writer."));
            } else {
                firstOleObjectId = mediaCatalog.NextObjectId;
            }
            if (!LegacyPptWriter.TryReadOleObjects(presentation.Slides,
                    firstOleObjectId,
                    out LegacyPptWriter.LegacyPptWriterOleObjectCatalog
                        oleCatalog,
                    out string? oleReason)) {
                findings.Add(new LegacyPptWriteFinding(
                    LegacyPptFeature.EmbeddedOle, "PPT-WRITE-OLE",
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
                        "PPT-WRITE-PERSIST-CAPACITY", exception.Message));
                }
            }
            if (!LegacyPptWriter.TryReadPictureCatalog(presentation,
                    shapeTextFonts, pictureBullets,
                    convertUnsupportedTables: true,
                    out _, out LegacyPptFeature pictureFailureFeature,
                    out string? pictureReason)) {
                findings.Add(new LegacyPptWriteFinding(
                    pictureFailureFeature,
                    pictureFailureFeature == LegacyPptFeature.Charts
                        ? "PPT-WRITE-CHART"
                        : pictureFailureFeature == LegacyPptFeature.Tables
                            ? "PPT-WRITE-TABLE"
                        : pictureFailureFeature == LegacyPptFeature.SmartArt
                            ? "PPT-WRITE-SMARTART"
                        : pictureFailureFeature == LegacyPptFeature.Backgrounds
                            ? "PPT-WRITE-BACKGROUND"
                        : pictureFailureFeature == LegacyPptFeature.Layouts
                            ? "PPT-WRITE-LAYOUT-PICTURE"
                        : "PPT-WRITE-PICTURE",
                    pictureReason
                    ?? "A visual cannot be encoded by the native binary writer."));
            }
            for (int slideIndex = 0; slideIndex < presentation.Slides.Count; slideIndex++) {
                PowerPointSlide slide = presentation.Slides[slideIndex];
                IReadOnlyList<PowerPointShape> shapes = LegacyPptWriter
                    .ReadSlideShapesForWrite(slide, out string? layoutShapeReason);
                if (layoutShapeReason != null) {
                    findings.Add(new LegacyPptWriteFinding(LegacyPptFeature.Layouts,
                        "PPT-WRITE-LAYOUT-SHAPE", layoutShapeReason, slideIndex));
                }
                shapes = LegacyPptWriter.FlattenShapeTreeForWrite(shapes,
                    out string? groupShapeReason);
                if (groupShapeReason != null) {
                    findings.Add(new LegacyPptWriteFinding(
                        LegacyPptFeature.Groups,
                        "PPT-WRITE-GROUP-SHAPE", groupShapeReason,
                        slideIndex));
                }
                if (HasUnsupportedRichNotes(slide)) {
                    findings.Add(new LegacyPptWriteFinding(LegacyPptFeature.RichNotes,
                        "PPT-WRITE-RICH-NOTES",
                        "Notes-page drawings beyond the speaker-notes body are not encoded by the native binary writer.",
                        slideIndex));
                }
                P.Slide? slideRoot = slide.SlidePart.Slide;
                if (!LegacyPptWriter.TryReadTransition(slide, out _,
                        out string? transitionReason)) {
                    findings.Add(new LegacyPptWriteFinding(LegacyPptFeature.Transitions, "PPT-WRITE-TRANSITION",
                        transitionReason ?? "The slide transition cannot be encoded by the native binary writer.",
                        slideIndex));
                }
                if (slideRoot?.Timing != null) {
                    var animationSounds = new LegacyPptWriter.LegacyPptWriterSoundCatalog();
                    if (!LegacyPptWriter.TryReadClassicAnimations(new[] { slide },
                            animationSounds, out _, out string? animationReason)) {
                        findings.Add(new LegacyPptWriteFinding(LegacyPptFeature.Animations,
                            "PPT-WRITE-TIMING",
                            animationReason ?? "The timing tree cannot be encoded as classic binary PowerPoint animations.",
                            slideIndex));
                    }
                }
                if (!LegacyPptWriter.TryReadBackground(slide, out _,
                        out string? backgroundReason)) {
                    findings.Add(new LegacyPptWriteFinding(LegacyPptFeature.Backgrounds, "PPT-WRITE-BACKGROUND",
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
                            "PPT-WRITE-PLACEHOLDER",
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
                        findings.Add(new LegacyPptWriteFinding(MapShapeFeature(shape), "PPT-WRITE-SHAPE",
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
                            "PPT-WRITE-TABLE",
                            tableReason
                            ?? "The DrawingML table cannot be encoded as a native binary PowerPoint table.",
                            slideIndex, shapeIndex));
                    } else if (shape is PowerPointChart) {
                        findings.Add(new LegacyPptWriteFinding(
                            LegacyPptFeature.Charts,
                            "PPT-WRITE-CHART-CONVERTED",
                            "The editable DrawingML chart will be converted to a static PNG picture in binary PowerPoint. Chart data and editability will not survive the conversion.",
                            slideIndex, shapeIndex));
                    } else if (shape is PowerPointSmartArt) {
                        findings.Add(new LegacyPptWriteFinding(
                            LegacyPptFeature.SmartArt,
                            "PPT-WRITE-SMARTART-CONVERTED",
                            "The editable SmartArt diagram will be converted to a static PNG picture in binary PowerPoint. Diagram structure and editability will not survive the conversion.",
                            slideIndex, shapeIndex));
                    }
                    if (HasUnsupportedVisualStyle(shape)) {
                        findings.Add(new LegacyPptWriteFinding(LegacyPptFeature.ShapeStyles, "PPT-WRITE-SHAPE-STYLE",
                            "Fill, outline, transform, effects, hyperlink, or visibility styling is not encoded.",
                            slideIndex, shapeIndex));
                    }
                    if (!LegacyPptWriter.TryReadShapeMetadataForWrite(shape,
                            out _, out string? metadataReason)) {
                        findings.Add(new LegacyPptWriteFinding(
                            LegacyPptFeature.AccessibilityMetadata,
                            "PPT-WRITE-ACCESSIBILITY-METADATA",
                            metadataReason
                            ?? "The shape accessibility metadata has no classic binary representation.",
                            slideIndex, shapeIndex));
                    }
                    if (LegacyPptWriter.IsLayoutShape(shape)
                        && HasUnsupportedMasterInteraction(shape)) {
                        findings.Add(new LegacyPptWriteFinding(
                            LegacyPptFeature.Layouts,
                            "PPT-WRITE-LAYOUT-INTERACTION",
                            "Interactions on a materialized ordinary-layout shape are not encoded by the native binary writer.",
                            slideIndex, shapeIndex));
                    }
                    if (shape is PowerPointTextBox textBox
                        && !LegacyPptWriter.TryReadTextBoxForWrite(textBox,
                            shapeTextFonts, pictureBullets,
                            out string? textReason)) {
                        findings.Add(new LegacyPptWriteFinding(LegacyPptFeature.RichText, "PPT-WRITE-RICH-TEXT",
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
                        "PPT-WRITE-IMPORT-LOSS",
                        $"Imported legacy content was not fully projected: {diagnostic.Code}."));
                }
            }
            return new LegacyPptWritePreflightReport(findings);
        }

        internal static bool CanWriteSlideLosslessly(PowerPointSlide slide) {
            if (slide == null) throw new ArgumentNullException(nameof(slide));
            if (LegacyPptWriter.ShouldWriteNotesPage(slide, out _)) return false;
            P.Slide? slideRoot = slide.SlidePart.Slide;
            bool animationsSupported = slideRoot?.Timing == null
                || LegacyPptWriter.TryReadClassicAnimations(new[] { slide },
                    new LegacyPptWriter.LegacyPptWriterSoundCatalog(),
                    out _, out _);
            if (!LegacyPptWriter.TryReadTransition(slide, out _, out _)
                || !animationsSupported
                || !LegacyPptWriter.TryReadBackground(slide, out _, out _)) {
                return false;
            }
            if (!LegacyPptWriter.TryReadInteractions(new[] { slide },
                    out _, out _)) return false;
            IReadOnlyList<PowerPointShape> shapes = LegacyPptWriter
                .ReadSlideShapesForWrite(slide, out string? layoutShapeReason);
            if (layoutShapeReason != null) return false;
            shapes = LegacyPptWriter.FlattenShapeTreeForWrite(shapes,
                out string? groupShapeReason);
            if (groupShapeReason != null) return false;
            LegacyPptWriter.LegacyPptWriterFontCatalog fonts =
                LegacyPptWriter.CreateFontCatalogForWrite();
            PresentationPart? presentationPart = slide.SlidePart
                .GetParentParts().OfType<PresentationPart>().FirstOrDefault();
            if (presentationPart == null
                || !LegacyPptWriter.TryReadPictureBulletCatalog(
                    presentationPart, out LegacyPptWriter
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
            || includePictures && shape is PowerPointPicture
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

        private static bool HasUnsupportedRichNotes(PowerPointSlide slide) {
            P.ShapeTree? tree = slide.SlidePart.NotesSlidePart?.NotesSlide?
                .CommonSlideData?.ShapeTree;
            if (tree == null) return false;
            int bodyPlaceholderCount = 0;
            foreach (OpenXmlElement child in tree.ChildElements) {
                if (child is P.NonVisualGroupShapeProperties or P.GroupShapeProperties) continue;
                if (child is not P.Shape shape) return true;
                P.PlaceholderShape? placeholder = shape.NonVisualShapeProperties?
                    .ApplicationNonVisualDrawingProperties?.PlaceholderShape;
                if (placeholder?.Type?.Value != P.PlaceholderValues.Body) return true;
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
                    "PPT-WRITE-MASTER-GROUP", $"{ownerName}: {groupReason}"));
            }
            for (int shapeIndex = 0; shapeIndex < shapes.Count; shapeIndex++) {
                PowerPointShape shape = shapes[shapeIndex];
                string location = $"{ownerName}, shape {shapeIndex}";
                if (!LegacyPptWriter.TryReadPlaceholderForWrite(shape,
                        shapeContext, out _, out string? placeholderReason)) {
                    findings.Add(new LegacyPptWriteFinding(
                        LegacyPptFeature.Placeholders,
                        "PPT-WRITE-MASTER-PLACEHOLDER",
                        $"{location}: {placeholderReason}"));
                }
                if (!LegacyPptWriter.IsSupportedMasterShape(shape)) {
                    findings.Add(new LegacyPptWriteFinding(
                        MapShapeFeature(shape), "PPT-WRITE-MASTER-SHAPE",
                        $"{location}: {shape.ShapeContentType} content is outside the native writer's supported master-shape subset."));
                    continue;
                }
                if (HasUnsupportedVisualStyle(shape)
                    || HasUnsupportedMasterInteraction(shape)) {
                    findings.Add(new LegacyPptWriteFinding(
                        LegacyPptFeature.ShapeStyles,
                        "PPT-WRITE-MASTER-SHAPE-STYLE",
                        $"{location}: visual styling or interactive content is not encoded on binary masters."));
                }
                if (!LegacyPptWriter.TryReadShapeMetadataForWrite(shape,
                        out _, out string? metadataReason)) {
                    findings.Add(new LegacyPptWriteFinding(
                        LegacyPptFeature.AccessibilityMetadata,
                        "PPT-WRITE-MASTER-ACCESSIBILITY-METADATA",
                        $"{location}: {metadataReason}"));
                }
                if (shape is PowerPointTextBox textBox
                    && !LegacyPptWriter.TryReadTextBoxForWrite(textBox,
                        fonts, pictureBullets,
                        out string? textReason)) {
                    findings.Add(new LegacyPptWriteFinding(
                        LegacyPptFeature.RichText,
                        "PPT-WRITE-MASTER-RICH-TEXT",
                        $"{location}: {textReason}"));
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
            || shape.Element.Descendants<A.EffectDag>().Any();

        private static bool HasUnsupportedMasterInteraction(
            PowerPointShape shape) => shape.Element
                .Descendants<A.HyperlinkOnClick>().Any()
            || shape.Element.Descendants<A.HyperlinkOnHover>().Any()
            || shape.Element.Descendants<A.HyperlinkOnMouseOver>().Any();

        internal static void ThrowIfBlocked(LegacyPptWritePreflightReport report, PowerPointSaveOptions? options) {
            if (!report.HasConversionLoss || options?.LossPolicy == PowerPointConversionLossPolicy.Allow) return;
            string details = string.Join("; ", report.Findings.Take(8));
            throw new NotSupportedException(
                "Native PPT/POT/PPS saving is blocked because known content cannot be encoded without loss. "
                + details + " Set PowerPointSaveOptions.LossPolicy to Allow only when that loss is intentional.");
        }
    }
}
