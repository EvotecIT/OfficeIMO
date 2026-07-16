using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Capabilities;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static class LegacyPptWritePreflight {
        internal static LegacyPptWritePreflightReport Analyze(PowerPointPresentation presentation) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            var findings = new List<LegacyPptWriteFinding>();
            if (presentation.CanPreserveOriginalLegacyPackage
                || LegacyPptPreservingWriter.CanWritePresentation(presentation)) {
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
            if (presentation.OpenXmlDocument.PresentationPart?.VbaProjectPart != null) {
                findings.Add(new LegacyPptWriteFinding(LegacyPptFeature.VbaProjects, "PPT-WRITE-VBA",
                    "VBA projects are not encoded by the native binary writer."));
            }
            int masterCount = presentation.OpenXmlDocument.PresentationPart?
                .SlideMasterParts.Count() ?? 0;
            if (masterCount == 0 || masterCount > LegacyPptWriter.MaxNativeMasterCount) {
                findings.Add(new LegacyPptWriteFinding(LegacyPptFeature.Masters,
                    "PPT-WRITE-MASTER-COUNT",
                    masterCount == 0
                        ? "The presentation has no slide master to encode."
                        : $"The native binary writer currently supports at most {LegacyPptWriter.MaxNativeMasterCount} slide masters; the presentation contains {masterCount}."));
            }
            foreach (SlideMasterPart masterPart in presentation.OpenXmlDocument
                         .PresentationPart?.SlideMasterParts
                     ?? Enumerable.Empty<SlideMasterPart>()) {
                if (!LegacyPptWriter.TryReadBackground(masterPart, out _,
                        out string? masterBackgroundReason)) {
                    findings.Add(new LegacyPptWriteFinding(
                        LegacyPptFeature.Backgrounds, "PPT-WRITE-BACKGROUND",
                        masterBackgroundReason
                        ?? "A slide-master background cannot be encoded by the native binary writer."));
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
            if (!LegacyPptWriter.TryReadInteractions(presentation, out _,
                    out string? interactionReason)) {
                findings.Add(new LegacyPptWriteFinding(LegacyPptFeature.Hyperlinks,
                    "PPT-WRITE-INTERACTION",
                    interactionReason ?? "A hyperlink or action cannot be encoded by the native binary writer."));
            }
            for (int slideIndex = 0; slideIndex < presentation.Slides.Count; slideIndex++) {
                PowerPointSlide slide = presentation.Slides[slideIndex];
                IReadOnlyList<PowerPointShape> shapes = LegacyPptWriter
                    .ReadSlideShapesForWrite(slide, out string? layoutShapeReason);
                if (layoutShapeReason != null) {
                    findings.Add(new LegacyPptWriteFinding(LegacyPptFeature.Layouts,
                        "PPT-WRITE-LAYOUT-SHAPE", layoutShapeReason, slideIndex));
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
                    if (!IsSupportedShape(shape)) {
                        findings.Add(new LegacyPptWriteFinding(MapShapeFeature(shape), "PPT-WRITE-SHAPE",
                            $"{shape.ShapeContentType} content is outside the native writer's text/rectangle/ellipse/line subset.",
                            slideIndex, shapeIndex));
                        continue;
                    }
                    if (HasUnsupportedVisualStyle(shape)) {
                        findings.Add(new LegacyPptWriteFinding(LegacyPptFeature.ShapeStyles, "PPT-WRITE-SHAPE-STYLE",
                            "Fill, outline, transform, effects, hyperlink, visibility, or alternative-text styling is not encoded.",
                            slideIndex, shapeIndex));
                    }
                    if (shape is PowerPointTextBox textBox && HasRichTextFormatting(textBox)) {
                        findings.Add(new LegacyPptWriteFinding(LegacyPptFeature.RichText, "PPT-WRITE-RICH-TEXT",
                            "Rich run or paragraph formatting is flattened to plain text.", slideIndex, shapeIndex));
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
            if (slide.SlidePart.NotesSlidePart != null && !string.IsNullOrWhiteSpace(slide.Notes.Text)) return false;
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
            foreach (PowerPointShape shape in shapes) {
                if (!IsSupportedShape(shape) || HasUnsupportedVisualStyle(shape)) return false;
                if (shape is PowerPointTextBox textBox && HasRichTextFormatting(textBox)) return false;
            }
            return true;
        }

        private static bool IsSupportedShape(PowerPointShape shape) => shape is PowerPointTextBox
            || shape is PowerPointAutoShape autoShape
            && (autoShape.ShapeType == A.ShapeTypeValues.Rectangle
                || autoShape.ShapeType == A.ShapeTypeValues.Ellipse
                || autoShape.ShapeType == A.ShapeTypeValues.Line);

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

        private static LegacyPptFeature MapShapeFeature(PowerPointShape shape) {
            switch (shape.ShapeContentType) {
                case PowerPointShapeContentType.Picture: return LegacyPptFeature.RasterPictures;
                case PowerPointShapeContentType.Table: return LegacyPptFeature.Tables;
                case PowerPointShapeContentType.Chart: return LegacyPptFeature.Charts;
                case PowerPointShapeContentType.Group: return LegacyPptFeature.Groups;
                case PowerPointShapeContentType.Media: return LegacyPptFeature.Media;
                case PowerPointShapeContentType.SmartArt: return LegacyPptFeature.SmartArt;
                case PowerPointShapeContentType.OleObject: return LegacyPptFeature.EmbeddedOle;
                case PowerPointShapeContentType.Connector: return LegacyPptFeature.Connectors;
                case PowerPointShapeContentType.AutoShape: return LegacyPptFeature.AutoShapes;
                case PowerPointShapeContentType.TextBox: return LegacyPptFeature.RichText;
                default: return LegacyPptFeature.UnknownRecordsAndStreams;
            }
        }

        private static bool HasUnsupportedVisualStyle(PowerPointShape shape) =>
            shape.FillColor != null
            || shape.FillTransparency != null
            || shape.OutlineColor != null
            || shape.OutlineWidthPoints != null
            || shape.OutlineDash != null
            || shape.Rotation != null
            || shape.HorizontalFlip != null
            || shape.VerticalFlip != null
            || shape.Hidden
            || !string.IsNullOrWhiteSpace(shape.AltText)
            || shape.Element.Descendants<A.EffectList>().Any();

        private static bool HasRichTextFormatting(PowerPointTextBox textBox) {
            P.Shape? shape = textBox.Element as P.Shape;
            if (shape?.TextBody == null) return false;
            return shape.TextBody.Descendants<A.RunProperties>().Any(properties =>
                       properties.HasAttributes || properties.ChildElements.Any(child =>
                           child is not A.HyperlinkOnClick
                               and not A.HyperlinkOnMouseOver))
                || shape.TextBody.Descendants<A.ParagraphProperties>().Any(properties =>
                    properties.HasAttributes || properties.HasChildren);
        }

        internal static void ThrowIfBlocked(LegacyPptWritePreflightReport report, PowerPointSaveOptions? options) {
            if (!report.HasConversionLoss || options?.LossPolicy == PowerPointConversionLossPolicy.Allow) return;
            string details = string.Join("; ", report.Findings.Take(8));
            throw new NotSupportedException(
                "Native PPT/POT/PPS saving is blocked because known content cannot be encoded without loss. "
                + details + " Set PowerPointSaveOptions.LossPolicy to Allow only when that loss is intentional.");
        }
    }
}
