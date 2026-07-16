using System.Collections.ObjectModel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>
    /// Retains a global package guard plus independent normalized slide and master-theme guards. Existing projected
    /// slides can be reordered or removed, and mapped master themes can be edited, while unsupported mutations to a
    /// retained slide or shared package part are rejected.
    /// </summary>
    internal sealed class LegacyPptProjectionFingerprint {
        private delegate bool TryGetShapeProjection(uint openXmlShapeId,
            out LegacyPptShapeProjection? projection);

        private const string ClassicAnimationExtensionUri =
            "{5BA743F1-2B69-4BB9-B2E0-4A418B7E7435}";
        private readonly IReadOnlyDictionary<string, string> _slides;

        private LegacyPptProjectionFingerprint(string global, IReadOnlyDictionary<string, string> slides) {
            Global = global;
            _slides = new ReadOnlyDictionary<string, string>(slides.ToDictionary(
                pair => pair.Key, pair => pair.Value, StringComparer.Ordinal));
        }

        internal string Global { get; }

        internal static LegacyPptProjectionFingerprint Create(PresentationDocument document,
            LegacyPptProjectionMap projectionMap) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (projectionMap == null) throw new ArgumentNullException(nameof(projectionMap));
            var slides = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (SlidePart slidePart in document.PresentationPart?.SlideParts ?? Enumerable.Empty<SlidePart>()) {
                if (projectionMap.Slides.Any(slide => string.Equals(slide.SlidePartUri,
                        slidePart.Uri.ToString(), StringComparison.Ordinal))) {
                    slides.Add(slidePart.Uri.ToString(), CreateSlide(document, slidePart, projectionMap));
                }
            }
            if (slides.Count != projectionMap.Slides.Count) {
                throw new InvalidDataException("The projected slide fingerprint set is incomplete.");
            }
            return new LegacyPptProjectionFingerprint(CreateGlobal(document,
                projectionMap), slides);
        }

        internal bool Matches(PresentationDocument document, LegacyPptProjectionMap projectionMap) {
            if (!string.Equals(Global, CreateGlobal(document, projectionMap),
                    StringComparison.Ordinal)) return false;
            SlidePart[] currentSlides = document.PresentationPart?.SlideParts.ToArray() ?? Array.Empty<SlidePart>();
            foreach (SlidePart slidePart in currentSlides) {
                string uri = slidePart.Uri.ToString();
                if (_slides.TryGetValue(uri, out string? expected)
                    && !string.Equals(expected, CreateSlide(document, slidePart, projectionMap),
                        StringComparison.Ordinal)) {
                    return false;
                }
            }
            return true;
        }

        private static string CreateGlobal(PresentationDocument document,
            LegacyPptProjectionMap projectionMap) {
            ISet<string> materializedLayoutThemePartUris = new HashSet<string>(
                document.PresentationPart?.SlideMasterParts
                    .SelectMany(master => master.SlideLayoutParts)
                    .Where(layout => projectionMap
                        .IsEditableProjectedLayoutThemePart(
                            layout.Uri.ToString()))
                    .Select(layout => layout.ThemeOverridePart?.Uri.ToString())
                    .Where(uri => uri != null).Cast<string>()
                ?? Enumerable.Empty<string>(), StringComparer.Ordinal);
            return PowerPointPackageFingerprint.Create(document,
                (part, root) => {
                    if (part is PresentationPart) NormalizePresentationTopology(root);
                    if (part is ExtendedFilePropertiesPart) {
                        NormalizeProjectedExtendedProperties(root);
                    }
                    if (part is SlideMasterPart masterPart
                        && projectionMap.TryGetMaster(masterPart,
                            out LegacyPptMasterProjection? masterProjection)
                        && masterProjection != null) {
                        NormalizeProjectedMaster(root, masterProjection);
                    }
                    if (part is NotesMasterPart or HandoutMasterPart
                        && projectionMap.TryGetSpecialMaster(part,
                            out LegacyPptMasterProjection? specialProjection)
                        && specialProjection != null) {
                        NormalizeProjectedSpecialMaster(root, specialProjection);
                    }
                    if (part is ThemePart or ThemeOverridePart
                        && projectionMap.IsProjectedMasterThemePart(part.Uri.ToString())) {
                        NormalizeProjectedMasterTheme(root);
                    }
                    if (part is SlideLayoutPart
                        && projectionMap.IsProjectedLayoutPart(part.Uri.ToString())) {
                        NormalizeProjectedHeaderFooter(root);
                    }
                    if (part is SlideLayoutPart backgroundLayout
                        && projectionMap.IsEditableProjectedLayoutBackgroundPart(
                            backgroundLayout.Uri.ToString())
                        && root is P.SlideLayout normalizedLayout) {
                        NormalizeProjectedBackground(
                            normalizedLayout.CommonSlideData);
                    }
                    if (part is SlideLayoutPart themeLayout
                        && projectionMap.IsEditableProjectedLayoutThemePart(
                            themeLayout.Uri.ToString())) {
                        NormalizeProjectedLayoutTheme(root);
                    }
                    if (part is SlideLayoutPart titlePart
                        && projectionMap.TryGetTitleMaster(titlePart,
                            out LegacyPptMasterProjection? titleProjection)
                        && titleProjection != null) {
                        NormalizeProjectedTitleMaster(root, titleProjection);
                    }
                },
                part => !(part is SlidePart or NotesSlidePart or SlideCommentsPart
                    or CommentAuthorsPart or CoreFilePropertiesPart
                    or CustomFilePropertiesPart or VbaProjectPart)
                    && !projectionMap.IsProjectedOlePart(
                        part.Uri.ToString())
                    && !materializedLayoutThemePartUris.Contains(
                        part.Uri.ToString()),
                (owner, relationship) => !(relationship.OpenXmlPart is SlidePart
                    or SlideCommentsPart or CommentAuthorsPart or VbaProjectPart)
                    && !(owner is SlideLayoutPart layout
                        && relationship.OpenXmlPart is ThemeOverridePart
                        && projectionMap.IsEditableProjectedLayoutThemePart(
                            layout.Uri.ToString())),
                includePackageProperties: false);
        }

        private static void NormalizeProjectedMaster(OpenXmlElement root,
            LegacyPptMasterProjection projection) {
            if (root is not P.SlideMaster master) return;
            if (master.ColorMap != null) {
                master.ColorMap.ClearAllAttributes();
                master.ColorMap.RemoveAllChildren();
            }
            NormalizeProjectedBackground(master.CommonSlideData);
            NormalizeProjectedShapes(root, projection.TryGetShape,
                normalizeInteractions: false);
        }

        private static void NormalizeProjectedMasterTheme(OpenXmlElement root) {
            if (root is A.Theme theme) {
                theme.ClearAllAttributes();
                theme.RemoveAllChildren();
            } else if (root is A.ThemeOverride themeOverride) {
                themeOverride.ClearAllAttributes();
                themeOverride.RemoveAllChildren();
            }
        }

        private static void NormalizeProjectedTitleMaster(OpenXmlElement root,
            LegacyPptMasterProjection projection) {
            if (root is not P.SlideLayout layout) return;
            layout.ShowMasterShapes = null;
            if (layout.ColorMapOverride != null) {
                layout.ColorMapOverride.ClearAllAttributes();
                layout.ColorMapOverride.RemoveAllChildren();
            }
            NormalizeProjectedBackground(layout.CommonSlideData);
            NormalizeProjectedShapes(root, projection.TryGetShape,
                normalizeInteractions: false);
        }

        private static void NormalizeProjectedSpecialMaster(OpenXmlElement root,
            LegacyPptMasterProjection projection) {
            switch (root) {
                case P.NotesMaster notes when notes.ColorMap != null:
                    notes.ColorMap.ClearAllAttributes();
                    notes.ColorMap.RemoveAllChildren();
                    NormalizeProjectedBackground(notes.CommonSlideData);
                    break;
                case P.HandoutMaster handout when handout.ColorMap != null:
                    handout.ColorMap.ClearAllAttributes();
                    handout.ColorMap.RemoveAllChildren();
                    NormalizeProjectedBackground(handout.CommonSlideData);
                    break;
                case P.NotesMaster notes:
                    NormalizeProjectedBackground(notes.CommonSlideData);
                    break;
                case P.HandoutMaster handout:
                    NormalizeProjectedBackground(handout.CommonSlideData);
                    break;
            }
            NormalizeProjectedShapes(root, projection.TryGetShape,
                normalizeInteractions: false);
        }

        private static void NormalizeProjectedBackground(
            P.CommonSlideData? commonSlideData) {
            P.Background? background = commonSlideData?.Background;
            if (background == null) return;
            background.Remove();
        }

        private static string CreateSlide(PresentationDocument document, SlidePart slidePart,
            LegacyPptProjectionMap projectionMap) => PowerPointPackageFingerprint.Create(document,
            (part, root) => NormalizeProjectedSlide(root, slidePart.Uri, projectionMap),
            part => string.Equals(part.Uri.ToString(), slidePart.Uri.ToString(),
                        StringComparison.Ordinal)
                    || part is NotesSlidePart notesPart
                    && ReferenceEquals(notesPart.SlidePart, slidePart),
            (owner, relationship) => relationship.OpenXmlPart is not SlideCommentsPart
                and not SlidePart,
            (owner, relationship) => relationship is not HyperlinkRelationship,
            includePackageProperties: false);

        private static void NormalizePresentationTopology(OpenXmlElement root) {
            if (root is not P.Presentation presentation) return;
            presentation.SlideIdList?.RemoveAllChildren<P.SlideId>();
            presentation.RemoveAllChildren<P.CustomShowList>();
        }

        private static void NormalizeProjectedExtendedProperties(
            OpenXmlElement root) {
            if (root is not DocumentFormat.OpenXml.ExtendedProperties.Properties
                properties) return;
            properties.Application?.Remove();
            properties.TotalTime?.Remove();
            properties.PresentationFormat?.Remove();
            properties.Slides?.Remove();
            properties.Notes?.Remove();
            properties.HiddenSlides?.Remove();
            properties.Manager?.Remove();
            properties.Company?.Remove();
        }

        private static void NormalizeProjectedHeaderFooter(OpenXmlElement root) {
            if (root is not P.SlideLayout layout) return;
            layout.RemoveAllChildren<P.HeaderFooter>();
            foreach (P.Shape shape in layout.CommonSlideData?.ShapeTree?
                         .Elements<P.Shape>() ?? Enumerable.Empty<P.Shape>()) {
                P.PlaceholderValues? type = shape.NonVisualShapeProperties?
                    .ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value;
                if (type != P.PlaceholderValues.DateAndTime
                    && type != P.PlaceholderValues.Footer
                    && type != P.PlaceholderValues.SlideNumber) continue;
                shape.TextBody?.RemoveAllChildren<A.Paragraph>();
            }
        }

        private static void NormalizeProjectedLayoutTheme(OpenXmlElement root) {
            if (root is not P.SlideLayout layout) return;
            layout.ColorMapOverride?.Remove();
        }

        private static void NormalizeProjectedSlide(OpenXmlElement root, Uri partUri,
            LegacyPptProjectionMap projectionMap) {
            LegacyPptSlideProjection? slideProjection = projectionMap.Slides.FirstOrDefault(slide =>
                string.Equals(slide.SlidePartUri, partUri.ToString(), StringComparison.Ordinal));
            if (slideProjection == null) return;
            if (root is P.NotesSlide notesRoot && slideProjection.Notes != null) {
                notesRoot.ShowMasterShapes = null;
                if (notesRoot.ColorMapOverride != null) {
                    notesRoot.ColorMapOverride.ClearAllAttributes();
                    notesRoot.ColorMapOverride.RemoveAllChildren();
                }
                NormalizeProjectedBackground(notesRoot.CommonSlideData);
                foreach (P.Shape shape in notesRoot.CommonSlideData?.ShapeTree?
                             .Elements<P.Shape>() ?? Enumerable.Empty<P.Shape>()) {
                    P.PlaceholderShape? placeholder = shape.NonVisualShapeProperties?
                        .ApplicationNonVisualDrawingProperties?.PlaceholderShape;
                    if (placeholder?.Type?.Value == P.PlaceholderValues.Body
                        && shape.TextBody != null) {
                        shape.TextBody.RemoveAllChildren<A.Paragraph>();
                    }
                }
                return;
            }
            if (root is P.Slide slideRoot) {
                slideRoot.Show = null;
                slideRoot.ShowMasterShapes = null;
                slideRoot.Transition = null;
                if (slideRoot.ColorMapOverride != null) {
                    slideRoot.ColorMapOverride.ClearAllAttributes();
                    slideRoot.ColorMapOverride.RemoveAllChildren();
                }
                NormalizeProjectedBackground(slideRoot.CommonSlideData);
                P.SlideExtensionList? extensions = slideRoot
                    .GetFirstChild<P.SlideExtensionList>();
                P.SlideExtension? classicAnimations = extensions?
                    .Elements<P.SlideExtension>().FirstOrDefault(extension =>
                        string.Equals(extension.Uri?.Value,
                            ClassicAnimationExtensionUri,
                            StringComparison.Ordinal));
                if (classicAnimations != null
                    && HasOnlyTopLevelAnimationTargets(classicAnimations,
                        slideProjection)) {
                    slideRoot.Timing?.Remove();
                    classicAnimations.Remove();
                    if (extensions!.ChildElements.Count == 0) extensions.Remove();
                }
            }

            NormalizeProjectedShapes(root, slideProjection.TryGetShape,
                normalizeInteractions: true);
        }

        private static void NormalizeProjectedShapes(OpenXmlElement root,
            TryGetShapeProjection tryGetShape, bool normalizeInteractions) {
            foreach (P.Shape shape in root.Descendants<P.Shape>()) {
                uint? shapeId = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value;
                if (!shapeId.HasValue || !tryGetShape(shapeId.Value,
                        out LegacyPptShapeProjection? shapeProjection)
                    || shapeProjection == null) continue;
                shape.NonVisualShapeProperties?
                    .ApplicationNonVisualDrawingProperties?.PlaceholderShape?
                    .Remove();
                if (normalizeInteractions && shapeProjection.CanEditInteractions) {
                    shape.NonVisualShapeProperties?.NonVisualDrawingProperties?
                        .RemoveAllChildren<A.HyperlinkOnClick>();
                    shape.NonVisualShapeProperties?.NonVisualDrawingProperties?
                        .RemoveAllChildren<A.HyperlinkOnHover>();
                }
                var transform = shape.ShapeProperties?.Transform2D;
                if (transform?.Offset != null) {
                    transform.Offset.X = 0L;
                    transform.Offset.Y = 0L;
                }
                if (transform?.Extents != null) {
                    transform.Extents.Cx = 0L;
                    transform.Extents.Cy = 0L;
                }
                if (shape.TextBody != null) shape.TextBody.RemoveAllChildren<A.Paragraph>();
            }
            foreach (P.Picture picture in root.Descendants<P.Picture>()) {
                uint? shapeId = picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value;
                if (!shapeId.HasValue || !tryGetShape(shapeId.Value,
                        out LegacyPptShapeProjection? shapeProjection)
                    || shapeProjection == null) continue;
                picture.NonVisualPictureProperties?
                    .ApplicationNonVisualDrawingProperties?.PlaceholderShape?
                    .Remove();
                if (normalizeInteractions && shapeProjection.CanEditInteractions) {
                    picture.NonVisualPictureProperties?.NonVisualDrawingProperties?
                        .RemoveAllChildren<A.HyperlinkOnClick>();
                    picture.NonVisualPictureProperties?.NonVisualDrawingProperties?
                        .RemoveAllChildren<A.HyperlinkOnHover>();
                }
                A.Transform2D? transform = picture.ShapeProperties?.Transform2D;
                if (transform?.Offset != null) {
                    transform.Offset.X = 0L;
                    transform.Offset.Y = 0L;
                }
                if (transform?.Extents != null) {
                    transform.Extents.Cx = 0L;
                    transform.Extents.Cy = 0L;
                }
                if (shapeProjection.CanEditPictureFormatting) {
                    picture.BlipFill?.SourceRectangle?.Remove();
                    A.Blip? blip = picture.BlipFill?.Blip;
                    blip?.RemoveAllChildren<A.LuminanceEffect>();
                    blip?.RemoveAllChildren<A.Grayscale>();
                    blip?.RemoveAllChildren<A.BiLevel>();
                    blip?.RemoveAllChildren<A.ColorChange>();
                    blip?.RemoveAllChildren<A.ColorReplacement>();
                }
            }
            foreach (P.ConnectionShape connection in root.Descendants<P.ConnectionShape>()) {
                uint? shapeId = connection.NonVisualConnectionShapeProperties?
                    .NonVisualDrawingProperties?.Id?.Value;
                if (!shapeId.HasValue || !tryGetShape(shapeId.Value,
                        out LegacyPptShapeProjection? shapeProjection)
                    || shapeProjection == null) continue;
                connection.NonVisualConnectionShapeProperties?
                    .ApplicationNonVisualDrawingProperties?.PlaceholderShape?
                    .Remove();
                if (normalizeInteractions && shapeProjection.CanEditInteractions) {
                    connection.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties?
                        .RemoveAllChildren<A.HyperlinkOnClick>();
                    connection.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties?
                        .RemoveAllChildren<A.HyperlinkOnHover>();
                }
                A.Transform2D? transform = connection.ShapeProperties?.Transform2D;
                if (transform?.Offset != null) {
                    transform.Offset.X = 0L;
                    transform.Offset.Y = 0L;
                }
                if (transform?.Extents != null) {
                    transform.Extents.Cx = 0L;
                    transform.Extents.Cy = 0L;
                }
            }
            foreach (P.GroupShape group in root.Descendants<P.GroupShape>()) {
                uint? shapeId = group.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Id?.Value;
                if (!shapeId.HasValue || !tryGetShape(shapeId.Value,
                        out LegacyPptShapeProjection? shapeProjection)
                    || shapeProjection == null) continue;
                if (normalizeInteractions && shapeProjection.CanEditInteractions) {
                    group.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?
                        .RemoveAllChildren<A.HyperlinkOnClick>();
                    group.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?
                        .RemoveAllChildren<A.HyperlinkOnHover>();
                }
                A.TransformGroup? transform = group.GroupShapeProperties?.TransformGroup;
                if (transform?.Offset != null) {
                    transform.Offset.X = 0L;
                    transform.Offset.Y = 0L;
                }
                if (transform?.Extents != null) {
                    transform.Extents.Cx = 0L;
                    transform.Extents.Cy = 0L;
                }
            }
            foreach (P.GraphicFrame frame in root.Descendants<P.GraphicFrame>()) {
                uint? shapeId = frame.NonVisualGraphicFrameProperties?
                    .NonVisualDrawingProperties?.Id?.Value;
                if (!shapeId.HasValue || !tryGetShape(shapeId.Value,
                        out LegacyPptShapeProjection? shapeProjection)
                    || shapeProjection?.OleObject == null) continue;
                P.Transform? transform = frame.Transform;
                if (transform?.Offset != null) {
                    transform.Offset.X = 0L;
                    transform.Offset.Y = 0L;
                }
                if (transform?.Extents != null) {
                    transform.Extents.Cx = 0L;
                    transform.Extents.Cy = 0L;
                }
                P.OleObject? ole = frame.Graphic?.GraphicData?
                    .GetFirstChild<P.OleObject>();
                if (ole == null) continue;
                ole.ProgId = null;
                ole.ShowAsIcon = null;
                P.OleObjectEmbed? embed = ole
                    .GetFirstChild<P.OleObjectEmbed>();
                if (embed != null) embed.FollowColorScheme = null;
            }
        }

        private static bool HasOnlyTopLevelAnimationTargets(
            P.SlideExtension extension,
            LegacyPptSlideProjection slideProjection) {
            OpenXmlElement[] animations = extension.Descendants()
                .Where(element => element.NamespaceUri ==
                        "https://schemas.officeimo.net/powerpoint/2026/classic-animations"
                    && element.LocalName == "animation")
                .ToArray();
            return animations.Length > 0 && animations.All(animation =>
                uint.TryParse(animation.GetAttributes().FirstOrDefault(attribute =>
                        attribute.LocalName == "shapeId").Value,
                    System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out uint shapeId)
                && slideProjection.TryGetShape(shapeId, out _));
        }
    }
}
