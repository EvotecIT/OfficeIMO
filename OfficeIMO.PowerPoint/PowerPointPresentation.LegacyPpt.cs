using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        private const double EmusPerLegacyMasterUnit = 1587.5d;

        /// <summary>Loads a binary `.ppt`, `.pot`, or `.pps` file into the normal editable PowerPoint model.</summary>
        public static PowerPointPresentation LoadLegacyPpt(string path, LegacyPptImportOptions? options = null) {
            if (path == null) throw new ArgumentNullException(nameof(path));
            LegacyPptPresentation legacy = LegacyPptPresentation.Load(path, options);
            return ProjectLoadedLegacyPpt(legacy, path,
                PowerPointPresentationLoadRouting.GetFormat(path, legacyDefault: true), new PowerPointLoadOptions());
        }

        /// <summary>Loads a binary PowerPoint stream into the normal editable PowerPoint model.</summary>
        public static PowerPointPresentation LoadLegacyPpt(Stream stream, LegacyPptImportOptions? options = null) {
            LegacyPptPresentation legacy = LegacyPptPresentation.Load(stream, options);
            return ProjectLoadedLegacyPpt(legacy, sourcePath: null, PowerPointFileFormat.Ppt,
                new PowerPointLoadOptions());
        }

        /// <summary>Loads a binary PowerPoint file and returns its projected presentation and import report.</summary>
        public static LegacyPptLoadResult LoadLegacyPptWithReport(string path, LegacyPptImportOptions? options = null) {
            if (path == null) throw new ArgumentNullException(nameof(path));
            LegacyPptPresentation legacy = LegacyPptPresentation.Load(path, options);
            try {
                return new LegacyPptLoadResult(ProjectLoadedLegacyPpt(legacy, path,
                    PowerPointPresentationLoadRouting.GetFormat(path, legacyDefault: true), new PowerPointLoadOptions()), legacy);
            } catch (InvalidDataException exception) {
                return new LegacyPptLoadResult(document: null, legacy, exception);
            }
        }

        /// <summary>Loads a binary PowerPoint stream and returns its projected presentation and import report.</summary>
        public static LegacyPptLoadResult LoadLegacyPptWithReport(Stream stream, LegacyPptImportOptions? options = null) {
            LegacyPptPresentation legacy = LegacyPptPresentation.Load(stream, options);
            try {
                return new LegacyPptLoadResult(ProjectLoadedLegacyPpt(legacy, sourcePath: null,
                    PowerPointFileFormat.Ppt, new PowerPointLoadOptions()), legacy);
            } catch (InvalidDataException exception) {
                return new LegacyPptLoadResult(document: null, legacy, exception);
            }
        }

        private static PowerPointPresentation LoadLegacyPptFromNormalFlow(byte[] bytes, string? sourcePath,
            Stream? sourceStream, PowerPointLoadOptions options) {
            if (options.PersistenceMode == DocumentPersistenceMode.SaveOnDispose && sourceStream == null
                && string.IsNullOrEmpty(sourcePath)) {
                throw new NotSupportedException("SaveOnDispose requires an associated destination for binary PowerPoint sources.");
            }
            LegacyPptPresentation legacy = LegacyPptPresentation.Load(bytes, options.LegacyPptImportOptions);
            PowerPointFileFormat sourceFormat = PowerPointPresentationLoadRouting.GetFormat(sourcePath, legacyDefault: true);
            return ProjectLoadedLegacyPpt(legacy, sourcePath, sourceFormat, options, sourceStream);
        }

        private static PowerPointPresentation ProjectLoadedLegacyPpt(LegacyPptPresentation legacy,
            string? sourcePath, PowerPointFileFormat sourceFormat, PowerPointLoadOptions loadOptions,
            Stream? sourceStream = null) {
            if (legacy == null) throw new ArgumentNullException(nameof(legacy));
            using PowerPointPresentation projected = Create();
            projected.SlideSize.SetSizeEmus(ToEmus(legacy.SlideWidth), ToEmus(legacy.SlideHeight));
            IReadOnlyDictionary<uint, LegacyPptLayoutTarget> layoutTargets =
                ProjectLegacyMasters(projected, legacy);

            foreach (LegacyPptSlide legacySlide in legacy.Slides) {
                PowerPointSlide slide = layoutTargets.TryGetValue(legacySlide.MasterId,
                    out LegacyPptLayoutTarget target)
                    ? projected.AddSlide(target.MasterIndex, target.LayoutIndex)
                    : projected.AddSlide();
                slide.Hidden = legacySlide.Hidden;
                ProjectLegacySlideDesign(slide, legacySlide);
                var projectedShapeIds = new Dictionary<uint, uint>();
                foreach (LegacyPptShape shape in legacySlide.Shapes) {
                    OpenXmlElement? projectedShape = ProjectLegacyShape(slide, shape);
                    if (projectedShape != null) {
                        RegisterLegacyShapeIds(shape, projectedShape, projectedShapeIds);
                    }
                }
                ProjectLegacyConnectorRules(slide, legacySlide.ConnectorRules, projectedShapeIds);
                if (!string.IsNullOrWhiteSpace(legacySlide.NotesText)) {
                    slide.Notes.Text = legacySlide.NotesText;
                }
            }

            byte[] packageBytes = projected.ToBytes();
            PowerPointPresentation presentation = LoadPackage(packageBytes, sourcePath, sourceStream, loadOptions);
            LegacyPptProjectionMap projectionMap = LegacyPptProjectionMap.Create(presentation, legacy);
            presentation.MarkLoadedFromLegacyPpt(sourcePath, legacy, projectionMap, sourceFormat);
            return presentation;
        }

        private static OpenXmlElement? ProjectLegacyShape(PowerPointSlide slide, LegacyPptShape shape) {
            long left = ToEmus(shape.Bounds.Left);
            long top = ToEmus(shape.Bounds.Top);
            long width = Math.Max(1L, ToEmus(shape.Bounds.Width));
            long height = Math.Max(1L, ToEmus(shape.Bounds.Height));
            PowerPointShape? projectedShape = null;
            switch (shape.Kind) {
                case LegacyPptShapeKind.TextBox:
                    PowerPointTextBox textBox = shape.PlaceholderKind == LegacyPptPlaceholderKind.Title
                        || shape.PlaceholderKind == LegacyPptPlaceholderKind.CenterTitle
                        || shape.PlaceholderKind == LegacyPptPlaceholderKind.VerticalTitle
                        ? slide.AddTitle(shape.Text, left, top, width, height)
                        : slide.AddTextBox(shape.Text, left, top, width, height);
                    PlaceholderValues? placeholder = MapPlaceholder(shape.PlaceholderKind);
                    if (placeholder.HasValue) textBox.PlaceholderType = placeholder.Value;
                    if (shape.OfficeArtShapeType != 202
                        && LegacyPptShapeGeometryMapper.TryGetPreset(shape.OfficeArtShapeType,
                            out A.ShapeTypeValues textGeometry)
                        && textBox.Element is DocumentFormat.OpenXml.Presentation.Shape textShape
                        && textShape.ShapeProperties?.GetFirstChild<A.PresetGeometry>() is A.PresetGeometry preset) {
                        preset.Preset = textGeometry;
                    }
                    projectedShape = textBox;
                    break;
                case LegacyPptShapeKind.Rectangle:
                    projectedShape = slide.AddShape(A.ShapeTypeValues.Rectangle, left, top, width, height);
                    break;
                case LegacyPptShapeKind.Ellipse:
                    projectedShape = slide.AddShape(A.ShapeTypeValues.Ellipse, left, top, width, height);
                    break;
                case LegacyPptShapeKind.Line:
                    projectedShape = slide.AddShape(A.ShapeTypeValues.Line, left, top, width, height);
                    break;
                case LegacyPptShapeKind.AutoShape:
                    if (LegacyPptShapeGeometryMapper.TryGetPreset(shape.OfficeArtShapeType,
                            out A.ShapeTypeValues geometry)) {
                        projectedShape = slide.AddShape(geometry, left, top, width, height);
                    }
                    break;
                case LegacyPptShapeKind.Connector:
                    if (LegacyPptShapeGeometryMapper.TryGetPreset(shape.OfficeArtShapeType,
                            out A.ShapeTypeValues connectorGeometry)) {
                        projectedShape = slide.AddConnectionShape(connectorGeometry, left, top, width, height);
                    }
                    break;
                case LegacyPptShapeKind.Picture:
                    if (shape.Picture?.HasImportableImage == true && shape.Picture.ContentType != null) {
                        using var image = new MemoryStream(shape.Picture.ImageBytes, writable: false);
                        projectedShape = slide.AddPicture(image,
                            GetLegacyPicturePartType(shape.Picture.ContentType), left, top, width, height);
                    }
                    break;
                case LegacyPptShapeKind.Group:
                    ShapeTree tree = slide.SlidePart.Slide?.CommonSlideData?.ShapeTree
                        ?? throw new InvalidDataException("The projected slide has no shape tree.");
                    uint nextShapeId = tree.Descendants<NonVisualDrawingProperties>()
                        .Select(item => item.Id?.Value ?? 0U)
                        .DefaultIfEmpty(1U)
                        .Max() + 1U;
                    OpenXmlElement? group = CreateLegacyOpenXmlShape(slide.SlidePart, shape,
                        ref nextShapeId);
                    if (group != null) tree.Append(group);
                    return group;
            }
            DocumentFormat.OpenXml.Presentation.ShapeProperties? projectedProperties =
                projectedShape?.Element switch {
                    DocumentFormat.OpenXml.Presentation.Shape projected => projected.ShapeProperties,
                    DocumentFormat.OpenXml.Presentation.ConnectionShape projected => projected.ShapeProperties,
                    DocumentFormat.OpenXml.Presentation.Picture projected => projected.ShapeProperties,
                    _ => null
                };
            if (projectedProperties != null) {
                if (shape.Kind != LegacyPptShapeKind.Picture) {
                    ApplyLegacyShapeStyle(projectedProperties, shape);
                }
                projectedProperties.Transform2D ??= new A.Transform2D();
                ApplyLegacyShapeTransform(projectedProperties.Transform2D, shape);
                if (projectedProperties.GetFirstChild<A.PresetGeometry>() is A.PresetGeometry preset) {
                    LegacyPptShapeGeometryMapper.ApplyExactPresetAdjustments(shape.OfficeArtShapeType,
                        shape.Geometry, preset);
                }
            }
            if (projectedShape?.Element is Picture projectedPicture) {
                ApplyLegacyPictureCrop(projectedPicture.BlipFill, shape);
            }
            return projectedShape?.Element;
        }

        private static void RegisterLegacyShapeIds(LegacyPptShape source, OpenXmlElement projected,
            IDictionary<uint, uint> projectedShapeIds) {
            uint? projectedId = projected switch {
                DocumentFormat.OpenXml.Presentation.Shape item => item.NonVisualShapeProperties?
                    .NonVisualDrawingProperties?.Id?.Value,
                DocumentFormat.OpenXml.Presentation.ConnectionShape item => item.NonVisualConnectionShapeProperties?
                    .NonVisualDrawingProperties?.Id?.Value,
                DocumentFormat.OpenXml.Presentation.Picture item => item.NonVisualPictureProperties?
                    .NonVisualDrawingProperties?.Id?.Value,
                DocumentFormat.OpenXml.Presentation.GroupShape item => item.NonVisualGroupShapeProperties?
                    .NonVisualDrawingProperties?.Id?.Value,
                _ => null
            };
            if (projectedId.HasValue) projectedShapeIds[source.ShapeId] = projectedId.Value;
            if (source.Kind != LegacyPptShapeKind.Group
                || projected is not DocumentFormat.OpenXml.Presentation.GroupShape group) return;

            LegacyPptShape[] sourceChildren = source.Children
                .Where(child => child.Kind != LegacyPptShapeKind.Unsupported)
                .ToArray();
            OpenXmlElement[] projectedChildren = group.ChildElements
                .Where(IsLegacyDrawingElement)
                .ToArray();
            int count = Math.Min(sourceChildren.Length, projectedChildren.Length);
            for (int index = 0; index < count; index++) {
                RegisterLegacyShapeIds(sourceChildren[index], projectedChildren[index], projectedShapeIds);
            }
        }

        private static void ProjectLegacyConnectorRules(PowerPointSlide slide,
            IReadOnlyList<LegacyPptConnectorRule> rules,
            IReadOnlyDictionary<uint, uint> projectedShapeIds) {
            if (rules.Count == 0) return;
            ShapeTree? tree = slide.SlidePart.Slide?.CommonSlideData?.ShapeTree;
            if (tree == null) return;
            ProjectLegacyConnectorRules(tree, rules, projectedShapeIds);
        }

        private static void ProjectLegacyConnectorRules(ShapeTree tree,
            IReadOnlyList<LegacyPptConnectorRule> rules,
            IReadOnlyDictionary<uint, uint> projectedShapeIds) {
            foreach (LegacyPptConnectorRule rule in rules) {
                if (!projectedShapeIds.TryGetValue(rule.ConnectorShapeId, out uint connectorId)) continue;
                DocumentFormat.OpenXml.Presentation.ConnectionShape? connector = tree
                    .Descendants<DocumentFormat.OpenXml.Presentation.ConnectionShape>()
                    .FirstOrDefault(item => item.NonVisualConnectionShapeProperties?
                        .NonVisualDrawingProperties?.Id?.Value == connectorId);
                if (connector == null) continue;
                NonVisualConnectorShapeDrawingProperties drawingProperties =
                    connector.NonVisualConnectionShapeProperties?.NonVisualConnectorShapeDrawingProperties
                    ?? new NonVisualConnectorShapeDrawingProperties();
                connector.NonVisualConnectionShapeProperties ??= new NonVisualConnectionShapeProperties();
                connector.NonVisualConnectionShapeProperties.NonVisualConnectorShapeDrawingProperties =
                    drawingProperties;
                drawingProperties.RemoveAllChildren<A.StartConnection>();
                drawingProperties.RemoveAllChildren<A.EndConnection>();
                if (projectedShapeIds.TryGetValue(rule.StartShapeId, out uint startShapeId)) {
                    drawingProperties.Append(new A.StartConnection {
                        Id = startShapeId,
                        Index = rule.StartConnectionSiteIndex
                    });
                }
                if (projectedShapeIds.TryGetValue(rule.EndShapeId, out uint endShapeId)) {
                    drawingProperties.Append(new A.EndConnection {
                        Id = endShapeId,
                        Index = rule.EndConnectionSiteIndex
                    });
                }
            }
        }

        private static bool IsLegacyDrawingElement(OpenXmlElement element) =>
            element is DocumentFormat.OpenXml.Presentation.Shape
                or DocumentFormat.OpenXml.Presentation.ConnectionShape
                or DocumentFormat.OpenXml.Presentation.Picture
                or DocumentFormat.OpenXml.Presentation.GroupShape;

        private static ImagePartType GetLegacyPicturePartType(string contentType) => contentType switch {
            "image/png" => ImagePartType.Png,
            "image/jpeg" => ImagePartType.Jpeg,
            "image/bmp" => ImagePartType.Bmp,
            "image/tiff" => ImagePartType.Tiff,
            "image/x-emf" => ImagePartType.Emf,
            "image/x-wmf" => ImagePartType.Wmf,
            _ => throw new NotSupportedException($"Legacy picture content type '{contentType}' is not supported.")
        };

        private static PlaceholderValues? MapPlaceholder(LegacyPptPlaceholderKind placeholder) {
            switch (placeholder) {
                case LegacyPptPlaceholderKind.MasterTitle:
                case LegacyPptPlaceholderKind.Title: return PlaceholderValues.Title;
                case LegacyPptPlaceholderKind.MasterCenterTitle:
                case LegacyPptPlaceholderKind.CenterTitle: return PlaceholderValues.CenteredTitle;
                case LegacyPptPlaceholderKind.MasterSubtitle:
                case LegacyPptPlaceholderKind.Subtitle: return PlaceholderValues.SubTitle;
                case LegacyPptPlaceholderKind.MasterBody:
                case LegacyPptPlaceholderKind.Body: return PlaceholderValues.Body;
                case LegacyPptPlaceholderKind.VerticalTitle: return PlaceholderValues.Title;
                case LegacyPptPlaceholderKind.VerticalBody: return PlaceholderValues.Body;
                case LegacyPptPlaceholderKind.MasterNotesSlideImage:
                case LegacyPptPlaceholderKind.NotesSlideImage: return PlaceholderValues.SlideImage;
                case LegacyPptPlaceholderKind.MasterNotesBody:
                case LegacyPptPlaceholderKind.NotesBody: return PlaceholderValues.Body;
                case LegacyPptPlaceholderKind.MasterDate: return PlaceholderValues.DateAndTime;
                case LegacyPptPlaceholderKind.MasterSlideNumber: return PlaceholderValues.SlideNumber;
                case LegacyPptPlaceholderKind.MasterFooter: return PlaceholderValues.Footer;
                case LegacyPptPlaceholderKind.MasterHeader: return PlaceholderValues.Header;
                case LegacyPptPlaceholderKind.Graph: return PlaceholderValues.Chart;
                case LegacyPptPlaceholderKind.Table: return PlaceholderValues.Table;
                case LegacyPptPlaceholderKind.ClipArt: return PlaceholderValues.ClipArt;
                case LegacyPptPlaceholderKind.Media: return PlaceholderValues.Media;
                case LegacyPptPlaceholderKind.Picture: return PlaceholderValues.Picture;
                case LegacyPptPlaceholderKind.Object:
                case LegacyPptPlaceholderKind.OrganizationChart:
                case LegacyPptPlaceholderKind.VerticalObject: return PlaceholderValues.Object;
                default: return null;
            }
        }

        private static long ToEmus(int masterUnits) => checked((long)Math.Round(
            masterUnits * EmusPerLegacyMasterUnit, MidpointRounding.AwayFromZero));
    }
}
