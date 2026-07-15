using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        private static void ProjectLegacyMainMasterLayouts(LegacyPptLayoutCatalog catalog,
            LegacyPptPresentation legacy, LegacyPptMaster master, SlideMasterPart masterPart,
            int masterIndex) {
            LegacyPptSlide[] slides = legacy.Slides.Where(slide => slide.MasterId == master.MasterId)
                .ToArray();
            if (slides.Length == 0) {
                catalog.AddMasterFallback(master.MasterId, new LegacyPptLayoutTarget(masterIndex, 0));
                return;
            }

            SlideLayoutPart firstLayout = masterPart.SlideLayoutParts.First();
            LegacyPptSlide[] representativeSlides = slides
                .GroupBy(LegacyPptLayoutCatalog.CreateKey, StringComparer.Ordinal)
                .Select(group => group.First())
                .ToArray();
            for (int index = 0; index < representativeSlides.Length; index++) {
                LegacyPptSlide representative = representativeSlides[index];
                SlideLayoutPart layoutPart;
                int layoutIndex;
                if (index == 0) {
                    layoutPart = firstLayout;
                    layoutIndex = 0;
                } else {
                    layoutPart = AddLegacyLayoutPart(masterPart);
                    layoutIndex = masterPart.SlideLayoutParts.Count() - 1;
                }

                string name = GetLegacySlideLayoutName(master, representative);
                layoutPart.SlideLayout = CreateLegacySlideLayout(name, representative,
                    legacy.SlideWidth, legacy.SlideHeight);
                layoutPart.SlideLayout.Save();
                catalog.Add(representative, new LegacyPptLayoutTarget(masterIndex, layoutIndex));
            }
            catalog.AddMasterFallback(master.MasterId, new LegacyPptLayoutTarget(masterIndex, 0));
            masterPart.SlideMaster?.Save();
        }

        private static SlideLayoutPart AddLegacyLayoutPart(SlideMasterPart masterPart) {
            string relationshipId = GetNextRelationshipId(masterPart);
            SlideLayoutPart layoutPart = masterPart.AddNewPart<SlideLayoutPart>(relationshipId);
            layoutPart.AddPart(masterPart);
            SlideMaster slideMaster = masterPart.SlideMaster
                ?? throw new InvalidDataException("The projected PowerPoint package has no slide master.");
            SlideLayoutIdList layoutIds = slideMaster.SlideLayoutIdList ??= new SlideLayoutIdList();
            uint layoutId = layoutIds.Elements<SlideLayoutId>()
                .Select(item => item.Id?.Value ?? 2147483648U)
                .DefaultIfEmpty(2147483648U)
                .Max() + 1U;
            layoutIds.Append(new SlideLayoutId { Id = layoutId, RelationshipId = relationshipId });
            return layoutPart;
        }

        private static SlideLayout CreateLegacySlideLayout(string name,
            LegacyPptSlide source, int slideWidth, int slideHeight) {
            var layout = new SlideLayout(
                new CommonSlideData(CreateLegacyLayoutPlaceholderTree(source.Shapes,
                    source.HeaderFooter, slideWidth, slideHeight)) {
                    Name = name
                },
                new ColorMapOverride(new A.MasterColorMapping())) {
                    Type = MapLegacyLayoutType(source.Layout, source.LayoutPlaceholderTypes)
                };
            ApplyLegacyHeaderFooter(layout, layout.CommonSlideData,
                source.HeaderFooter, allowHeader: false);
            return layout;
        }

        private static ShapeTree CreateLegacyLayoutPlaceholderTree(
            IReadOnlyList<LegacyPptShape> shapes,
            LegacyPptHeaderFooterSettings? headerFooter,
            int slideWidth, int slideHeight) {
            var tree = new ShapeTree(
                new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                    new NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                PowerPointUtils.CreateDefaultGroupShapeProperties());
            uint shapeId = 2U;
            foreach (LegacyPptShape source in shapes.Where(shape => shape.Placeholder != null)) {
                tree.Append(CreateLegacyLayoutPlaceholderShape(source, shapeId++));
            }
            AppendLegacyHeaderFooterPlaceholders(tree, headerFooter, slideWidth,
                slideHeight, ref shapeId);
            return tree;
        }

        private static void AppendLegacyHeaderFooterPlaceholders(ShapeTree tree,
            LegacyPptHeaderFooterSettings? settings, int slideWidth,
            int slideHeight, ref uint shapeId) {
            if (settings == null) return;
            long width = ToEmus(slideWidth);
            long height = ToEmus(slideHeight);
            long margin = Math.Max(1L, width / 24L);
            long placeholderHeight = Math.Max(1L, height / 16L);
            long top = Math.Max(0L, height - margin - placeholderHeight);
            long sideWidth = Math.Max(1L, width / 5L);
            long centerWidth = Math.Max(1L, width / 3L);
            if (settings.ShowDate || settings.UserDateText.Length > 0) {
                AppendLegacyHeaderFooterPlaceholder(tree, PlaceholderValues.DateAndTime,
                    settings.UserDateText, margin, top, sideWidth, placeholderHeight,
                    ref shapeId);
            }
            if (settings.ShowFooter || settings.FooterText.Length > 0) {
                AppendLegacyHeaderFooterPlaceholder(tree, PlaceholderValues.Footer,
                    settings.FooterText, (width - centerWidth) / 2L, top,
                    centerWidth, placeholderHeight, ref shapeId);
            }
            if (settings.ShowSlideNumber) {
                AppendLegacyHeaderFooterPlaceholder(tree, PlaceholderValues.SlideNumber,
                    string.Empty, width - margin - sideWidth, top, sideWidth,
                    placeholderHeight, ref shapeId);
            }
        }

        private static void AppendLegacyHeaderFooterPlaceholder(ShapeTree tree,
            PlaceholderValues type, string text, long left, long top, long width,
            long height, ref uint shapeId) {
            bool exists = tree.Elements<Shape>().Any(shape => shape.NonVisualShapeProperties?
                .ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value == type);
            if (exists) return;
            string name = type == PlaceholderValues.DateAndTime
                ? "Binary Date Placeholder"
                : type == PlaceholderValues.Footer
                    ? "Binary Footer Placeholder"
                    : type == PlaceholderValues.SlideNumber
                        ? "Binary Slide Number Placeholder"
                        : "Binary Header/Footer Placeholder";
            tree.Append(new Shape(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties { Id = shapeId++, Name = name },
                    new NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties(
                        new PlaceholderShape { Type = type })),
                new ShapeProperties(new A.Transform2D(
                    new A.Offset { X = left, Y = top },
                    new A.Extents { Cx = width, Cy = height })),
                new TextBody(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(
                        new A.Run(new A.Text(text ?? string.Empty)),
                        new A.EndParagraphRunProperties()))));
        }

        private static Shape CreateLegacyLayoutPlaceholderShape(LegacyPptShape source,
            uint shapeId) {
            LegacyPptPlaceholder placeholder = source.Placeholder
                ?? throw new InvalidOperationException("A layout placeholder source has no PlaceholderAtom.");
            string name = source.Metadata.Name
                ?? $"Binary {placeholder.Kind} Placeholder {placeholder.Position}";
            var shape = new Shape(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties { Id = shapeId, Name = name },
                    new NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties(
                        CreateLegacyPlaceholderShape(placeholder))),
                new ShapeProperties(
                    new A.Transform2D(
                        new A.Offset {
                            X = ToEmus(source.Bounds.Left),
                            Y = ToEmus(source.Bounds.Top)
                        },
                        new A.Extents {
                            Cx = Math.Max(1L, ToEmus(source.Bounds.Width)),
                            Cy = Math.Max(1L, ToEmus(source.Bounds.Height))
                        })),
                new TextBody(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(new A.EndParagraphRunProperties())));
            if (source.Metadata.Description != null) {
                shape.NonVisualShapeProperties!.NonVisualDrawingProperties!.Description =
                    source.Metadata.Description;
            }
            return shape;
        }

        private static PlaceholderShape CreateLegacyPlaceholderShape(
            LegacyPptPlaceholder source) {
            PlaceholderValues? type = MapPlaceholder(source.Kind);
            var target = new PlaceholderShape {
                Index = checked((uint)source.Position),
                Size = source.Size switch {
                    LegacyPptPlaceholderSize.Full => PlaceholderSizeValues.Full,
                    LegacyPptPlaceholderSize.Half => PlaceholderSizeValues.Half,
                    LegacyPptPlaceholderSize.Quarter => PlaceholderSizeValues.Quarter,
                    _ => throw new ArgumentOutOfRangeException(nameof(source))
                }
            };
            if (type.HasValue) target.Type = type.Value;
            if (source.Kind == LegacyPptPlaceholderKind.VerticalTitle
                || source.Kind == LegacyPptPlaceholderKind.VerticalBody
                || source.Kind == LegacyPptPlaceholderKind.VerticalObject) {
                target.Orientation = DirectionValues.Vertical;
            }
            return target;
        }

        private static void ApplyLegacyPlaceholder(OpenXmlElement target,
            LegacyPptShape source) {
            if (source.Placeholder == null) return;
            ApplicationNonVisualDrawingProperties? applicationProperties = target switch {
                Shape shape => shape.NonVisualShapeProperties?
                    .ApplicationNonVisualDrawingProperties,
                ConnectionShape connector => connector.NonVisualConnectionShapeProperties?
                    .ApplicationNonVisualDrawingProperties,
                Picture picture => picture.NonVisualPictureProperties?
                    .ApplicationNonVisualDrawingProperties,
                GraphicFrame frame => frame.NonVisualGraphicFrameProperties?
                    .ApplicationNonVisualDrawingProperties,
                _ => null
            };
            if (applicationProperties == null) return;
            applicationProperties.PlaceholderShape = CreateLegacyPlaceholderShape(source.Placeholder);
        }

        private static SlideLayoutValues MapLegacyLayoutType(LegacyPptSlideLayoutType? type,
            IReadOnlyList<LegacyPptPlaceholderKind> placeholderTypes) => type switch {
                LegacyPptSlideLayoutType.TitleSlide => SlideLayoutValues.Title,
                LegacyPptSlideLayoutType.TitleBody => MapTitleBodyLayout(placeholderTypes),
                LegacyPptSlideLayoutType.MasterTitle => SlideLayoutValues.Title,
                LegacyPptSlideLayoutType.TitleOnly => SlideLayoutValues.TitleOnly,
                LegacyPptSlideLayoutType.TwoColumns => SlideLayoutValues.TwoColumnText,
                LegacyPptSlideLayoutType.TwoRows => SlideLayoutValues.TwoObjects,
                LegacyPptSlideLayoutType.ColumnTwoRows => SlideLayoutValues.ObjectAndTwoObjects,
                LegacyPptSlideLayoutType.TwoRowsColumn => SlideLayoutValues.TwoObjectsAndObject,
                LegacyPptSlideLayoutType.TwoColumnsRow => SlideLayoutValues.TwoObjectsOverText,
                LegacyPptSlideLayoutType.FourObjects => SlideLayoutValues.FourObjects,
                LegacyPptSlideLayoutType.BigObject => SlideLayoutValues.ObjectOnly,
                LegacyPptSlideLayoutType.Blank => SlideLayoutValues.Blank,
                LegacyPptSlideLayoutType.VerticalTitleBody => SlideLayoutValues.VerticalTitleAndText,
                LegacyPptSlideLayoutType.VerticalTwoRows =>
                    SlideLayoutValues.VerticalTitleAndTextOverChart,
                _ => SlideLayoutValues.Custom
            };

        private static SlideLayoutValues MapTitleBodyLayout(
            IReadOnlyList<LegacyPptPlaceholderKind> placeholderTypes) {
            LegacyPptPlaceholderKind content = placeholderTypes.Skip(1)
                .FirstOrDefault(value => value != LegacyPptPlaceholderKind.None);
            return content switch {
                LegacyPptPlaceholderKind.Table => SlideLayoutValues.Table,
                LegacyPptPlaceholderKind.Graph => SlideLayoutValues.Chart,
                LegacyPptPlaceholderKind.ClipArt => SlideLayoutValues.TextAndClipArt,
                LegacyPptPlaceholderKind.Media => SlideLayoutValues.TextAndMedia,
                LegacyPptPlaceholderKind.Object or LegacyPptPlaceholderKind.OrganizationChart
                    or LegacyPptPlaceholderKind.VerticalObject => SlideLayoutValues.TextAndObject,
                _ => SlideLayoutValues.Text
            };
        }

        private static string GetLegacySlideLayoutName(LegacyPptMaster master,
            LegacyPptSlide slide) {
            string type = slide.Layout?.ToString() ?? $"0x{slide.LayoutType:X8}";
            string signature = string.Join("-", slide.LayoutPlaceholderTypes
                .Select(value => ((byte)value).ToString("X2")));
            return $"{GetLegacyMasterName(master)} / {type} / {signature}";
        }

        private sealed class LegacyPptLayoutCatalog {
            private readonly Dictionary<string, LegacyPptLayoutTarget> _slides =
                new(StringComparer.Ordinal);
            private readonly Dictionary<uint, LegacyPptLayoutTarget> _fallbacks = new();

            internal void Add(LegacyPptSlide slide, LegacyPptLayoutTarget target) =>
                _slides[CreateKey(slide)] = target;

            internal void AddMasterFallback(uint masterId, LegacyPptLayoutTarget target) =>
                _fallbacks[masterId] = target;

            internal bool TryGet(LegacyPptSlide slide, out LegacyPptLayoutTarget target) =>
                _slides.TryGetValue(CreateKey(slide), out target)
                || _fallbacks.TryGetValue(slide.MasterId, out target);

            internal static string CreateKey(LegacyPptSlide slide) =>
                $"{slide.MasterId:X8}:{slide.LayoutType:X8}:"
                + string.Join("-", slide.LayoutPlaceholderTypes
                    .Select(value => ((byte)value).ToString("X2")))
                + ":" + (slide.HeaderFooter?.CreateLayoutKey() ?? string.Empty);
        }

        private readonly struct LegacyPptLayoutTarget {
            internal LegacyPptLayoutTarget(int masterIndex, int layoutIndex) {
                MasterIndex = masterIndex;
                LayoutIndex = layoutIndex;
            }

            internal int MasterIndex { get; }
            internal int LayoutIndex { get; }
        }
    }
}
