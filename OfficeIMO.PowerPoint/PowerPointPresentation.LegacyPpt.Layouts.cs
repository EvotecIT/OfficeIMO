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
                layoutPart.SlideLayout = CreateLegacySlideLayout(name, representative);
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
            LegacyPptSlide source) => new(
                new CommonSlideData(CreateLegacyLayoutPlaceholderTree(source.Shapes)) {
                    Name = name
                },
                new ColorMapOverride(new A.MasterColorMapping())) {
                    Type = MapLegacyLayoutType(source.Layout, source.LayoutPlaceholderTypes)
                };

        private static ShapeTree CreateLegacyLayoutPlaceholderTree(
            IReadOnlyList<LegacyPptShape> shapes) {
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
            return tree;
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
                    .Select(value => ((byte)value).ToString("X2")));
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
