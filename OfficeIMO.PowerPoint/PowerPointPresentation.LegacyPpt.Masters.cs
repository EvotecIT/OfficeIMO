using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        private static IReadOnlyDictionary<uint, LegacyPptLayoutTarget> ProjectLegacyMasters(
            PowerPointPresentation presentation, LegacyPptPresentation legacy) {
            LegacyPptMaster[] mainMasters = legacy.Masters.Where(master => master.IsMainMaster).ToArray();
            if (mainMasters.Length == 0) {
                return new Dictionary<uint, LegacyPptLayoutTarget>();
            }

            SlideMasterPart firstMasterPart = presentation._presentationPart.SlideMasterParts.First();
            ResetLegacyMasterScaffold(firstMasterPart, "Binary Main Master 1");
            var masterParts = new List<SlideMasterPart> { firstMasterPart };
            for (int index = 1; index < mainMasters.Length; index++) {
                masterParts.Add(presentation.CloneSlideMasterPart(firstMasterPart, out _));
            }

            var targets = new Dictionary<uint, LegacyPptLayoutTarget>();
            for (int index = 0; index < mainMasters.Length; index++) {
                LegacyPptMaster mainMaster = mainMasters[index];
                SlideMasterPart masterPart = masterParts[index];
                ApplyLegacyMaster(masterPart, mainMaster);
                targets.Add(mainMaster.MasterId, new LegacyPptLayoutTarget(index, 0));
            }

            foreach (LegacyPptMaster titleMaster in legacy.Masters.Where(master => !master.IsMainMaster)) {
                LegacyPptLayoutTarget parentTarget;
                if (!targets.TryGetValue(titleMaster.ParentMasterId, out parentTarget)) {
                    parentTarget = new LegacyPptLayoutTarget(0, 0);
                }
                SlideMasterPart parentPart = masterParts[parentTarget.MasterIndex];
                int layoutIndex = AddLegacyTitleMasterLayout(parentPart, titleMaster);
                targets[titleMaster.MasterId] = new LegacyPptLayoutTarget(parentTarget.MasterIndex, layoutIndex);
            }

            return targets;
        }

        private static void ResetLegacyMasterScaffold(SlideMasterPart masterPart, string name) {
            SlideLayoutPart[] layouts = masterPart.SlideLayoutParts.ToArray();
            if (layouts.Length == 0) {
                throw new InvalidDataException("The native PowerPoint scaffold has no slide layout.");
            }
            SlideLayoutPart blankLayout = layouts[0];
            foreach (SlideLayoutPart extraLayout in layouts.Skip(1)) {
                masterPart.DeletePart(extraLayout);
            }

            blankLayout.SlideLayout = CreateLegacyLayout("Binary Main Master", SlideLayoutValues.Blank,
                Array.Empty<LegacyPptShape>());
            SlideMaster slideMaster = masterPart.SlideMaster
                ?? throw new InvalidDataException("The native PowerPoint scaffold has no slide master.");
            slideMaster.CommonSlideData = new CommonSlideData(CreateLegacyShapeTree()) { Name = name };
            string relationshipId = masterPart.GetIdOfPart(blankLayout);
            slideMaster.SlideLayoutIdList = new SlideLayoutIdList(new SlideLayoutId {
                Id = 2147483649U,
                RelationshipId = relationshipId
            });
            blankLayout.SlideLayout.Save();
            slideMaster.Save();
        }

        private static void ApplyLegacyMaster(SlideMasterPart masterPart, LegacyPptMaster master) {
            SlideMaster slideMaster = masterPart.SlideMaster
                ?? throw new InvalidDataException("The projected PowerPoint package has no slide master.");
            slideMaster.CommonSlideData = new CommonSlideData(CreateLegacyShapeTree(master.Shapes)) {
                Name = GetLegacyMasterName(master)
            };
            SlideLayoutPart blankLayout = masterPart.SlideLayoutParts.First();
            if (blankLayout.SlideLayout?.CommonSlideData != null) {
                blankLayout.SlideLayout.CommonSlideData.Name = GetLegacyMasterName(master);
                blankLayout.SlideLayout.Save();
            }
            slideMaster.Save();
        }

        private static int AddLegacyTitleMasterLayout(SlideMasterPart masterPart, LegacyPptMaster titleMaster) {
            string relationshipId = GetNextRelationshipId(masterPart);
            SlideLayoutPart layoutPart = masterPart.AddNewPart<SlideLayoutPart>(relationshipId);
            layoutPart.SlideLayout = CreateLegacyLayout(GetLegacyMasterName(titleMaster),
                SlideLayoutValues.Title, titleMaster.Shapes);
            layoutPart.AddPart(masterPart);

            SlideMaster slideMaster = masterPart.SlideMaster
                ?? throw new InvalidDataException("The projected PowerPoint package has no slide master.");
            SlideLayoutIdList layoutIds = slideMaster.SlideLayoutIdList ??= new SlideLayoutIdList();
            uint layoutId = layoutIds.Elements<SlideLayoutId>()
                .Select(item => item.Id?.Value ?? 2147483648U)
                .DefaultIfEmpty(2147483648U)
                .Max() + 1U;
            layoutIds.Append(new SlideLayoutId { Id = layoutId, RelationshipId = relationshipId });
            layoutPart.SlideLayout.Save();
            slideMaster.Save();
            return masterPart.SlideLayoutParts.Count() - 1;
        }

        private static string GetLegacyMasterName(LegacyPptMaster master) =>
            $"Binary {(master.IsMainMaster ? "Main" : "Title")} Master {master.MasterId:X8}";

        private static SlideLayout CreateLegacyLayout(string name, SlideLayoutValues type,
            IReadOnlyList<LegacyPptShape> shapes) => new(
                new CommonSlideData(CreateLegacyShapeTree(shapes)) { Name = name },
                new ColorMapOverride(new A.MasterColorMapping())) { Type = type };

        private static ShapeTree CreateLegacyShapeTree(IReadOnlyList<LegacyPptShape>? shapes = null) {
            var tree = new ShapeTree(
                new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                    new NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                PowerPointUtils.CreateDefaultGroupShapeProperties());
            if (shapes == null) return tree;

            uint shapeId = 2U;
            foreach (LegacyPptShape source in shapes) {
                Shape? shape = CreateLegacyShape(source, shapeId);
                if (shape == null) continue;
                tree.Append(shape);
                shapeId++;
            }
            return tree;
        }

        private static Shape? CreateLegacyShape(LegacyPptShape source, uint shapeId) {
            A.ShapeTypeValues geometry;
            switch (source.Kind) {
                case LegacyPptShapeKind.TextBox:
                case LegacyPptShapeKind.Rectangle:
                    geometry = A.ShapeTypeValues.Rectangle;
                    break;
                case LegacyPptShapeKind.Ellipse:
                    geometry = A.ShapeTypeValues.Ellipse;
                    break;
                case LegacyPptShapeKind.Line:
                    geometry = A.ShapeTypeValues.Line;
                    break;
                default:
                    return null;
            }

            var applicationProperties = new ApplicationNonVisualDrawingProperties();
            PlaceholderValues? placeholder = MapPlaceholder(source.PlaceholderKind);
            if (placeholder.HasValue) {
                applicationProperties.Append(new PlaceholderShape { Type = placeholder.Value });
            }
            var shapeProperties = new ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = ToEmus(source.Bounds.Left), Y = ToEmus(source.Bounds.Top) },
                    new A.Extents {
                        Cx = Math.Max(1L, ToEmus(source.Bounds.Width)),
                        Cy = Math.Max(1L, ToEmus(source.Bounds.Height))
                    }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = geometry });
            if (geometry == A.ShapeTypeValues.Line) shapeProperties.Append(new A.NoFill());

            var shape = new Shape(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties { Id = shapeId, Name = $"Binary Shape {shapeId - 1U}" },
                    new NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                    applicationProperties),
                shapeProperties);
            if (source.Kind == LegacyPptShapeKind.TextBox) {
                shape.Append(new TextBody(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(new A.Run(new A.Text(source.Text)))));
            }
            return shape;
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
