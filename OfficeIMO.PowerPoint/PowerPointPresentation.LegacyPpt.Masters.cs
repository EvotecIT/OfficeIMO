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
            if (master.ColorScheme != null) {
                ApplyLegacyColorScheme(masterPart, master.ColorScheme);
                ApplyLegacyBackground(slideMaster.CommonSlideData, master.ColorScheme.Background);
            }
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
            if (!titleMaster.FollowsMasterObjects && layoutPart.SlideLayout.CommonSlideData != null) {
                SetShowsMasterShapes(layoutPart.SlideLayout.CommonSlideData, false);
            }
            if (!titleMaster.FollowsMasterColorScheme && titleMaster.ColorScheme != null) {
                ApplyLegacyColorScheme(layoutPart, titleMaster.ColorScheme);
            }
            if (!titleMaster.FollowsMasterBackground && titleMaster.ColorScheme != null
                && layoutPart.SlideLayout.CommonSlideData != null) {
                ApplyLegacyBackground(layoutPart.SlideLayout.CommonSlideData,
                    titleMaster.ColorScheme.Background);
            }

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

        private static void ProjectLegacySlideDesign(PowerPointSlide slide, LegacyPptSlide source) {
            Slide slideRoot = slide.SlidePart.Slide ??= new Slide();
            CommonSlideData commonSlideData = slideRoot.CommonSlideData ??= new CommonSlideData();
            if (!source.FollowsMasterObjects) SetShowsMasterShapes(commonSlideData, false);
            if (!source.FollowsMasterColorScheme && source.ColorScheme != null) {
                ApplyLegacyColorScheme(slide.SlidePart, source.ColorScheme);
            }
            if (!source.FollowsMasterBackground && source.ColorScheme != null) {
                ApplyLegacyBackground(commonSlideData, source.ColorScheme.Background);
            }
        }

        private static void SetShowsMasterShapes(CommonSlideData commonSlideData, bool value) {
            commonSlideData.SetAttribute(new OpenXmlAttribute(string.Empty, "showMasterSp", string.Empty,
                value ? "1" : "0"));
        }

        private static void ApplyLegacyBackground(CommonSlideData commonSlideData, string color) {
            commonSlideData.Background = new Background(
                new BackgroundProperties(
                    new A.SolidFill(new A.RgbColorModelHex { Val = color })));
        }

        private static void ApplyLegacyColorScheme(SlideMasterPart masterPart, LegacyPptColorScheme source) {
            A.ColorScheme target = EnsureColorScheme(masterPart);
            SetLegacyThemeColors(target, source);
            masterPart.ThemePart?.Theme?.Save();
        }

        private static void ApplyLegacyColorScheme(SlideLayoutPart layoutPart, LegacyPptColorScheme source) {
            A.ColorScheme target = CloneMasterColorScheme(layoutPart.SlideMasterPart);
            SetLegacyThemeColors(target, source);
            ThemeOverridePart overridePart = layoutPart.ThemeOverridePart
                ?? layoutPart.AddNewPart<ThemeOverridePart>();
            overridePart.ThemeOverride = new A.ThemeOverride(target);
            overridePart.ThemeOverride.Save();
        }

        private static void ApplyLegacyColorScheme(SlidePart slidePart, LegacyPptColorScheme source) {
            A.ColorScheme target = CloneMasterColorScheme(slidePart.SlideLayoutPart?.SlideMasterPart);
            SetLegacyThemeColors(target, source);
            ThemeOverridePart overridePart = slidePart.ThemeOverridePart
                ?? slidePart.AddNewPart<ThemeOverridePart>();
            overridePart.ThemeOverride = new A.ThemeOverride(target);
            overridePart.ThemeOverride.Save();
        }

        private static A.ColorScheme CloneMasterColorScheme(SlideMasterPart? masterPart) {
            A.ColorScheme? source = masterPart?.ThemePart?.Theme?.ThemeElements?.ColorScheme;
            return source?.CloneNode(true) as A.ColorScheme
                ?? new A.ColorScheme { Name = "Binary PowerPoint" };
        }

        private static void SetLegacyThemeColors(A.ColorScheme target, LegacyPptColorScheme source) {
            SetThemeColor(target, PowerPointThemeColor.Light1, source.Background);
            SetThemeColor(target, PowerPointThemeColor.Dark1, source.Text);
            SetThemeColor(target, PowerPointThemeColor.Accent4, source.Shadow);
            SetThemeColor(target, PowerPointThemeColor.Dark2, source.TitleText);
            SetThemeColor(target, PowerPointThemeColor.Light2, source.Fill);
            SetThemeColor(target, PowerPointThemeColor.Accent1, source.Accent1);
            SetThemeColor(target, PowerPointThemeColor.Accent2, source.Accent2);
            SetThemeColor(target, PowerPointThemeColor.Accent3, source.Accent3);
        }

        private static void SetThemeColor(A.ColorScheme scheme, PowerPointThemeColor color, string value) {
            OpenXmlCompositeElement element = GetOrCreateColorElement(scheme, color);
            element.RemoveAllChildren<A.RgbColorModelHex>();
            element.RemoveAllChildren<A.SystemColor>();
            element.Append(new A.RgbColorModelHex { Val = value });
        }

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
            ApplyLegacyShapeStyle(shapeProperties, source);

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

        internal static void ApplyLegacyShapeStyle(ShapeProperties properties, LegacyPptShape source) {
            OfficeIMO.Drawing.Binary.OfficeArtShapeStyle style = source.Style;
            if (style.FillEnabled == false) {
                SetLegacyShapeFill(properties, new A.NoFill());
            } else if (source.FillColor != null && style.FillType.GetValueOrDefault() == 0) {
                SetLegacyShapeFill(properties, CreateLegacySolidFill(source.FillColor, style.FillOpacity));
            }

            bool hasLineStyle = style.LineEnabled.HasValue || source.LineColor != null
                || style.LineOpacity.HasValue || style.LineWidthEmus.HasValue || style.LineDashing.HasValue
                || style.LineStartArrowhead.HasValue || style.LineEndArrowhead.HasValue
                || style.LineJoinStyle.HasValue || style.LineEndCapStyle.HasValue;
            if (!hasLineStyle) return;

            A.Outline outline = properties.GetFirstChild<A.Outline>() ?? new A.Outline();
            if (outline.Parent == null) properties.Append(outline);
            if (style.LineEnabled == false) {
                SetLegacyOutlineFill(outline, new A.NoFill());
                return;
            }

            if (source.LineColor != null) {
                SetLegacyOutlineFill(outline, CreateLegacySolidFill(source.LineColor, style.LineOpacity));
            }
            if (style.LineWidthEmus is >= 0) outline.Width = style.LineWidthEmus.Value;
            ApplyLegacyLineDash(outline, style.LineDashing);
            ApplyLegacyLineJoin(outline, style.LineJoinStyle);
            outline.CapType = MapLegacyLineCap(style.LineEndCapStyle);
            ApplyLegacyLineEnd(outline, isHead: true, style.LineStartArrowhead,
                style.LineStartArrowWidth, style.LineStartArrowLength);
            ApplyLegacyLineEnd(outline, isHead: false, style.LineEndArrowhead,
                style.LineEndArrowWidth, style.LineEndArrowLength);
        }

        private static A.SolidFill CreateLegacySolidFill(string color, double? opacity) {
            var rgb = new A.RgbColorModelHex { Val = color };
            if (opacity.HasValue) {
                rgb.Append(new A.Alpha { Val = checked((int)Math.Round(
                    Math.Max(0D, Math.Min(1D, opacity.Value)) * 100000D)) });
            }
            return new A.SolidFill(rgb);
        }

        private static void SetLegacyShapeFill(ShapeProperties properties, OpenXmlElement fill) {
            properties.RemoveAllChildren<A.NoFill>();
            properties.RemoveAllChildren<A.SolidFill>();
            properties.RemoveAllChildren<A.GradientFill>();
            properties.RemoveAllChildren<A.BlipFill>();
            properties.RemoveAllChildren<A.PatternFill>();
            A.Outline? outline = properties.GetFirstChild<A.Outline>();
            if (outline != null) properties.InsertBefore(fill, outline);
            else properties.Append(fill);
        }

        private static void SetLegacyOutlineFill(A.Outline outline, OpenXmlElement fill) {
            outline.RemoveAllChildren<A.NoFill>();
            outline.RemoveAllChildren<A.SolidFill>();
            outline.RemoveAllChildren<A.GradientFill>();
            outline.RemoveAllChildren<A.PatternFill>();
            OpenXmlElement? first = outline.FirstChild;
            if (first != null) outline.InsertBefore(fill, first);
            else outline.Append(fill);
        }

        private static void ApplyLegacyLineDash(A.Outline outline, uint? value) {
            outline.RemoveAllChildren<A.PresetDash>();
            A.PresetLineDashValues? dash = value switch {
                0 => A.PresetLineDashValues.Solid,
                1 => A.PresetLineDashValues.SystemDash,
                2 => A.PresetLineDashValues.SystemDot,
                3 => A.PresetLineDashValues.SystemDashDot,
                4 => A.PresetLineDashValues.SystemDashDotDot,
                5 => A.PresetLineDashValues.Dot,
                6 => A.PresetLineDashValues.Dash,
                7 => A.PresetLineDashValues.LargeDash,
                8 => A.PresetLineDashValues.DashDot,
                9 => A.PresetLineDashValues.LargeDashDot,
                10 => A.PresetLineDashValues.LargeDashDotDot,
                _ => null
            };
            if (dash.HasValue) InsertLegacyOutlineChild(outline, new A.PresetDash { Val = dash.Value });
        }

        private static void ApplyLegacyLineJoin(A.Outline outline, uint? value) {
            outline.RemoveAllChildren<A.Bevel>();
            outline.RemoveAllChildren<A.Miter>();
            outline.RemoveAllChildren<A.Round>();
            OpenXmlElement? join = value switch {
                0 => new A.Bevel(),
                1 => new A.Miter(),
                2 => new A.Round(),
                _ => null
            };
            if (join != null) InsertLegacyOutlineChild(outline, join);
        }

        private static A.LineCapValues? MapLegacyLineCap(uint? value) => value switch {
            0 => A.LineCapValues.Round,
            1 => A.LineCapValues.Square,
            2 => A.LineCapValues.Flat,
            _ => null
        };

        private static void ApplyLegacyLineEnd(A.Outline outline, bool isHead, uint? type,
            uint? width, uint? length) {
            if (!type.HasValue && !width.HasValue && !length.HasValue) return;
            if (isHead) {
                A.HeadEnd head = outline.GetFirstChild<A.HeadEnd>() ?? new A.HeadEnd();
                head.Type = MapLegacyLineEnd(type) ?? A.LineEndValues.None;
                head.Width = MapLegacyLineEndWidth(width);
                head.Length = MapLegacyLineEndLength(length);
                if (head.Parent == null) InsertLegacyOutlineChild(outline, head);
            } else {
                A.TailEnd tail = outline.GetFirstChild<A.TailEnd>() ?? new A.TailEnd();
                tail.Type = MapLegacyLineEnd(type) ?? A.LineEndValues.None;
                tail.Width = MapLegacyLineEndWidth(width);
                tail.Length = MapLegacyLineEndLength(length);
                if (tail.Parent == null) InsertLegacyOutlineChild(outline, tail);
            }
        }

        private static A.LineEndValues? MapLegacyLineEnd(uint? value) => value switch {
            0 => A.LineEndValues.None,
            1 => A.LineEndValues.Triangle,
            2 => A.LineEndValues.Stealth,
            3 => A.LineEndValues.Diamond,
            4 => A.LineEndValues.Oval,
            5 => A.LineEndValues.Arrow,
            _ => null
        };

        private static A.LineEndWidthValues? MapLegacyLineEndWidth(uint? value) => value switch {
            0 => A.LineEndWidthValues.Small,
            1 => A.LineEndWidthValues.Medium,
            2 => A.LineEndWidthValues.Large,
            _ => null
        };

        private static A.LineEndLengthValues? MapLegacyLineEndLength(uint? value) => value switch {
            0 => A.LineEndLengthValues.Small,
            1 => A.LineEndLengthValues.Medium,
            2 => A.LineEndLengthValues.Large,
            _ => null
        };

        private static void InsertLegacyOutlineChild(A.Outline outline, OpenXmlElement child) {
            int order = child switch {
                A.NoFill or A.SolidFill or A.GradientFill or A.PatternFill => 0,
                A.PresetDash or A.CustomDash => 1,
                A.Round or A.Bevel or A.Miter => 2,
                A.HeadEnd => 3,
                A.TailEnd => 4,
                _ => 100
            };
            OpenXmlElement? before = outline.ChildElements.FirstOrDefault(existing => (existing switch {
                A.NoFill or A.SolidFill or A.GradientFill or A.PatternFill => 0,
                A.PresetDash or A.CustomDash => 1,
                A.Round or A.Bevel or A.Miter => 2,
                A.HeadEnd => 3,
                A.TailEnd => 4,
                _ => 100
            }) > order);
            if (before != null) outline.InsertBefore(child, before);
            else outline.Append(child);
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
