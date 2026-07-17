using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using System.Runtime.CompilerServices;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        private static readonly ConditionalWeakTable<OpenXmlPart,
            Dictionary<ushort, string>> LegacyPictureBulletRelationships =
            new();

        private static LegacyPptLayoutCatalog ProjectLegacyMasters(
            PowerPointPresentation presentation, LegacyPptPresentation legacy,
            LegacyPptSoundProjectionContext soundContext,
            ICollection<LegacyPptDeferredProjection>
                deferredInteractions) {
            LegacyPptMaster[] mainMasters = legacy.Masters.Where(master => master.IsMainMaster).ToArray();
            var catalog = new LegacyPptLayoutCatalog();
            if (mainMasters.Length == 0) {
                return catalog;
            }

            SlideMasterPart firstMasterPart = presentation._presentationPart.SlideMasterParts.First();
            ResetLegacyMasterScaffold(firstMasterPart, "Binary Main Master 1");
            var masterParts = new List<SlideMasterPart> { firstMasterPart };
            for (int index = 1; index < mainMasters.Length; index++) {
                masterParts.Add(presentation.CloneSlideMasterPart(firstMasterPart, out _));
            }

            var masterTargets = new Dictionary<uint, LegacyPptLayoutTarget>();
            for (int index = 0; index < mainMasters.Length; index++) {
                LegacyPptMaster mainMaster = mainMasters[index];
                SlideMasterPart masterPart = masterParts[index];
                ApplyLegacyMaster(masterPart, mainMaster,
                    GetEffectiveLegacySlideHeaderFooter(legacy, mainMaster),
                    soundContext, deferredInteractions);
                var target = new LegacyPptLayoutTarget(index, 0);
                masterTargets.Add(mainMaster.MasterId, target);
                ProjectLegacyMainMasterLayouts(catalog, legacy, mainMaster,
                    masterPart, index);
            }

            foreach (LegacyPptMaster titleMaster in legacy.Masters.Where(master => !master.IsMainMaster)) {
                LegacyPptLayoutTarget parentTarget;
                if (!masterTargets.TryGetValue(titleMaster.ParentMasterId, out parentTarget)) {
                    parentTarget = new LegacyPptLayoutTarget(0, 0);
                }
                SlideMasterPart parentPart = masterParts[parentTarget.MasterIndex];
                int layoutIndex = AddLegacyTitleMasterLayout(parentPart, titleMaster,
                    soundContext, deferredInteractions);
                var titleTarget = new LegacyPptLayoutTarget(parentTarget.MasterIndex, layoutIndex);
                masterTargets[titleMaster.MasterId] = titleTarget;
                catalog.AddMasterFallback(titleMaster.MasterId, titleTarget);
            }

            return catalog;
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

            blankLayout.SlideLayout = CreateLegacyLayout(blankLayout, "Binary Main Master", SlideLayoutValues.Blank,
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

        private static void ApplyLegacyMaster(SlideMasterPart masterPart, LegacyPptMaster master,
            LegacyPptHeaderFooterSettings? headerFooter,
            LegacyPptSoundProjectionContext soundContext,
            ICollection<LegacyPptDeferredProjection>
                deferredInteractions) {
            SlideMaster slideMaster = masterPart.SlideMaster
                ?? throw new InvalidDataException("The projected PowerPoint package has no slide master.");
            slideMaster.CommonSlideData = new CommonSlideData(CreateLegacyShapeTree(masterPart, master.Shapes,
                master.ConnectorRules, soundContext: soundContext,
                deferredInteractions: deferredInteractions)) {
                Name = GetLegacyMasterName(master)
            };
            ApplyLegacyRoundTripTheme(masterPart, master.RoundTripTheme);
            if (master.ColorScheme != null
                && master.RoundTripTheme?.ThemeXml == null) {
                ApplyLegacyColorScheme(masterPart, master.ColorScheme);
            }
            if (master.Background != null) {
                ApplyLegacyBackground(masterPart, slideMaster.CommonSlideData,
                    master.Background);
            } else if (master.ColorScheme != null) {
                ApplyLegacyBackground(slideMaster.CommonSlideData, master.ColorScheme.Background);
            }
            ApplyLegacyMasterTextStyles(masterPart, slideMaster,
                master.TextMasterStyles);
            ApplyLegacyHeaderFooter(slideMaster, slideMaster.CommonSlideData,
                headerFooter, allowHeader: false);
            SlideLayoutPart blankLayout = masterPart.SlideLayoutParts.First();
            if (blankLayout.SlideLayout?.CommonSlideData != null) {
                blankLayout.SlideLayout.CommonSlideData.Name = GetLegacyMasterName(master);
                blankLayout.SlideLayout.Save();
            }
            slideMaster.Save();
        }

        private static int AddLegacyTitleMasterLayout(SlideMasterPart masterPart,
            LegacyPptMaster titleMaster,
            LegacyPptSoundProjectionContext soundContext,
            ICollection<LegacyPptDeferredProjection>
                deferredInteractions) {
            string relationshipId = GetNextRelationshipId(masterPart);
            SlideLayoutPart layoutPart = masterPart.AddNewPart<SlideLayoutPart>(relationshipId);
            layoutPart.SlideLayout = CreateLegacyLayout(layoutPart, GetLegacyMasterName(titleMaster),
                SlideLayoutValues.Title, titleMaster.Shapes,
                titleMaster.ConnectorRules, soundContext,
                deferredInteractions);
            layoutPart.AddPart(masterPart);
            if (!titleMaster.FollowsMasterObjects) {
                layoutPart.SlideLayout.ShowMasterShapes = false;
            }
            ApplyLegacyRoundTripTheme(layoutPart, titleMaster.RoundTripTheme);
            if (!titleMaster.FollowsMasterColorScheme
                && titleMaster.ColorScheme != null
                && titleMaster.RoundTripTheme?.ThemeXml == null) {
                ApplyLegacyColorScheme(layoutPart, titleMaster.ColorScheme);
            }
            if (!titleMaster.FollowsMasterBackground
                && layoutPart.SlideLayout.CommonSlideData != null) {
                if (titleMaster.Background != null) {
                    ApplyLegacyBackground(layoutPart,
                        layoutPart.SlideLayout.CommonSlideData, titleMaster.Background);
                } else if (titleMaster.ColorScheme != null) {
                    ApplyLegacyBackground(layoutPart.SlideLayout.CommonSlideData,
                        titleMaster.ColorScheme.Background);
                }
            }
            ApplyLegacyHeaderFooter(layoutPart.SlideLayout,
                layoutPart.SlideLayout.CommonSlideData, titleMaster.HeaderFooter,
                allowHeader: false);

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
            ApplyLegacyRoundTripTheme(slide.SlidePart,
                source.RoundTripTheme);
            if (!source.FollowsMasterObjects) slideRoot.ShowMasterShapes = false;
            if (!source.FollowsMasterColorScheme && source.ColorScheme != null
                && source.RoundTripTheme?.ThemeXml == null) {
                ApplyLegacyColorScheme(slide.SlidePart, source.ColorScheme);
            }
            if (!source.FollowsMasterBackground) {
                if (source.Background != null) {
                    ApplyLegacyBackground(slide.SlidePart, commonSlideData,
                        source.Background);
                } else if (source.ColorScheme != null) {
                    ApplyLegacyBackground(commonSlideData, source.ColorScheme.Background);
                }
            }
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

        private static SlideLayout CreateLegacyLayout(OpenXmlPart ownerPart, string name, SlideLayoutValues type,
            IReadOnlyList<LegacyPptShape> shapes,
            IReadOnlyList<LegacyPptConnectorRule>? connectorRules = null,
            LegacyPptSoundProjectionContext? soundContext = null,
            ICollection<LegacyPptDeferredProjection>?
                deferredInteractions = null) => new(
            new CommonSlideData(CreateLegacyShapeTree(ownerPart, shapes,
                connectorRules, soundContext: soundContext,
                deferredInteractions: deferredInteractions)) { Name = name },
                new ColorMapOverride(new A.MasterColorMapping())) { Type = type };

        private static ShapeTree CreateLegacyShapeTree(OpenXmlPart? ownerPart = null,
            IReadOnlyList<LegacyPptShape>? shapes = null,
            IReadOnlyList<LegacyPptConnectorRule>? connectorRules = null,
            IReadOnlyDictionary<uint, SlidePart>? slidePartsByLegacyId = null,
            LegacyPptSoundProjectionContext? soundContext = null,
            ICollection<LegacyPptDeferredProjection>?
                deferredInteractions = null) {
            var tree = new ShapeTree(
                new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                    new NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                PowerPointUtils.CreateDefaultGroupShapeProperties());
            if (shapes == null) return tree;
            if (ownerPart == null) {
                throw new InvalidOperationException("An owning Open XML part is required to project legacy shapes.");
            }

            uint nextShapeId = 2U;
            var projectedShapeIds = new Dictionary<uint, uint>();
            foreach (LegacyPptShape source in shapes) {
                OpenXmlElement? shape = CreateLegacyOpenXmlShape(ownerPart, source,
                    ref nextShapeId, slidePartsByLegacyId, soundContext,
                    projectedShapeIds, deferredInteractions);
                if (shape == null) continue;
                tree.Append(shape);
            }
            if (connectorRules != null) {
                ProjectLegacyConnectorRules(tree, connectorRules, projectedShapeIds);
            }
            return tree;
        }

        internal static OpenXmlElement? CreateLegacyOpenXmlShape(OpenXmlPart ownerPart,
            LegacyPptShape source, ref uint nextShapeId,
            IReadOnlyDictionary<uint, SlidePart>? slidePartsByLegacyId = null,
            LegacyPptSoundProjectionContext? soundContext = null,
            IDictionary<uint, uint>? projectedShapeIds = null,
            ICollection<LegacyPptDeferredProjection>?
                deferredInteractions = null) {
            if (ownerPart == null) throw new ArgumentNullException(nameof(ownerPart));
            if (source == null) throw new ArgumentNullException(nameof(source));
            uint shapeId = nextShapeId++;
            OpenXmlElement? projected = source.Kind switch {
                LegacyPptShapeKind.Picture => CreateLegacyPicture(ownerPart, source, shapeId),
                LegacyPptShapeKind.Table => CreateLegacyTableFrame(ownerPart,
                    source, shapeId, slidePartsByLegacyId, soundContext,
                    deferredInteractions),
                LegacyPptShapeKind.Group => CreateLegacyGroupShape(ownerPart, source, shapeId,
                    ref nextShapeId, slidePartsByLegacyId, soundContext,
                    projectedShapeIds, deferredInteractions),
                _ => CreateLegacyShape(ownerPart, source, shapeId,
                    slidePartsByLegacyId, soundContext,
                    deferredInteractions)
            };
            if (projected != null) {
                RegisterLegacyShapeId(source, projected,
                    projectedShapeIds);
                ApplyLegacyShapeMetadata(projected, source);
                ApplyLegacyShapeInteractions(ownerPart, projected, source,
                    slidePartsByLegacyId, soundContext,
                    deferredInteractions);
            }
            return projected;
        }

        internal static bool CanCreateLegacyOpenXmlShape(
            LegacyPptShape source) {
            if (source == null) throw new ArgumentNullException(nameof(source));
            return source.Kind switch {
                LegacyPptShapeKind.TextBox or
                LegacyPptShapeKind.Rectangle or
                LegacyPptShapeKind.Ellipse or
                LegacyPptShapeKind.Line => true,
                LegacyPptShapeKind.AutoShape or
                LegacyPptShapeKind.Connector =>
                    LegacyPptShapeGeometryMapper.TryGetPreset(
                        source.OfficeArtShapeType, out _),
                LegacyPptShapeKind.Picture =>
                    source.Picture?.HasImportableImage == true
                    && source.Picture.ContentType != null,
                LegacyPptShapeKind.Table => source.Table != null,
                LegacyPptShapeKind.Group =>
                    source.GroupCoordinateBounds.HasValue
                    && source.Children.Count > 0,
                _ => false
            };
        }

        private static void ApplyLegacyShapeMetadata(OpenXmlElement target, LegacyPptShape source) {
            NonVisualDrawingProperties? properties = target switch {
                Shape shape => shape.NonVisualShapeProperties?.NonVisualDrawingProperties,
                ConnectionShape connector => connector.NonVisualConnectionShapeProperties?
                    .NonVisualDrawingProperties,
                Picture picture => picture.NonVisualPictureProperties?.NonVisualDrawingProperties,
                GroupShape group => group.NonVisualGroupShapeProperties?.NonVisualDrawingProperties,
                GraphicFrame frame => frame.NonVisualGraphicFrameProperties?
                    .NonVisualDrawingProperties,
                _ => null
            };
            if (properties == null) return;
            if (source.Metadata.Name != null) properties.Name = source.Metadata.Name;
            if (source.Metadata.Description != null) properties.Description = source.Metadata.Description;
            properties.Hidden = source.Style.Hidden == true ? true : null;
        }

        private static OpenXmlElement? CreateLegacyShape(OpenXmlPart ownerPart,
            LegacyPptShape source, uint shapeId,
            IReadOnlyDictionary<uint, SlidePart>? slidePartsByLegacyId,
            LegacyPptSoundProjectionContext? soundContext,
            ICollection<LegacyPptDeferredProjection>?
                deferredInteractions) {
            if (source.Kind == LegacyPptShapeKind.Connector
                && LegacyPptShapeGeometryMapper.TryGetPreset(source.OfficeArtShapeType,
                    out A.ShapeTypeValues connectorGeometry)) {
                return CreateLegacyConnectionShape(source, shapeId, connectorGeometry);
            }
            A.ShapeTypeValues geometry;
            switch (source.Kind) {
                case LegacyPptShapeKind.TextBox:
                    geometry = source.OfficeArtShapeType != 202
                        && LegacyPptShapeGeometryMapper.TryGetPreset(source.OfficeArtShapeType,
                            out A.ShapeTypeValues textGeometry)
                        ? textGeometry
                        : A.ShapeTypeValues.Rectangle;
                    break;
                case LegacyPptShapeKind.Rectangle:
                    geometry = A.ShapeTypeValues.Rectangle;
                    break;
                case LegacyPptShapeKind.Ellipse:
                    geometry = A.ShapeTypeValues.Ellipse;
                    break;
                case LegacyPptShapeKind.Line:
                    geometry = A.ShapeTypeValues.Line;
                    break;
                case LegacyPptShapeKind.AutoShape:
                    if (!LegacyPptShapeGeometryMapper.TryGetPreset(source.OfficeArtShapeType,
                            out geometry)) return null;
                    break;
                default:
                    return null;
            }

            var applicationProperties = new ApplicationNonVisualDrawingProperties();
            if (source.Placeholder != null) {
                applicationProperties.Append(CreateLegacyPlaceholderShape(source.Placeholder));
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
            LegacyPptShapeGeometryMapper.ApplyExactPresetAdjustments(source.OfficeArtShapeType,
                source.Geometry, shapeProperties.GetFirstChild<A.PresetGeometry>()!);
            ApplyLegacyShapeStyle(shapeProperties, source);
            ApplyLegacyShapeTransform(shapeProperties.Transform2D!, source);

            var shape = new Shape(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties { Id = shapeId, Name = $"Binary Shape {shapeId - 1U}" },
                    new NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                    applicationProperties),
                shapeProperties);
            if (source.Kind == LegacyPptShapeKind.TextBox) {
                bool hasRichText = source.TextBody.HasExplicitCharacterFormatting
                    || source.TextBody.HasParagraphFormatting
                    || source.TextBody.HasInteractions;
                bool deferTextInteractions =
                    ShouldDeferLegacyTextInteractions(source.TextBody,
                        slidePartsByLegacyId, deferredInteractions);
                TextBody textBody = hasRichText
                    ? LegacyPptTextProjection.CreateTextBody(source.TextBody,
                        source.TextFrame,
                        interaction => deferTextInteractions
                            ? Array.Empty<OpenXmlElement>()
                            : ProjectLegacyInteraction(ownerPart, interaction,
                                slidePartsByLegacyId: slidePartsByLegacyId,
                                soundContext: soundContext),
                        pictureBullet => ProjectLegacyPictureBullet(ownerPart,
                            pictureBullet))
                    : new TextBody(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(new A.Run(new A.Text(source.Text))));
                if (!hasRichText) {
                    LegacyPptTextProjection.ApplyTextFrame(
                        textBody.BodyProperties, source.TextFrame);
                }
                shape.Append(textBody);
                if (deferTextInteractions) {
                    deferredInteractions!.Add(new LegacyPptDeferredProjection(
                        projectedSlides => {
                            shape.TextBody = LegacyPptTextProjection
                                .CreateTextBody(source.TextBody,
                                    source.TextFrame,
                                    interaction => ProjectLegacyInteraction(
                                        ownerPart, interaction,
                                        slidePartsByLegacyId: projectedSlides,
                                        soundContext: soundContext),
                                    pictureBullet =>
                                        ProjectLegacyPictureBullet(ownerPart,
                                            pictureBullet));
                            ownerPart.RootElement?.Save();
                        }));
                }
            }
            return shape;
        }

        private static ConnectionShape CreateLegacyConnectionShape(LegacyPptShape source, uint shapeId,
            A.ShapeTypeValues geometry) {
            var properties = new ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = ToEmus(source.Bounds.Left), Y = ToEmus(source.Bounds.Top) },
                    new A.Extents {
                        Cx = Math.Max(1L, ToEmus(source.Bounds.Width)),
                        Cy = Math.Max(1L, ToEmus(source.Bounds.Height))
                    }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = geometry });
            LegacyPptShapeGeometryMapper.ApplyExactPresetAdjustments(source.OfficeArtShapeType,
                source.Geometry, properties.GetFirstChild<A.PresetGeometry>()!);
            ApplyLegacyShapeStyle(properties, source);
            ApplyLegacyShapeTransform(properties.Transform2D!, source);
            return new ConnectionShape(
                new NonVisualConnectionShapeProperties(
                    new NonVisualDrawingProperties { Id = shapeId, Name = $"Binary Connector {shapeId - 1U}" },
                    new NonVisualConnectorShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                properties);
        }

        private static GroupShape? CreateLegacyGroupShape(OpenXmlPart ownerPart, LegacyPptShape source,
            uint shapeId, ref uint nextShapeId,
            IReadOnlyDictionary<uint, SlidePart>? slidePartsByLegacyId,
            LegacyPptSoundProjectionContext? soundContext,
            IDictionary<uint, uint>? projectedShapeIds,
            ICollection<LegacyPptDeferredProjection>?
                deferredInteractions) {
            if (!source.GroupCoordinateBounds.HasValue || source.Children.Count == 0) return null;
            LegacyPptBounds coordinate = source.GroupCoordinateBounds.Value;
            var transform = new A.TransformGroup(
                new A.Offset {
                    X = ToEmus(source.Bounds.Left),
                    Y = ToEmus(source.Bounds.Top)
                },
                new A.Extents {
                    Cx = Math.Max(1L, ToEmus(source.Bounds.Width)),
                    Cy = Math.Max(1L, ToEmus(source.Bounds.Height))
                },
                new A.ChildOffset {
                    X = ToEmus(coordinate.Left),
                    Y = ToEmus(coordinate.Top)
                },
                new A.ChildExtents {
                    Cx = Math.Max(1L, ToEmus(coordinate.Width)),
                    Cy = Math.Max(1L, ToEmus(coordinate.Height))
                });
            ApplyLegacyShapeTransform(transform, source);
            var group = new GroupShape(
                new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties { Id = shapeId, Name = $"Binary Group {shapeId - 1U}" },
                    new NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new GroupShapeProperties(transform));
            foreach (LegacyPptShape child in source.Children) {
                OpenXmlElement? projected = CreateLegacyOpenXmlShape(ownerPart, child,
                    ref nextShapeId, slidePartsByLegacyId, soundContext,
                    projectedShapeIds, deferredInteractions);
                if (projected != null) group.Append(projected);
            }
            return group;
        }

        private static Picture? CreateLegacyPicture(OpenXmlPart ownerPart, LegacyPptShape source,
            uint shapeId) {
            if (source.Picture?.HasImportableImage != true || source.Picture.ContentType == null) return null;
            ImagePart imagePart = AddLegacyImagePart(ownerPart, source.Picture);
            string relationshipId = ownerPart.GetIdOfPart(imagePart);
            var properties = new ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = ToEmus(source.Bounds.Left), Y = ToEmus(source.Bounds.Top) },
                    new A.Extents {
                        Cx = Math.Max(1L, ToEmus(source.Bounds.Width)),
                        Cy = Math.Max(1L, ToEmus(source.Bounds.Height))
                    }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle });
            ApplyLegacyShapeTransform(properties.Transform2D!, source);
            var blipFill = new BlipFill(
                new A.Blip { Embed = relationshipId },
                new A.Stretch(new A.FillRectangle()));
            ApplyLegacyPictureCrop(blipFill, source);
            ApplyLegacyPictureEffects(blipFill.Blip, source);
            return new Picture(
                new NonVisualPictureProperties(
                    new NonVisualDrawingProperties { Id = shapeId, Name = $"Binary Picture {shapeId - 1U}" },
                    CreateLegacyNonVisualPictureDrawingProperties(source),
                    new ApplicationNonVisualDrawingProperties()),
                blipFill,
                properties);
        }

        private static ImagePart AddLegacyImagePart(OpenXmlPart ownerPart,
            OfficeIMO.Drawing.Binary.OfficeArtBlipStoreEntry picture) {
            if (picture.ContentType == null || !picture.HasImportableImage) {
                throw new ArgumentException("The OfficeArt BLIP has no importable image payload.",
                    nameof(picture));
            }
            return AddLegacyImagePart(ownerPart, picture.ContentType,
                picture.ImageBytes);
        }

        private static string? ProjectLegacyPictureBullet(
            OpenXmlPart ownerPart, LegacyPptPictureBullet pictureBullet) {
            if (!pictureBullet.HasImportableImage
                || pictureBullet.ContentType == null) return null;
            Dictionary<ushort, string> relationships =
                LegacyPictureBulletRelationships.GetOrCreateValue(
                    ownerPart);
            lock (relationships) {
                if (relationships.TryGetValue(pictureBullet.Index,
                        out string? existingRelationship)) {
                    return existingRelationship;
                }
                byte[] imageBytes = pictureBullet.ImageBytes;
                ImagePart imagePart = AddLegacyImagePart(ownerPart,
                    pictureBullet.ContentType, imageBytes,
                    reuseExisting: false);
                string relationshipId = ownerPart.GetIdOfPart(imagePart);
                relationships.Add(pictureBullet.Index, relationshipId);
                return relationshipId;
            }
        }

        private static ImagePart AddLegacyImagePart(OpenXmlPart ownerPart,
            string contentType, byte[] imageBytes,
            bool reuseExisting = true) {
            foreach (IdPartPair pair in reuseExisting
                         ? ownerPart.Parts
                         : Enumerable.Empty<IdPartPair>()) {
                if (pair.OpenXmlPart is not ImagePart existing
                    || !string.Equals(existing.ContentType, contentType,
                        StringComparison.OrdinalIgnoreCase)) continue;
                using Stream source = existing.GetStream(FileMode.Open,
                    FileAccess.Read);
                if (source.Length != imageBytes.Length) continue;
                using var copy = new MemoryStream();
                source.CopyTo(copy);
                if (copy.ToArray().SequenceEqual(imageBytes)) {
                    return existing;
                }
            }
            PartTypeInfo partType = GetLegacyPicturePartType(contentType)
                .ToPartTypeInfo();
            ImagePart imagePart = ownerPart switch {
                SlidePart slidePart => slidePart.AddImagePart(partType),
                SlideMasterPart masterPart => masterPart.AddImagePart(partType),
                SlideLayoutPart layoutPart => layoutPart.AddImagePart(partType),
                NotesSlidePart notesSlidePart => notesSlidePart.AddImagePart(partType),
                NotesMasterPart notesPart => notesPart.AddImagePart(partType),
                HandoutMasterPart handoutPart => handoutPart.AddImagePart(partType),
                _ => throw new NotSupportedException(
                    $"Legacy pictures cannot be attached to {ownerPart.GetType().Name}.")
            };
            using var stream = new MemoryStream(imageBytes, writable: false);
            imagePart.FeedData(stream);
            return imagePart;
        }

        private static void ApplyLegacyPictureCrop(BlipFill? target, LegacyPptShape source) {
            if (target == null || !source.PictureProperties.HasCrop) return;
            target.SourceRectangle = new A.SourceRectangle {
                Left = ToOpenXmlCrop(source.PictureProperties.CropFromLeft),
                Top = ToOpenXmlCrop(source.PictureProperties.CropFromTop),
                Right = ToOpenXmlCrop(source.PictureProperties.CropFromRight),
                Bottom = ToOpenXmlCrop(source.PictureProperties.CropFromBottom)
            };
        }

        private static int? ToOpenXmlCrop(double? fraction) {
            if (!fraction.HasValue || double.IsNaN(fraction.Value)
                || double.IsInfinity(fraction.Value)) return null;
            double value = Math.Round(fraction.Value * 100000D, MidpointRounding.AwayFromZero);
            return value < int.MinValue || value > int.MaxValue ? null : (int)value;
        }

        private static void ApplyLegacyPictureEffects(A.Blip? target, LegacyPptShape source) {
            if (target == null) return;
            if (source.PictureTransparentColor != null) {
                target.Append(new A.ColorChange {
                    ColorFrom = new A.ColorFrom(
                        new A.RgbColorModelHex {
                            Val = source.PictureTransparentColor
                        }),
                    ColorTo = new A.ColorTo(
                        new A.RgbColorModelHex(
                            new A.Alpha { Val = 0 }) {
                            Val = source.PictureTransparentColor
                        })
                });
            }
            if (source.PictureRecolorColor != null) {
                target.Append(new A.ColorReplacement(
                    new A.RgbColorModelHex {
                        Val = source.PictureRecolorColor
                    }));
            }
            int? brightness = ToOpenXmlPictureAdjustment(
                source.PictureProperties.BrightnessAdjustment);
            int? contrast = ToOpenXmlPictureAdjustment(
                source.PictureProperties.ContrastAdjustment);
            if (brightness.HasValue || contrast.HasValue) {
                target.Append(new A.LuminanceEffect {
                    Brightness = brightness,
                    Contrast = contrast
                });
            }
            if (source.PictureProperties.BiLevel == true) {
                target.Append(new A.BiLevel { Threshold = 50000 });
            } else if (source.PictureProperties.Grayscale == true) {
                target.Append(new A.Grayscale());
            }
        }

        private static int? ToOpenXmlPictureAdjustment(double? fraction) {
            if (!fraction.HasValue || double.IsNaN(fraction.Value)
                || double.IsInfinity(fraction.Value)) return null;
            return (int)Math.Round(Math.Max(-1D, Math.Min(1D, fraction.Value)) * 100000D,
                MidpointRounding.AwayFromZero);
        }

        internal static void ApplyLegacyShapeTransform(A.Transform2D transform, LegacyPptShape source) {
            transform.Rotation = ToOpenXmlRotation(source.Transform.RotationDegrees);
            transform.HorizontalFlip = source.Transform.FlipHorizontal ? true : null;
            transform.VerticalFlip = source.Transform.FlipVertical ? true : null;
        }

        private static void ApplyLegacyShapeTransform(A.TransformGroup transform, LegacyPptShape source) {
            transform.Rotation = ToOpenXmlRotation(source.Transform.RotationDegrees);
            transform.HorizontalFlip = source.Transform.FlipHorizontal ? true : null;
            transform.VerticalFlip = source.Transform.FlipVertical ? true : null;
        }

        private static int? ToOpenXmlRotation(double? degrees) {
            if (!degrees.HasValue) return null;
            double value = degrees.Value * 60000D;
            if (value < int.MinValue || value > int.MaxValue) return null;
            return (int)Math.Round(value);
        }

        internal static void ApplyLegacyShapeStyle(ShapeProperties properties, LegacyPptShape source) {
            OfficeIMO.Drawing.Binary.OfficeArtShapeStyle style = source.Style;
            if (style.FillEnabled == false) {
                SetLegacyShapeFill(properties, new A.NoFill());
            } else if (CreateLegacyShapeGradientFill(source) is A.GradientFill gradient) {
                SetLegacyShapeFill(properties, gradient);
            } else if (source.FillColor != null && style.FillType.GetValueOrDefault() == 0) {
                SetLegacyShapeFill(properties, CreateLegacySolidFill(source.FillColor, style.FillOpacity));
            }
            bool hasLineStyle = style.LineEnabled.HasValue || source.LineColor != null
                || style.LineOpacity.HasValue || style.LineWidthEmus.HasValue || style.LineDashing.HasValue
                || style.LineStartArrowhead.HasValue || style.LineEndArrowhead.HasValue
                || style.LineJoinStyle.HasValue || style.LineEndCapStyle.HasValue;
            if (!hasLineStyle) {
                ApplyLegacyShapeShadow(properties, source);
                return;
            }

            A.Outline outline = properties.GetFirstChild<A.Outline>() ?? new A.Outline();
            if (outline.Parent == null) properties.Append(outline);
            if (style.LineEnabled == false) {
                SetLegacyOutlineFill(outline, new A.NoFill());
                ApplyLegacyShapeShadow(properties, source);
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
            ApplyLegacyShapeShadow(properties, source);
        }

        private static void ApplyLegacyShapeShadow(ShapeProperties properties,
            LegacyPptShape source) {
            OfficeIMO.Drawing.Binary.OfficeArtShapeStyle style = source.Style;
            if (!style.HasProjectableShadow) return;

            int offsetX = style.ShadowOffsetXEmus ?? 0x6338;
            int offsetY = style.ShadowOffsetYEmus ?? 0x6338;
            double angle = Math.Atan2(offsetY, offsetX) * 180D / Math.PI;
            if (angle < 0D) angle += 360D;
            long distance = (long)Math.Round(Math.Sqrt((double)offsetX * offsetX
                + (double)offsetY * offsetY), MidpointRounding.AwayFromZero);
            var color = new A.RgbColorModelHex { Val = source.ShadowColor ?? "808080" };
            color.Append(new A.Alpha { Val = checked((int)Math.Round(
                Math.Max(0D, Math.Min(1D, style.ShadowOpacity ?? 1D)) * 100000D)) });
            var shadow = new A.OuterShadow(color) {
                BlurRadius = Math.Max(0, style.ShadowSoftnessEmus ?? 0),
                Distance = distance,
                Direction = (int)Math.Round(angle * 60000D, MidpointRounding.AwayFromZero),
                RotateWithShape = false
            };
            A.EffectList effects = properties.GetFirstChild<A.EffectList>() ?? new A.EffectList();
            effects.RemoveAllChildren<A.OuterShadow>();
            effects.Append(shadow);
            if (effects.Parent == null) {
                OpenXmlElement? insertBefore = properties.ChildElements.FirstOrDefault(child =>
                    child is not A.Transform2D
                    && child is not A.CustomGeometry
                    && child is not A.PresetGeometry
                    && child is not A.NoFill
                    && child is not A.SolidFill
                    && child is not A.GradientFill
                    && child is not A.BlipFill
                    && child is not A.PatternFill
                    && child is not A.GroupFill
                    && child is not A.Outline);
                if (insertBefore != null) properties.InsertBefore(effects, insertBefore);
                else properties.Append(effects);
            }
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

    }
}
