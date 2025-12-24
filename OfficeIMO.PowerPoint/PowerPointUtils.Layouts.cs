using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointUtils {
        private static Slide CreateBlankSlide() {
            return new Slide(
                new CommonSlideData(
                    new ShapeTree(
                        new P.NonVisualGroupShapeProperties(
                            new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                            new P.NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()),
                        CreateDefaultGroupShapeProperties())),
                new ColorMapOverride(new D.MasterColorMapping()));
        }

        private static SlideMaster CreateSlideMasterSkeleton() {
            return new SlideMaster(
                new CommonSlideData(new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    CreateDefaultGroupShapeProperties())),
                new P.ColorMap() {
                    Background1 = D.ColorSchemeIndexValues.Light1,
                    Text1 = D.ColorSchemeIndexValues.Dark1,
                    Background2 = D.ColorSchemeIndexValues.Light2,
                    Text2 = D.ColorSchemeIndexValues.Dark2,
                    Accent1 = D.ColorSchemeIndexValues.Accent1,
                    Accent2 = D.ColorSchemeIndexValues.Accent2,
                    Accent3 = D.ColorSchemeIndexValues.Accent3,
                    Accent4 = D.ColorSchemeIndexValues.Accent4,
                    Accent5 = D.ColorSchemeIndexValues.Accent5,
                    Accent6 = D.ColorSchemeIndexValues.Accent6,
                    Hyperlink = D.ColorSchemeIndexValues.Hyperlink,
                    FollowedHyperlink = D.ColorSchemeIndexValues.FollowedHyperlink
                },
                new SlideLayoutIdList(new SlideLayoutId() { Id = (UInt32Value)2147483649U, RelationshipId = "rId1" }),
                new TextStyles(new TitleStyle(), new BodyStyle(), new OtherStyle()));
        }

        private static void CreateAdditionalSlideLayouts(SlideMasterPart slideMasterPart, SlideLayoutPart initialLayoutPart) {
            List<(SlideLayoutPart Part, string RelationshipId, uint LayoutId)> layoutEntries = new();

            string initialRelationshipId = slideMasterPart.GetIdOfPart(initialLayoutPart);
            layoutEntries.Add((initialLayoutPart, initialRelationshipId, 2147483649U));

            foreach (SlideLayoutDefinition definition in GetDefaultSlideLayoutDefinitions()) {
                SlideLayoutPart layoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>(definition.RelationshipId);
                layoutPart.SlideLayout = definition.CreateLayout();
                layoutPart.AddPart(slideMasterPart);
                layoutEntries.Add((layoutPart, definition.RelationshipId, definition.LayoutId));
            }

            SlideLayoutIdList slideLayoutIdList = slideMasterPart.SlideMaster.SlideLayoutIdList ?? new SlideLayoutIdList();
            slideLayoutIdList.RemoveAllChildren<SlideLayoutId>();

            foreach ((SlideLayoutPart Part, string RelationshipId, uint LayoutId) entry in layoutEntries) {
                slideLayoutIdList.Append(new SlideLayoutId() { Id = (UInt32Value)entry.LayoutId, RelationshipId = entry.RelationshipId });
            }

            slideMasterPart.SlideMaster.SlideLayoutIdList = slideLayoutIdList;
        }

        private static IEnumerable<SlideLayoutDefinition> GetDefaultSlideLayoutDefinitions() {
            yield return new SlideLayoutDefinition("rId2", 2147483650U, CreateTitleAndContentLayout);
            yield return new SlideLayoutDefinition("rId3", 2147483651U, CreateSectionHeaderLayout);
            yield return new SlideLayoutDefinition("rId4", 2147483652U, CreateTwoContentLayout);
            yield return new SlideLayoutDefinition("rId5", 2147483653U, CreateComparisonLayout);
            yield return new SlideLayoutDefinition("rId6", 2147483654U, CreateTitleOnlyLayout);
            yield return new SlideLayoutDefinition("rId7", 2147483655U, CreateBlankLayout);
            yield return new SlideLayoutDefinition("rId8", 2147483656U, CreatePictureWithCaptionLayout);
            yield return new SlideLayoutDefinition("rId9", 2147483657U, CreateTitleAndVerticalTextLayout);
            yield return new SlideLayoutDefinition("rId10", 2147483658U, CreateVerticalTitleAndTextLayout);
            yield return new SlideLayoutDefinition("rId11", 2147483659U, CreateTwoContentWithCaptionLayout);
        }

        private static SlideLayout CreateTitleSlideLayout() {
            P.Shape titleShape = CreateLayoutPlaceholderShape(2U, "Title Placeholder 1", PlaceholderValues.Title, 0U, 838200L, 365125L, 7772400L, 1470025L);
            P.Shape subtitleShape = CreateLayoutPlaceholderShape(3U, "Subtitle Placeholder 2", PlaceholderValues.SubTitle, 1U, 838200L, 2174875L, 7772400L, 1470025L);

            return CreateSlideLayout("Title Slide", SlideLayoutValues.Title, titleShape, subtitleShape);
        }

        private static SlideLayout CreateTitleAndContentLayout() {
            P.Shape titleShape = CreateLayoutPlaceholderShape(2U, "Title Placeholder 1", PlaceholderValues.Title, 0U, 838200L, 365125L, 7772400L, 1470025L);
            P.Shape contentShape = CreateLayoutPlaceholderShape(3U, "Content Placeholder 2", PlaceholderValues.Body, 1U, 838200L, 2174875L, 7772400L, 3962400L);

            return CreateSlideLayout("Title and Content", SlideLayoutValues.Text, titleShape, contentShape);
        }

        private static SlideLayout CreateSectionHeaderLayout() {
            P.Shape titleShape = CreateLayoutPlaceholderShape(2U, "Section Header Title 1", PlaceholderValues.CenteredTitle, 0U, 838200L, 365125L, 7772400L, 1470025L);
            P.Shape contentShape = CreateLayoutPlaceholderShape(3U, "Section Header Text 2", PlaceholderValues.Body, 1U, 838200L, 2174875L, 7772400L, 1470025L);

            return CreateSlideLayout("Section Header", SlideLayoutValues.SectionHeader, titleShape, contentShape);
        }

        private static SlideLayout CreateTwoContentLayout() {
            P.Shape titleShape = CreateLayoutPlaceholderShape(2U, "Two Content Title 1", PlaceholderValues.Title, 0U, 838200L, 365125L, 7772400L, 1470025L);
            P.Shape leftContent = CreateLayoutPlaceholderShape(3U, "Content Placeholder 2", PlaceholderValues.Body, 1U, 685800L, 2174875L, 3657600L, 3962400L);
            P.Shape rightContent = CreateLayoutPlaceholderShape(4U, "Content Placeholder 3", PlaceholderValues.Body, 2U, 4127500L, 2174875L, 3657600L, 3962400L);

            return CreateSlideLayout("Two Content", SlideLayoutValues.TwoColumnText, titleShape, leftContent, rightContent);
        }

        private static SlideLayout CreateTitleOnlyLayout() {
            P.Shape titleShape = CreateLayoutPlaceholderShape(2U, "Title Placeholder 1", PlaceholderValues.Title, 0U, 838200L, 365125L, 7772400L, 1470025L);

            return CreateSlideLayout("Title Only", SlideLayoutValues.TitleOnly, titleShape);
        }

        private static SlideLayout CreateBlankLayout() {
            return CreateSlideLayout("Blank", SlideLayoutValues.Blank);
        }

        private static SlideLayout CreateComparisonLayout() {
            P.Shape titleShape = CreateLayoutPlaceholderShape(2U, "Comparison Title 1", PlaceholderValues.Title, 0U, 457200L, 365125L, 8048625L, 457200L);
            P.Shape leftContent = CreateLayoutPlaceholderShape(3U, "Left Text Placeholder 2", PlaceholderValues.Body, 1U, 457200L, 899158L, 3889375L, 3504892L);
            P.Shape rightContent = CreateLayoutPlaceholderShape(4U, "Right Text Placeholder 3", PlaceholderValues.Body, 2U, 457200L + 4000000L, 899158L, 3889375L, 3504892L);
            return CreateSlideLayout("Comparison", SlideLayoutValues.TwoObjects, titleShape, leftContent, rightContent);
        }

        private static SlideLayout CreatePictureWithCaptionLayout() {
            P.Shape titleShape = CreateLayoutPlaceholderShape(2U, "Picture Title 1", PlaceholderValues.Title, 0U, 838200L, 365125L, 7772400L, 457200L);
            P.Shape caption = CreateLayoutPlaceholderShape(3U, "Caption Placeholder 2", PlaceholderValues.Body, 1U, 838200L, 1912625L, 7772400L, 1143000L);
            P.Shape picture = CreateLayoutPlaceholderShape(4U, "Picture Placeholder 3", PlaceholderValues.Picture, 2U, 838200L, 760000L, 7772400L, 1016000L);
            return CreateSlideLayout("Picture with Caption", SlideLayoutValues.PictureText, titleShape, picture, caption);
        }

        private static SlideLayout CreateTitleAndVerticalTextLayout() {
            P.Shape titleShape = CreateLayoutPlaceholderShape(2U, "Vertical Title 1", PlaceholderValues.Title, 0U, 914400L, 365125L, 1828800L, 6858000L);
            P.Shape verticalText = CreateLayoutPlaceholderShape(3U, "Vertical Text 2", PlaceholderValues.Body, 1U, 2743200L, 365125L, 5486400L, 6858000L);
            return CreateSlideLayout("Title and Vertical Text", SlideLayoutValues.VerticalTitleAndText, titleShape, verticalText);
        }

        private static SlideLayout CreateVerticalTitleAndTextLayout() {
            P.Shape verticalTitle = CreateLayoutPlaceholderShape(2U, "Vertical Title 1", PlaceholderValues.Title, 0U, 914400L, 365125L, 2743200L, 6858000L);
            P.Shape text = CreateLayoutPlaceholderShape(3U, "Text Placeholder 2", PlaceholderValues.Body, 1U, 365125L, 365125L, 914400L, 6858000L);
            return CreateSlideLayout("Vertical Title and Text", SlideLayoutValues.VerticalText, verticalTitle, text);
        }

        private static SlideLayout CreateTwoContentWithCaptionLayout() {
            P.Shape title = CreateLayoutPlaceholderShape(2U, "Title Placeholder 1", PlaceholderValues.Title, 0U, 838200L, 365125L, 7772400L, 1470025L);
            P.Shape leftContent = CreateLayoutPlaceholderShape(3U, "Content Placeholder 2", PlaceholderValues.Body, 1U, 685800L, 2174875L, 3657600L, 3962400L);
            P.Shape rightContent = CreateLayoutPlaceholderShape(4U, "Content Placeholder 3", PlaceholderValues.Body, 2U, 4127500L, 2174875L, 3657600L, 3962400L);
            P.Shape caption = CreateLayoutPlaceholderShape(5U, "Caption Placeholder 4", PlaceholderValues.Object, 3U, 685800L, 6200000L, 7090000L, 600000L);
            return CreateSlideLayout("Two Content with Caption", SlideLayoutValues.TwoObjects, title, leftContent, rightContent, caption);
        }

        private static SlideLayout CreateSlideLayout(string name, SlideLayoutValues layoutType, params OpenXmlElement[] shapes) {
            P.ShapeTree shapeTree = new(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties() { Id = 1U, Name = "" },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                CreateDefaultGroupShapeProperties());

            foreach (OpenXmlElement shape in shapes) {
                shapeTree.Append(shape);
            }

            return new SlideLayout(
                new CommonSlideData(shapeTree) { Name = name },
                new ColorMapOverride(new D.MasterColorMapping())) { Type = layoutType };
        }

        private static P.Shape CreateLayoutPlaceholderShape(uint id, string name, PlaceholderValues type, uint index, long left, long top, long width, long height) {
            PlaceholderShape placeholderShape = new() { Type = type, Index = index };

            return new P.Shape(
                new P.NonVisualShapeProperties(
                    new P.NonVisualDrawingProperties { Id = id, Name = name },
                    new P.NonVisualShapeDrawingProperties(new D.ShapeLocks { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties(placeholderShape)),
                new P.ShapeProperties(
                    new D.Transform2D(
                        new D.Offset { X = left, Y = top },
                        new D.Extents { Cx = width, Cy = height }),
                    new D.PresetGeometry(new D.AdjustValueList()) { Preset = D.ShapeTypeValues.Rectangle }),
                new P.TextBody(
                    new D.BodyProperties(),
                    new D.ListStyle(),
                    new D.Paragraph(new D.EndParagraphRunProperties { Language = "en-US" }))); 
        }

        private readonly struct SlideLayoutDefinition {
            public SlideLayoutDefinition(string relationshipId, uint layoutId, Func<SlideLayout> createLayout) {
                RelationshipId = relationshipId;
                LayoutId = layoutId;
                CreateLayout = createLayout;
            }

            public string RelationshipId { get; }
            public uint LayoutId { get; }
            public Func<SlideLayout> CreateLayout { get; }
        }

    }
}
