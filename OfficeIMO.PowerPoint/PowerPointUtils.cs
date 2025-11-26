using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// PowerPoint utility methods based on the working open-xml-sdk-snippets implementation.
    /// CRITICAL: This class contains the exact initialization pattern required to prevent
    /// PowerPoint from showing a "repair" dialog. The order and relationship IDs used here
    /// are very specific and must not be changed.
    /// </summary>
    internal static class PowerPointUtils {
        private const int DefaultRestoredLeftSize = 15989;
        private const int DefaultRestoredTopSize = 94660;
        private const string DefaultTableStyleGuid = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";
        private const string DefaultDocumentAuthor = "OfficeIMO";

        public static PresentationDocument CreatePresentation(string filepath) {
            // Create a presentation at a specified file path. The presentation document type is pptx by default.
            PresentationDocument presentationDoc = PresentationDocument.Create(filepath, PresentationDocumentType.Presentation);
            PresentationPart presentationPart = presentationDoc.AddPresentationPart();
            presentationPart.Presentation = new Presentation();

            CreatePresentationParts(presentationDoc, presentationPart);

            return presentationDoc;
        }

        internal static void CreatePresentationParts(PresentationDocument presentationDocument, PresentationPart presentationPart) {
            SlideMasterIdList slideMasterIdList1 = new SlideMasterIdList(new SlideMasterId() { Id = (UInt32Value)2147483648U, RelationshipId = "rId1" });
            SlideIdList slideIdList1 = new SlideIdList(new SlideId() { Id = (UInt32Value)256U, RelationshipId = "rId2" });
            // Match the common 16:9 widescreen default (same as the shipped blank template)
            SlideSize slideSize1 = new SlideSize() { Cx = 12192000, Cy = 6858000, Type = SlideSizeValues.Screen16x9 };
            NotesSize notesSize1 = new NotesSize() { Cx = 6858000, Cy = 9144000 };
            DefaultTextStyle defaultTextStyle1 = new DefaultTextStyle();

            presentationPart.Presentation.Append(slideMasterIdList1, slideIdList1, slideSize1, notesSize1, defaultTextStyle1);
            presentationPart.Presentation.SaveSubsetFonts = true;

            // Create master and layouts directly under the presentation part so they land at ppt/slideMasters and ppt/slideLayouts.
            SlideMasterPart slideMasterPart1 = presentationPart.AddNewPart<SlideMasterPart>("rId1");
            slideMasterPart1.SlideMaster = CreateSlideMasterSkeleton();

            // Initial layout (Title Slide)
            SlideLayoutPart titleLayout = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId1");
            titleLayout.SlideLayout = CreateTitleSlideLayout();
            titleLayout.AddPart(slideMasterPart1);

            // Additional default layouts (full set of 11 matching a blank PowerPoint)
            CreateAdditionalSlideLayouts(slideMasterPart1, titleLayout);

            // Theme stored under /ppt/theme/theme1.xml and linked from master
            ThemePart themePart1 = CreateTheme(presentationPart);
            slideMasterPart1.AddPart(themePart1, "rId12");

            // Create initial slide and link to the title layout
            SlidePart slidePart1 = presentationPart.AddNewPart<SlidePart>("rId2");
            slidePart1.Slide = CreateBlankSlide();
            slidePart1.AddPart(titleLayout, "rId1");

            CreatePresentationPropertiesPart(presentationPart);
            CreateViewPropertiesPart(presentationPart);
            CreateTableStylesPart(presentationPart);
            EnsureDocumentProperties(presentationPart);
            EnsureThumbnail(presentationDocument);
        }

        private const string RelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        internal static void NormalizeContentTypes(string filePath) {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath)) return;

            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true);
            var entry = archive.GetEntry("[Content_Types].xml");
            if (entry == null) return;

            string xml;
            using (var reader = new StreamReader(entry.Open())) {
                xml = reader.ReadToEnd();
            }

            if (string.IsNullOrWhiteSpace(xml)) return;

            XDocument xdoc;
            try { xdoc = XDocument.Parse(xml, LoadOptions.PreserveWhitespace); }
            catch { xdoc = XDocument.Parse(xml, LoadOptions.None); }

            XNamespace ns = xdoc.Root?.Name.Namespace ?? "http://schemas.openxmlformats.org/package/2006/content-types";
            var types = xdoc.Root ?? new XElement(ns + "Types");

            void EnsureDefault(string ext, string contentType) {
                var matches = types.Elements(ns + "Default").Where(e => string.Equals((string?)e.Attribute("Extension"), ext, StringComparison.OrdinalIgnoreCase)).ToList();
                if (matches.Count == 0) {
                    types.AddFirst(new XElement(ns + "Default", new XAttribute("Extension", ext), new XAttribute("ContentType", contentType)));
                } else {
                    matches[0].SetAttributeValue("ContentType", contentType);
                    for (int i = 1; i < matches.Count; i++) matches[i].Remove();
                }
            }

            void EnsureOverride(string partName, string contentType) {
                var el = types.Elements(ns + "Override").FirstOrDefault(e => string.Equals((string?)e.Attribute("PartName"), partName, StringComparison.OrdinalIgnoreCase));
                if (el == null) {
                    types.Add(new XElement(ns + "Override", new XAttribute("PartName", partName), new XAttribute("ContentType", contentType)));
                } else {
                    el.SetAttributeValue("ContentType", contentType);
                }
            }

            EnsureDefault("rels", "application/vnd.openxmlformats-package.relationships+xml");
            EnsureDefault("xml", "application/xml");
            EnsureOverride("/ppt/presentation.xml", "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml");

            // Replace entry content
            entry.Delete();
            var newEntry = archive.CreateEntry("[Content_Types].xml", CompressionLevel.Optimal);
            using var writer = new StreamWriter(newEntry.Open());
            xdoc.Save(writer);
        }

        // Thumbnail extracted from Assets/PowerPointTemplates/PowerPointBlank.pptx (docProps/thumbnail.jpeg)
        private static readonly byte[] ThumbnailBytes = Convert.FromBase64String(
            "/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAMCAgMCAgMDAwMEAwMEBQgFBQQEBQoHBwYIDAoMDAsKCwsNDhIQDQ4RDgsLEBYQERMUFRUVDA8XGBYUGBIUFRT/2wBDAQMEBAUEBQkFBQkUDQsNFBQUEBAUEBQUFBQUEBQUFBQUEBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQT/wAARCAAkAA4DAREAAhEBAxEB/8QAHQABAAEEAwEAAAAAAAAAAAAAAAYHAgUEAwIBCf/EADgQAAECAwQJBgUEAwAAAAAAAAECAwAEBRESBiExBxNBUWFxFCIyobHBFBZCksEVM1NiY//EABsBAQEAAwEBAQAAAAAAAAAAAAABAgMEAQUG/8QALREBAAIBBAECBQMEAwAAAAAAAAECAxEEEiExQQUTIlFhcZGh0fAUI0JSsf/aAAwDAQACEQMRAD8A9P8A0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAARvbLcm6goc0JdU5pUW4RjjdM9i9Wxj6WutKLa9c7WJXu1L05PwnHA/0eFX3c/B/5UX5xvV+zY+q2cL6ZrGdOHFjT3yL1clOVruomDfaYy6USsvXxP6c8g3k3kDWK6byPXF0jHGU0kqdZSk0nFJJySgAAAAAAAAAAAAAAAAAAABK/pv3iQvqVps2fwXyrGnUsPW9jHTV6hBZ6wqn6UdZacK3WN+S2z87zuzp/8AflmWzNtZ0rJQ1pqSikopSnVKUqk1SlKSgAAbcRLRqy5F06ys3xcrKzJSmkoxVFLKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKU//Z");

        internal static NotesMasterPart EnsureNotesMasterPart(PresentationPart presentationPart) {
            NotesMasterPart notesMasterPart = presentationPart.NotesMasterPart ?? presentationPart.AddNewPart<NotesMasterPart>();

            if (notesMasterPart.NotesMaster == null) {
                notesMasterPart.NotesMaster = CreateDefaultNotesMaster();
            }

            Presentation presentation = presentationPart.Presentation ??= new Presentation();
            NotesMasterIdList notesMasterIdList = presentation.NotesMasterIdList ??= new NotesMasterIdList();

            string relationshipId = presentationPart.GetIdOfPart(notesMasterPart);
            bool hasEntry = notesMasterIdList
                .Elements<NotesMasterId>()
                .Any(existing => GetRelationshipId(existing) == relationshipId);
            if (!hasEntry) {
                NotesMasterId notesMasterId = new NotesMasterId();
                SetRelationshipId(notesMasterId, relationshipId);
                notesMasterIdList.AppendChild(notesMasterId);
            }

            return notesMasterPart;
        }

        private static NotesMaster CreateDefaultNotesMaster() {
            ShapeTree shapeTree = new ShapeTree();
            shapeTree.Append(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 1U, Name = "Notes Group Shape" },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new P.GroupShapeProperties(new D.TransformGroup()));

            shapeTree.Append(
                CreatePlaceholderShape(2U, "Notes Placeholder", PlaceholderValues.Body, 1U, includeEndParagraph: true),
                CreatePlaceholderShape(3U, "Slide Image Placeholder", PlaceholderValues.SlideImage, 2U, includeEndParagraph: false),
                CreatePlaceholderShape(4U, "Date Placeholder", PlaceholderValues.DateAndTime, 3U, includeEndParagraph: true),
                CreatePlaceholderShape(5U, "Slide Number Placeholder", PlaceholderValues.SlideNumber, 4U, includeEndParagraph: true),
                CreatePlaceholderShape(6U, "Footer Placeholder", PlaceholderValues.Footer, 5U, includeEndParagraph: true));

            Background background = new Background(new BackgroundProperties(new D.NoFill()));

            return new NotesMaster(
                new CommonSlideData(background, shapeTree),
                new P.ColorMap {
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
                new NotesStyle(
                    new D.Level1ParagraphProperties(new D.DefaultRunProperties { Language = "en-US" }),
                    new D.Level2ParagraphProperties(new D.DefaultRunProperties { Language = "en-US" }),
                    new D.Level3ParagraphProperties(new D.DefaultRunProperties { Language = "en-US" }),
                    new D.Level4ParagraphProperties(new D.DefaultRunProperties { Language = "en-US" }),
                    new D.Level5ParagraphProperties(new D.DefaultRunProperties { Language = "en-US" }),
                    new D.Level6ParagraphProperties(new D.DefaultRunProperties { Language = "en-US" }),
                    new D.Level7ParagraphProperties(new D.DefaultRunProperties { Language = "en-US" }),
                    new D.Level8ParagraphProperties(new D.DefaultRunProperties { Language = "en-US" }),
                    new D.Level9ParagraphProperties(new D.DefaultRunProperties { Language = "en-US" }))
            );
        }

        private static P.Shape CreatePlaceholderShape(uint id, string name, PlaceholderValues type, uint index, bool includeEndParagraph) {
            P.Shape shape = new P.Shape(
                new P.NonVisualShapeProperties(
                    new P.NonVisualDrawingProperties { Id = id, Name = name },
                    new P.NonVisualShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape { Type = type, Index = index })),
                new P.ShapeProperties(),
                new P.TextBody(
                    new D.BodyProperties(),
                    new D.ListStyle()));

            D.Paragraph paragraph = new D.Paragraph();
            if (includeEndParagraph) {
                paragraph.Append(new D.EndParagraphRunProperties { Language = "en-US" });
            }

            shape.TextBody!.Append(paragraph);
            return shape;
        }

        private static string? GetRelationshipId(NotesMasterId notesMasterId) {
            OpenXmlAttribute attribute = notesMasterId.GetAttribute("id", RelationshipNamespace);
            return string.IsNullOrEmpty(attribute.Value) ? null : attribute.Value;
        }

        private static void SetRelationshipId(NotesMasterId notesMasterId, string relationshipId) {
            notesMasterId.SetAttribute(new OpenXmlAttribute("r", "id", RelationshipNamespace, relationshipId));
        }

        private static Slide CreateBlankSlide() {
            return new Slide(
                new CommonSlideData(
                    new ShapeTree(
                        new P.NonVisualGroupShapeProperties(
                            new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                            new P.NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()),
                        new GroupShapeProperties(new TransformGroup()))),
                new ColorMapOverride(new MasterColorMapping()));
        }

        private static SlideMaster CreateSlideMasterSkeleton() {
            return new SlideMaster(
                new CommonSlideData(new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new TransformGroup()))),
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
                new P.GroupShapeProperties(new D.TransformGroup()));

            foreach (OpenXmlElement shape in shapes) {
                shapeTree.Append(shape);
            }

            return new SlideLayout(
                new CommonSlideData(shapeTree) { Name = name },
                new ColorMapOverride(new MasterColorMapping())) { Type = layoutType };
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

        private static void CreatePresentationPropertiesPart(PresentationPart presentationPart) {
            PresentationPropertiesPart part = presentationPart.PresentationPropertiesPart ?? presentationPart.AddNewPart<PresentationPropertiesPart>();

            part.PresentationProperties ??= new PresentationProperties();

            ShowProperties showProperties = part.PresentationProperties.ShowProperties ??= new ShowProperties();
            showProperties.ShowNarration = false;
            showProperties.ShowAnimation = true;
            showProperties.UseTimings = true;
        }

        private static void CreateViewPropertiesPart(PresentationPart presentationPart) {
            ViewPropertiesPart viewPart = presentationPart.ViewPropertiesPart ?? presentationPart.AddNewPart<ViewPropertiesPart>();

            NormalViewProperties normalViewProperties = new NormalViewProperties(
                new RestoredLeft() { Size = DefaultRestoredLeftSize, AutoAdjust = false },
                new RestoredTop() { Size = DefaultRestoredTopSize }
            );

            SlideViewProperties slideViewProperties = new SlideViewProperties();

            CommonSlideViewProperties commonSlideViewProperties = new CommonSlideViewProperties() { SnapToGrid = false };
            CommonViewProperties commonViewProperties = new CommonViewProperties() { VariableScale = true };

            ScaleFactor scaleFactor = new ScaleFactor();
            scaleFactor.Append(new D.ScaleX() { Numerator = 142, Denominator = 100 });
            scaleFactor.Append(new D.ScaleY() { Numerator = 142, Denominator = 100 });
            commonViewProperties.Append(scaleFactor);
            commonViewProperties.Append(new Origin() { X = 0L, Y = 0L });

            commonSlideViewProperties.Append(commonViewProperties);
            slideViewProperties.Append(commonSlideViewProperties);

            NotesTextViewProperties notesTextViewProperties = new NotesTextViewProperties();
            CommonViewProperties notesCommonViewProperties = new CommonViewProperties();
            ScaleFactor notesScaleFactor = new ScaleFactor();
            notesScaleFactor.Append(new D.ScaleX() { Numerator = 1, Denominator = 1 });
            notesScaleFactor.Append(new D.ScaleY() { Numerator = 1, Denominator = 1 });
            notesCommonViewProperties.Append(notesScaleFactor);
            notesCommonViewProperties.Append(new Origin() { X = 0L, Y = 0L });
            notesTextViewProperties.Append(notesCommonViewProperties);

            GridSpacing gridSpacing = new GridSpacing() { Cx = 72008L, Cy = 72008L };

            ViewProperties viewProperties = new ViewProperties();
            viewProperties.Append(normalViewProperties);
            viewProperties.Append(slideViewProperties);
            viewProperties.Append(notesTextViewProperties);
            viewProperties.Append(gridSpacing);

            viewPart.ViewProperties = viewProperties;
        }

        private static void CreateTableStylesPart(PresentationPart presentationPart) {
            TableStylesPart tableStylesPart = presentationPart.TableStylesPart ?? presentationPart.AddNewPart<TableStylesPart>();

            D.TableStyleList tableStyleList = new D.TableStyleList() { Default = DefaultTableStyleGuid };
            tableStyleList.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            tableStylesPart.TableStyleList = tableStyleList;
        }

        private static void EnsureDocumentProperties(PresentationPart presentationPart) {
            if (presentationPart.OpenXmlPackage is not PresentationDocument presentationDocument) {
                return;
            }

            ExtendedFilePropertiesPart extendedPart = presentationDocument.ExtendedFilePropertiesPart ?? presentationDocument.AddExtendedFilePropertiesPart();
            if (extendedPart.Properties == null) {
                extendedPart.Properties = new Ap.Properties();
            }

            extendedPart.Properties.TotalTime ??= new Ap.TotalTime() { Text = "0" };
            extendedPart.Properties.Application ??= new Ap.Application() { Text = "Microsoft Office PowerPoint" };
            extendedPart.Properties.PresentationFormat ??= new Ap.PresentationFormat() { Text = "Widescreen" };
            extendedPart.Properties.Slides ??= new Ap.Slides() { Text = "1" };
            extendedPart.Properties.Notes ??= new Ap.Notes() { Text = "0" };
            extendedPart.Properties.HiddenSlides ??= new Ap.HiddenSlides() { Text = "0" };

            DateTime timestamp = DateTime.UtcNow;
            CoreFilePropertiesPart corePart = presentationDocument.CoreFilePropertiesPart ?? presentationDocument.AddCoreFilePropertiesPart();
            bool coreHasContent;

            using (Stream coreStream = corePart.GetStream(FileMode.OpenOrCreate, FileAccess.Read)) {
                coreHasContent = coreStream.Length > 0;
            }

            if (!coreHasContent) {
                InitializeCoreFilePropertiesPart(corePart, timestamp);
            }

            var packageProperties = presentationDocument.PackageProperties;

            if (string.IsNullOrEmpty(packageProperties.Creator)) {
                packageProperties.Creator = DefaultDocumentAuthor;
            }

            if (string.IsNullOrEmpty(packageProperties.LastModifiedBy)) {
                packageProperties.LastModifiedBy = DefaultDocumentAuthor;
            }

            if (packageProperties.Created == null) {
                packageProperties.Created = timestamp;
            }

            if (packageProperties.Modified == null) {
                packageProperties.Modified = timestamp;
            }
        }

        private static void EnsureThumbnail(PresentationDocument doc) {
            if (doc.ThumbnailPart != null) return;
            ThumbnailPart thumbnailPart = doc.AddThumbnailPart("image/jpeg");
            using Stream stream = thumbnailPart.GetStream(FileMode.Create, FileAccess.Write);
            stream.Write(ThumbnailBytes, 0, ThumbnailBytes.Length);
        }

        private static void InitializeCoreFilePropertiesPart(CoreFilePropertiesPart corePart, DateTime timestamp) {
            XNamespace cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
            XNamespace dc = "http://purl.org/dc/elements/1.1/";
            XNamespace dcterms = "http://purl.org/dc/terms/";
            XNamespace dcmitype = "http://purl.org/dc/dcmitype/";
            XNamespace xsi = "http://www.w3.org/2001/XMLSchema-instance";

            string serializedTimestamp = timestamp.ToString("s", CultureInfo.InvariantCulture) + "Z";

            XDocument coreDocument = new XDocument(
                new XElement(cp + "coreProperties",
                    new XAttribute(XNamespace.Xmlns + "cp", cp),
                    new XAttribute(XNamespace.Xmlns + "dc", dc),
                    new XAttribute(XNamespace.Xmlns + "dcterms", dcterms),
                    new XAttribute(XNamespace.Xmlns + "dcmitype", dcmitype),
                    new XAttribute(XNamespace.Xmlns + "xsi", xsi),
                    new XElement(dc + "creator", DefaultDocumentAuthor),
                    new XElement(cp + "lastModifiedBy", DefaultDocumentAuthor),
                    new XElement(dcterms + "created",
                        new XAttribute(xsi + "type", "dcterms:W3CDTF"),
                        serializedTimestamp),
                    new XElement(dcterms + "modified",
                        new XAttribute(xsi + "type", "dcterms:W3CDTF"),
                        serializedTimestamp))
            );

            using Stream stream = corePart.GetStream(FileMode.Create, FileAccess.Write);
            coreDocument.Save(stream);
        }

        private static ThemePart CreateTheme(PresentationPart presentationPart) {
            // Theme should live under /ppt/theme/theme1.xml; create it on the presentation part
            ThemePart themePart1 = presentationPart.AddNewPart<ThemePart>();
            D.Theme theme1 = new D.Theme() { Name = "Office Theme" };

            D.ThemeElements themeElements1 = new D.ThemeElements(
            new D.ColorScheme(
              new D.Dark1Color(new D.SystemColor() { Val = D.SystemColorValues.WindowText, LastColor = "000000" }),
              new D.Light1Color(new D.SystemColor() { Val = D.SystemColorValues.Window, LastColor = "FFFFFF" }),
              new D.Dark2Color(new D.RgbColorModelHex() { Val = "1F497D" }),
              new D.Light2Color(new D.RgbColorModelHex() { Val = "EEECE1" }),
              new D.Accent1Color(new D.RgbColorModelHex() { Val = "4F81BD" }),
              new D.Accent2Color(new D.RgbColorModelHex() { Val = "C0504D" }),
              new D.Accent3Color(new D.RgbColorModelHex() { Val = "9BBB59" }),
              new D.Accent4Color(new D.RgbColorModelHex() { Val = "8064A2" }),
              new D.Accent5Color(new D.RgbColorModelHex() { Val = "4BACC6" }),
              new D.Accent6Color(new D.RgbColorModelHex() { Val = "F79646" }),
              new D.Hyperlink(new D.RgbColorModelHex() { Val = "0000FF" }),
              new D.FollowedHyperlinkColor(new D.RgbColorModelHex() { Val = "800080" })) { Name = "Office" },
              new D.FontScheme(
              new D.MajorFont(
              new D.LatinFont() { Typeface = "Calibri" },
              new D.EastAsianFont() { Typeface = "" },
              new D.ComplexScriptFont() { Typeface = "" }),
              new D.MinorFont(
              new D.LatinFont() { Typeface = "Calibri" },
              new D.EastAsianFont() { Typeface = "" },
              new D.ComplexScriptFont() { Typeface = "" })) { Name = "Office" },
              new D.FormatScheme(
              new D.FillStyleList(
              new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.PhColor }),
              new D.GradientFill(
                new D.GradientStopList(
                new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 50000 },
                  new D.SaturationModulation() { Val = 300000 }) { Val = D.SchemeColorValues.PhColor }) { Position = 0 },
                new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 37000 },
                 new D.SaturationModulation() { Val = 300000 }) { Val = D.SchemeColorValues.PhColor }) { Position = 35000 },
                new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 15000 },
                 new D.SaturationModulation() { Val = 350000 }) { Val = D.SchemeColorValues.PhColor }) { Position = 100000 }
                ),
                new D.LinearGradientFill() { Angle = 16200000, Scaled = true }),
              new D.NoFill(),
              new D.PatternFill(),
              new D.GroupFill()),
              new D.LineStyleList(
              new D.Outline(
                new D.SolidFill(
                new D.SchemeColor(
                  new D.Shade() { Val = 95000 },
                  new D.SaturationModulation() { Val = 105000 }) { Val = D.SchemeColorValues.PhColor }),
                new D.PresetDash() { Val = D.PresetLineDashValues.Solid }) {
                  Width = 9525,
                  CapType = D.LineCapValues.Flat,
                  CompoundLineType = D.CompoundLineValues.Single,
                  Alignment = D.PenAlignmentValues.Center
              },
              new D.Outline(
                new D.SolidFill(
                new D.SchemeColor(
                  new D.Shade() { Val = 95000 },
                  new D.SaturationModulation() { Val = 105000 }) { Val = D.SchemeColorValues.PhColor }),
                new D.PresetDash() { Val = D.PresetLineDashValues.Solid }) {
                  Width = 9525,
                  CapType = D.LineCapValues.Flat,
                  CompoundLineType = D.CompoundLineValues.Single,
                  Alignment = D.PenAlignmentValues.Center
              },
              new D.Outline(
                new D.SolidFill(
                new D.SchemeColor(
                  new D.Shade() { Val = 95000 },
                  new D.SaturationModulation() { Val = 105000 }) { Val = D.SchemeColorValues.PhColor }),
                new D.PresetDash() { Val = D.PresetLineDashValues.Solid }) {
                  Width = 9525,
                  CapType = D.LineCapValues.Flat,
                  CompoundLineType = D.CompoundLineValues.Single,
                  Alignment = D.PenAlignmentValues.Center
              }),
              new D.EffectStyleList(
              new D.EffectStyle(
                new D.EffectList(
                new D.OuterShadow(
                  new D.RgbColorModelHex(
                  new D.Alpha() { Val = 38000 }) { Val = "000000" }) { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false })),
              new D.EffectStyle(
                new D.EffectList(
                new D.OuterShadow(
                  new D.RgbColorModelHex(
                  new D.Alpha() { Val = 38000 }) { Val = "000000" }) { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false })),
              new D.EffectStyle(
                new D.EffectList(
                new D.OuterShadow(
                  new D.RgbColorModelHex(
                  new D.Alpha() { Val = 38000 }) { Val = "000000" }) { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false }))),
              new D.BackgroundFillStyleList(
              new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.PhColor }),
              new D.GradientFill(
                new D.GradientStopList(
                new D.GradientStop(
                  new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 }) { Val = D.SchemeColorValues.PhColor }) { Position = 0 },
                new D.GradientStop(
                  new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 }) { Val = D.SchemeColorValues.PhColor }) { Position = 0 },
                new D.GradientStop(
                  new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 }) { Val = D.SchemeColorValues.PhColor }) { Position = 0 }),
                new D.LinearGradientFill() { Angle = 16200000, Scaled = true }),
              new D.GradientFill(
                new D.GradientStopList(
                new D.GradientStop(
                  new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 }) { Val = D.SchemeColorValues.PhColor }) { Position = 0 },
                new D.GradientStop(
                  new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 }) { Val = D.SchemeColorValues.PhColor }) { Position = 0 }),
                new D.LinearGradientFill() { Angle = 16200000, Scaled = true }))) { Name = "Office" });

            theme1.Append(themeElements1);
            theme1.Append(new D.ObjectDefaults());
            theme1.Append(new D.ExtraColorSchemeList());

            themePart1.Theme = theme1;
            return themePart1;
        }
    }
}
