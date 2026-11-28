using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.OpenXml;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public class PowerPointTemplateWorkflowTests {
        [Fact]
        public void ClearsTheSlideListAndRetainsTheTemplateAndExtensionMarkup() {
            const string showName = "Template show";
            const string decoy = "RelationshipInsideExtension";
            using var source = PowerPointDocument.Create();
            PowerPointSlide first = source.AddSlide();
            first.AddTextBox("Template slide");
            PowerPointSlide second = source.AddSlide();
            source.AddCustomShow(showName, new[] { first, second });
            source.AddSection("Template section", 0);
            byte[] seed = source.ToBytes();
            using var original = OfficeOpcPackage.Open(new MemoryStream(seed, writable: false));
            string main = original.GetRelatedPartName(null, OfficeOpcPackage.OfficeDocumentRelationshipType)!;
            using var edit = OfficeOpcEditablePackage.Open(new MemoryStream(seed, writable: false));
            OfficeXmlTree root = edit.GetXml(main);
            const int ns = OfficeXmlNamespaces.PresentationML;
            int extensions = root.FindChild(root.Root, ns, "extLst");
            int extension = root.CreateElement(ns, "ext");
            root.SetAttribute(extension, OfficeXmlNamespaces.None, "uri", "urn:officeimo:fixture:slide-list");
            root.Append(extensions, extension);
            int shadow = root.CreateElement(ns, "sldIdLst");
            root.Append(extension, shadow);
            int item = root.CreateElement(ns, "sldId");
            root.SetAttribute(item, OfficeXmlNamespaces.Relationships, "id", decoy);
            root.Append(shadow, item);
            using var bytes = new MemoryStream();
            edit.Save(bytes);
            byte[] fixture = bytes.ToArray();

            using var baseline = PowerPointDocument.Load(new MemoryStream(fixture, writable: false));
            baseline.AddSlide();
            _ = baseline.OpenXmlDocument;
            baseline.ClearSlides();
            using var presentation = PowerPointDocument.Load(new MemoryStream(fixture, writable: false));
            PowerPointSlide held = presentation.Slides[0];
            presentation.AddSlide();
            presentation.ClearSlides();
            Assert.Empty(presentation.Slides);
            Assert.True(held.RemovedFromPackage);
            Assert.Equal(baseline.GetSections().Select(section => section.Name), presentation.GetSections().Select(section => section.Name));
            byte[] saved = presentation.ToBytes();
            Assert.False(presentation.IsSdkLoaded);

            using var actual = OfficeOpcPackage.Open(new MemoryStream(saved, writable: false));
            Assert.DoesNotContain(actual.GetRelationships(main), edge => edge.Type == OfficeOpcPackage.SlideRelationshipType);
            const string masterType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "slideMaster";
            Assert.Equal(original.GetRelatedPartName(main, masterType), actual.GetRelatedPartName(main, masterType));
            using var expected = OfficeOpcPackage.Open(new MemoryStream(baseline.ToBytes(), writable: false));
            using var expectedMain = expected.ReadPart(main);
            using var actualMain = actual.ReadPart(main);
            var expectedXml = XElement.Parse(Encoding.UTF8.GetString(expectedMain.Array, 0, expectedMain.Length));
            var actualXml = XElement.Parse(Encoding.UTF8.GetString(actualMain.Array, 0, actualMain.Length));
            XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
            XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            XElement expectedShow = Assert.Single(expectedXml.Element(p + "custShowLst")!.Elements(p + "custShow"));
            XElement actualShow = Assert.Single(actualXml.Element(p + "custShowLst")!.Elements(p + "custShow"));
            Assert.Equal(showName, (string?)actualShow.Attribute("name"));
            Assert.Equal((uint?)expectedShow.Attribute("id"), (uint?)actualShow.Attribute("id"));
            Assert.Equal(expectedShow.Descendants(p + "sld").Select(slide => (string?)slide.Attribute(r + "id")),
                actualShow.Descendants(p + "sld").Select(slide => (string?)slide.Attribute(r + "id")));
            Assert.Contains(actualXml.Descendants(p + "sldId"), element => element.Attributes().Any(attribute => attribute.Value == decoy));

            baseline.AddSlide();
            presentation.AddSlide();
            using var expectedReuse = OfficeOpcPackage.Open(new MemoryStream(baseline.ToBytes(), writable: false));
            using var actualReuse = OfficeOpcPackage.Open(new MemoryStream(presentation.ToBytes(), writable: false));
            OfficeOpcRelationship expectedSlide = Assert.Single(expectedReuse.GetRelationships(main), edge => edge.Type == OfficeOpcPackage.SlideRelationshipType);
            OfficeOpcRelationship actualSlide = Assert.Single(actualReuse.GetRelationships(main), edge => edge.Type == OfficeOpcPackage.SlideRelationshipType);
            Assert.Equal(expectedSlide.Id, actualSlide.Id);
            Assert.Equal(expectedSlide.TargetPartName, actualSlide.TargetPartName);
            Assert.False(presentation.IsSdkLoaded);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void ClearingSlidesKeepsTheTextOfHeldContent(bool holdParagraph) {
            using var source = PowerPointDocument.Create();
            source.AddSlide().AddTextBox("Held text");
            using var presentation = PowerPointDocument.Load(new MemoryStream(source.ToBytes(), writable: false));
            PowerPointTextBox held = Assert.Single(presentation.Slides[0].TextBoxes);
            PowerPointParagraph? paragraph = holdParagraph ? Assert.Single(held.Paragraphs) : null;
            Assert.False(presentation.IsSdkLoaded);

            presentation.ClearSlides();

            Assert.Empty(presentation.Slides);
            Assert.Equal("Held text", held.Text);
            if (paragraph != null) Assert.Equal("Held text", paragraph.Text);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void ClearsTheInternalMainPartAndDeletesItsLastRelationship(bool leadingExternalMain) {
            const string main = "/ppt/presentation.xml";
            const string slide = "/ppt/slides/slide1.xml";
            const int p = OfficeXmlNamespaces.PresentationML;
            const int r = OfficeXmlNamespaces.Relationships;
            using var package = OfficeOpcEditablePackage.Create();
            OfficeXmlTree presentationRoot = package.CreateXmlPart(main,
                "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml", p, "presentation");
            presentationRoot.DeclareNamespace(presentationRoot.Root, "p", "http://schemas.openxmlformats.org/presentationml/2006/main");
            presentationRoot.DeclareNamespace(presentationRoot.Root, "r", OfficeXmlNamespaces.TransitionalRelationshipTypePrefix.TrimEnd('/'));
            int list = PowerPointDocument.Append(presentationRoot, presentationRoot.Root, p, "sldIdLst");
            int entry = PowerPointDocument.Append(presentationRoot, list, p, "sldId");
            presentationRoot.SetAttribute(entry, OfficeXmlNamespaces.None, "id", "256");
            presentationRoot.SetAttribute(entry, r, "id", "slideRel");
            int p14 = package.Namespaces.GetId(PowerPointTreeSections.P14Namespace);
            presentationRoot.DeclareNamespace(presentationRoot.Root, "p14", PowerPointTreeSections.P14Namespace);
            int extensions = PowerPointDocument.Append(presentationRoot, presentationRoot.Root, p, "extLst");
            int extension = PowerPointDocument.Append(presentationRoot, extensions, p, "ext");
            presentationRoot.SetAttribute(extension, OfficeXmlNamespaces.None, "uri", PowerPointTreeSections.SectionListUri);
            int sections = PowerPointDocument.Append(presentationRoot, extension, p14, "sectionLst");
            int section = PowerPointDocument.Append(presentationRoot, sections, p14, "section");
            presentationRoot.SetAttribute(section, OfficeXmlNamespaces.None, "name", "Waiting section");
            presentationRoot.SetAttribute(section, OfficeXmlNamespaces.None, "id", "{61E3A0F1-5F9A-4B60-B953-6C3F35C66ACA}");
            PowerPointDocument.Append(presentationRoot, section, p14, "sldIdLst");
            OfficeXmlTree slideRoot = package.CreateXmlPart(slide,
                "application/vnd.openxmlformats-officedocument.presentationml.slide+xml", p, "sld");
            slideRoot.DeclareNamespace(slideRoot.Root, "p", "http://schemas.openxmlformats.org/presentationml/2006/main");
            int content = PowerPointDocument.Append(slideRoot, slideRoot.Root, p, "cSld");
            PowerPointDocument.Append(slideRoot, content, p, "spTree");
            if (leadingExternalMain) {
                package.CreateXmlPart("/other.xml", "application/xml", p, "presentation");
                package.AddRelationship(null, OfficeOpcPackage.OfficeDocumentRelationshipType, "other.xml", external: true);
            }
            package.AddRelationship(null, OfficeOpcPackage.OfficeDocumentRelationshipType, main);
            package.AddRelationship(main, OfficeOpcPackage.SlideRelationshipType, slide, id: "slideRel");
            using var fixture = new MemoryStream();
            package.Save(fixture);
            fixture.Position = 0;
            using var document = PowerPointDocument.Load(fixture);
            Assert.Single(document.Slides);

            document.ClearSlides();
            Assert.Equal("Waiting section", Assert.Single(document.GetSections()).Name);
            byte[] saved = document.ToBytes();

            Assert.False(document.IsSdkLoaded);
            using var actual = OfficeOpcPackage.Open(new MemoryStream(saved, writable: false));
            Assert.False(actual.PartExists(OfficeOpcPartNames.GetRelationshipsPartName(main)));
            Assert.False(actual.PartExists(slide));
            using var reopened = PowerPointDocument.Load(new MemoryStream(saved, writable: false));
            Assert.Empty(reopened.Slides);
        }

        [Fact]
        public void ClearingAnEmptyDeckKeepsTheSdkErrorForAMissingMaster() {
            using var source = PowerPointDocument.Create();
            byte[] seed = source.ToBytes();
            using var edit = OfficeOpcEditablePackage.Open(new MemoryStream(seed, writable: false));
            const string masterType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "slideMaster";
            edit.AddRelationship("/ppt/presentation.xml", masterType, "/ppt/slideMasters/missing.xml");
            using var output = new MemoryStream();
            edit.Save(output);
            byte[] fixture = output.ToArray();
            using var baseline = PowerPointDocument.Load(new MemoryStream(fixture, writable: false));
            Exception expected = Assert.ThrowsAny<Exception>(() => {
                _ = baseline.OpenXmlDocument;
                baseline.ClearSlides();
                _ = baseline.ToBytes();
            });
            using var presentation = PowerPointDocument.Load(new MemoryStream(fixture, writable: false));
            Assert.IsType(expected.GetType(), Record.Exception(() => {
                presentation.ClearSlides();
                _ = presentation.ToBytes();
            }));
        }

        [Fact]
        public void InspectTemplate_MapsLayoutPlaceholdersThemeAndBrandTokens() {
            string templatePath = CreateTemplatePresentation();
            string targetPath = TempPath(".pptx");
            try {
                PowerPointTemplateInventory inventory = PowerPointTemplate.Inspect(templatePath);

                PowerPointTemplateMasterInfo master = Assert.Single(inventory.Masters);
                Assert.Equal("Corporate Test Theme", master.ThemeName);
                Assert.Equal("0B7FAB", master.ThemeColors[PowerPointThemeColor.Accent1]);
                Assert.Equal("Aptos Display", master.ThemeFonts.MajorLatin);
                Assert.NotEmpty(master.Layouts);
                Assert.Equal(3, inventory.SourceSlideCount);

                PowerPointTemplateLayoutInfo layout = inventory.ResolveLayout(master.Layouts[0].Name);
                Assert.True(layout.SafeArea.Width.ToEmus() > 0);
                PowerPointTemplatePlaceholderInfo body = layout.ResolvePlaceholder("Executive Summary Body");
                Assert.Equal(PowerPointTemplatePlaceholderRole.Body, body.Role);
                Assert.True(body.Bounds.HasValue);
                PowerPointTemplateResolutionException ambiguous = Assert.Throws<PowerPointTemplateResolutionException>(
                    () => layout.ResolvePlaceholder(PowerPointTemplatePlaceholderRole.Body));
                Assert.Equal("Template.PlaceholderAmbiguous", ambiguous.Code);
                Assert.True(ambiguous.Candidates.Count >= 2);
                Assert.Contains("CONFIDENTIAL", inventory.FooterContents);

                PowerPointDesignBrief brief = inventory.CreateDesignBrief("template-brand", "executive review");
                Assert.Equal("0B7FAB", brief.AccentColor);
                Assert.Equal("Aptos Display", brief.HeadingFontFamily);
                Assert.Equal("Aptos", brief.BodyFontFamily);
                Assert.Equal("CONFIDENTIAL", brief.FooterLeft);

                using PowerPointDocument target = PowerPointDocument.Create(targetPath);
                inventory.ApplyBrandTo(target);
                Assert.Equal("Corporate Test Theme", target.ThemeName);
                Assert.Equal("0B7FAB", target.GetThemeColor(PowerPointThemeColor.Accent1));
            } finally {
                Delete(templatePath, targetPath);
            }
        }

        [Fact]
        public void InspectTemplate_InventoriesPicturesNestedInsideGroups() {
            string templatePath = CreateTemplatePresentation();
            try {
                using (PresentationDocument document = PresentationDocument.Open(templatePath, true)) {
                    SlideLayoutPart layoutPart = document.PresentationPart!.SlideMasterParts.First()
                        .SlideLayoutParts.First();
                    ShapeTree tree = layoutPart.SlideLayout.CommonSlideData!.ShapeTree!;
                    tree.Append(new GroupShape(
                        new NonVisualGroupShapeProperties(
                            new NonVisualDrawingProperties { Id = 900U, Name = "Brand Group" },
                            new NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()),
                        new GroupShapeProperties(new A.TransformGroup(
                            new A.Offset { X = 2000000L, Y = 1000000L },
                            new A.Extents { Cx = 2000000L, Cy = 1000000L },
                            new A.ChildOffset { X = 0L, Y = 0L },
                            new A.ChildExtents { Cx = 1000000L, Cy = 500000L })),
                        new DocumentFormat.OpenXml.Presentation.Picture(
                            new NonVisualPictureProperties(
                                new NonVisualDrawingProperties { Id = 901U, Name = "Grouped Logo" },
                                new NonVisualPictureDrawingProperties(),
                                new ApplicationNonVisualDrawingProperties()),
                            new BlipFill(new A.Blip(), new A.Stretch(new A.FillRectangle())),
                            new ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset { X = 100000L, Y = 50000L },
                                    new A.Extents { Cx = 200000L, Cy = 100000L }),
                                new A.PresetGeometry { Preset = A.ShapeTypeValues.Rectangle }))));
                    tree.Append(new DocumentFormat.OpenXml.Presentation.Picture(
                        new NonVisualPictureProperties(
                            new NonVisualDrawingProperties { Id = 912U, Name = "Long Boundary Logo" },
                            new NonVisualPictureDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()),
                        new BlipFill(new A.Blip(), new A.Stretch(new A.FillRectangle())),
                        new ShapeProperties(
                            new A.Transform2D(
                                new A.Offset { X = long.MaxValue, Y = 0L },
                                new A.Extents { Cx = 1L, Cy = 1L }),
                            new A.PresetGeometry { Preset = A.ShapeTypeValues.Rectangle })));
                    layoutPart.SlideLayout.Save();
                }

                PowerPointTemplateInventory inventory = PowerPointTemplate.Inspect(templatePath);
                PowerPointTemplateAssetInfo asset = Assert.Single(inventory.Assets,
                    candidate => candidate.Name == "Grouped Logo" &&
                                 candidate.Kind == PowerPointTemplateAssetKind.Logo);
                PowerPointLayoutBox bounds = Assert.IsType<PowerPointLayoutBox>(asset.Bounds);
                Assert.Equal(2200000L, bounds.Left.ToEmus());
                Assert.Equal(1100000L, bounds.Top.ToEmus());
                Assert.Equal(400000L, bounds.Width.ToEmus());
                Assert.Equal(200000L, bounds.Height.ToEmus());
                Assert.Null(Assert.Single(
                    inventory.Assets,
                    candidate => candidate.Name == "Long Boundary Logo").Bounds);
            } finally {
                Delete(templatePath);
            }
        }

        [Fact]
        public void InspectTemplate_DropsNonFiniteGroupedAssetBoundsInsteadOfOverflowing() {
            string templatePath = CreateTemplatePresentation();
            try {
                using (PresentationDocument document = PresentationDocument.Open(templatePath, true)) {
                    SlideLayoutPart layoutPart = document.PresentationPart!.SlideMasterParts.First()
                        .SlideLayoutParts.First();
                    ShapeTree tree = layoutPart.SlideLayout.CommonSlideData!.ShapeTree!;
                    tree.Append(new GroupShape(
                        new NonVisualGroupShapeProperties(
                            new NonVisualDrawingProperties { Id = 910U, Name = "Extreme Group" },
                            new NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()),
                        new GroupShapeProperties(new A.TransformGroup(
                            new A.Offset { X = long.MaxValue, Y = long.MaxValue },
                            new A.Extents { Cx = long.MaxValue, Cy = long.MaxValue },
                            new A.ChildOffset { X = 0L, Y = 0L },
                            new A.ChildExtents { Cx = 1L, Cy = 1L })),
                        new DocumentFormat.OpenXml.Presentation.Picture(
                            new NonVisualPictureProperties(
                                new NonVisualDrawingProperties { Id = 911U, Name = "Extreme Logo" },
                                new NonVisualPictureDrawingProperties(),
                                new ApplicationNonVisualDrawingProperties()),
                            new BlipFill(new A.Blip(), new A.Stretch(new A.FillRectangle())),
                            new ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset { X = long.MaxValue, Y = long.MaxValue },
                                    new A.Extents { Cx = long.MaxValue, Cy = long.MaxValue }),
                                new A.PresetGeometry { Preset = A.ShapeTypeValues.Rectangle }))));
                    layoutPart.SlideLayout.Save();
                }

                PowerPointTemplateInventory inventory = PowerPointTemplate.Inspect(templatePath);
                PowerPointTemplateAssetInfo asset = Assert.Single(
                    inventory.Assets,
                    candidate => candidate.Name == "Extreme Logo");

                Assert.Null(asset.Bounds);
            } finally {
                Delete(templatePath);
            }
        }

        [Fact]
        public void CreateFromTemplate_RemovesSourceSlidesAndPreservesNamedLayout() {
            string templatePath = CreateTemplatePresentation();
            string outputPath = TempPath(".pptx");
            try {
                PowerPointTemplateInventory inventory = PowerPointTemplate.Inspect(templatePath);
                PowerPointTemplateLayoutInfo layout = inventory.Masters[0].Layouts[0];
                PowerPointTemplatePlaceholderInfo body = layout.ResolvePlaceholder("Executive Summary Body");

                using (PowerPointDocument presentation = PowerPointTemplate.CreateDocument(templatePath,
                           outputPath)) {
                    Assert.Empty(presentation.Slides);
                    PowerPointSlide slide = presentation.AddSlide(layout);
                    slide.AddTextToPlaceholder(body, "Generated from the corporate layout");
                    presentation.Save();
                    Assert.Empty(presentation.ValidateDocument());
                }

                using PowerPointDocument reopened = PowerPointDocument.Load(outputPath);
                Assert.Single(reopened.Slides);
                Assert.Equal("Corporate Test Theme", reopened.ThemeName);
                Assert.NotEmpty(reopened.GetSlideLayouts());
                Assert.Contains(reopened.Slides[0].TextBoxes,
                    textBox => textBox.Text == "Generated from the corporate layout" &&
                               textBox.Name == "Executive Summary Body");
            } finally {
                Delete(templatePath, outputPath);
            }
        }

        [Theory]
        [InlineData(PowerPointTemplateSlideRetention.Selected, false)]
        [InlineData(PowerPointTemplateSlideRetention.Selected, true)]
        [InlineData(PowerPointTemplateSlideRetention.All, true)]
        public void CreateFromTemplate_HidesRetainedSlidesWithoutChangingExtensionMarkup(PowerPointTemplateSlideRetention retention,
            bool useTemplateSubtype) {
            string templatePath = CreateTemplatePresentation();
            string potx = TempPath(".potx");
            string outputPath = TempPath(".pptx");
            try {
                string sourcePath = templatePath;
                if (useTemplateSubtype) {
                    File.Copy(templatePath, potx);
                    ChangeTemplateDocumentTypeThroughSdk(potx, PresentationDocumentType.Template);
                    sourcePath = potx;
                }
                EditTemplateCarrierFixture(sourcePath, "visibility extension");
                byte[] sourceBytes = File.ReadAllBytes(sourcePath);
                var options = new PowerPointTemplateCreationOptions {
                    SlideRetention = retention,
                    HideRetainedSourceSlides = true
                };
                options.SourceSlideIndexes.Add(2);
                options.SourceSlideIndexes.Add(0);

                using (PowerPointDocument presentation = PowerPointTemplate.CreateDocument(sourcePath,
                           outputPath, options)) {
                    Assert.All(presentation.Slides, slide => Assert.True(slide.Hidden));
                    string[] texts = presentation.Slides.Select(slide => Assert.Single(slide.TextBoxes).Text).ToArray();
                    Assert.Equal(retention == PowerPointTemplateSlideRetention.Selected
                        ? new[] { "Source slide 1", "Source slide 3" }
                        : new[] { "Source slide 1", "Source slide 2", "Source slide 3" }, texts);
                }

                Assert.Equal(sourceBytes, File.ReadAllBytes(sourcePath));
                using var package = OfficeOpcPackage.Open(File.OpenRead(outputPath));
                string main = package.GetRelatedPartName(null, OfficeOpcPackage.OfficeDocumentRelationshipType)!;
                string slidePart = package.GetRelatedPartName(main, OfficeOpcPackage.SlideRelationshipType)!;
                XElement root = ReadTemplateXml(package, slidePart);
                XNamespace p = root.Name.Namespace;
                XElement decoy = Assert.Single(root.Element(p + "extLst")!.Descendants(p + "sld"));
                Assert.Equal("1", (string?)decoy.Attribute("show"));
            } finally {
                Delete(templatePath, potx, outputPath);
            }
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void CreateFromTemplate_DoesNotExposeCopiedTemplateWhenRetentionValidationFails(bool useTemplateSubtype) {
            string templatePath = CreateTemplatePresentation();
            string potx = TempPath(".potx");
            string outputPath = TempPath(".pptx");
            try {
                if (useTemplateSubtype) {
                    File.Copy(templatePath, potx);
                    ChangeTemplateDocumentTypeThroughSdk(potx, PresentationDocumentType.Template);
                }
                var options = new PowerPointTemplateCreationOptions {
                    SlideRetention = PowerPointTemplateSlideRetention.Selected
                };
                options.SourceSlideIndexes.Add(99);

                Assert.Throws<ArgumentOutOfRangeException>(() =>
                    PowerPointTemplate.CreateDocument(useTemplateSubtype ? potx : templatePath, outputPath, options));

                Assert.False(File.Exists(outputPath));
                string directory = Path.GetDirectoryName(outputPath)!;
                string prefix = "." + Path.GetFileNameWithoutExtension(outputPath) + ".";
                Assert.DoesNotContain(Directory.EnumerateFiles(directory), path =>
                    Path.GetFileName(path).StartsWith(prefix, StringComparison.Ordinal));
            } finally {
                Delete(templatePath, potx, outputPath);
            }
        }

        [Theory]
        [InlineData(null)]
        [InlineData("external relationship")]
        [InlineData("reserved main name")]
        [InlineData("mixed relationship types")]
        [InlineData("discarded Strict slide")]
        [InlineData("unbound content type")]
        [InlineData("invalid content type catalog")]
        [InlineData("invalid media type")]
        public void CreateFromTemplate_ConvertsPotxToEditablePptx(string? carrierCase) {
            string sourcePptx = CreateTemplatePresentation();
            string sourcePotx = TempPath(".potx");
            string outputPath = TempPath(".pptx");
            string sdkCopy = TempPath(".pptx");
            try {
                File.Copy(sourcePptx, sourcePotx);
                ChangeTemplateDocumentTypeThroughSdk(sourcePotx, PresentationDocumentType.Template);
                if (carrierCase != null) EditTemplateCarrierFixture(sourcePotx, carrierCase);
                File.Copy(sourcePotx, sdkCopy);
                if (carrierCase is "invalid content type catalog" or "invalid media type" or "discarded Strict slide") {
                    Exception expectedError = Assert.ThrowsAny<Exception>(() => ChangeTemplateDocumentTypeThroughSdk(sdkCopy, PresentationDocumentType.Presentation));
                    Assert.IsType(expectedError.GetType(), Record.Exception(() => PowerPointTemplate.CreateDocument(sourcePotx, outputPath).Dispose()));
                    Assert.False(File.Exists(outputPath));
                    return;
                }
                ChangeTemplateDocumentTypeThroughSdk(sdkCopy, PresentationDocumentType.Presentation);
                bool retainAll = carrierCase is "unbound content type" or "mixed relationship types";
                using (PowerPointDocument reference = PowerPointDocument.Load(sdkCopy)) {
                    if (!retainAll) reference.ClearSlides();
                    reference.Save();
                }
                var options = new PowerPointTemplateCreationOptions {
                    SlideRetention = retainAll ? PowerPointTemplateSlideRetention.All : PowerPointTemplateSlideRetention.None
                };

                using (PowerPointDocument presentation = PowerPointTemplate.CreateDocument(sourcePotx,
                           outputPath, options)) {
                    Assert.Equal(retainAll ? 3 : 0, presentation.Slides.Count);
                    using (var expected = OfficeOpcPackage.Open(File.OpenRead(sdkCopy)))
                    using (var actual = OfficeOpcPackage.Open(File.OpenRead(outputPath))) {
                        OfficeOpcRelationship expectedMain = Assert.Single(expected.GetRelationships(null),
                            edge => edge.Type == OfficeOpcPackage.OfficeDocumentRelationshipType && !edge.IsExternal);
                        OfficeOpcRelationship actualMain = Assert.Single(actual.GetRelationships(null),
                            edge => edge.Type == OfficeOpcPackage.OfficeDocumentRelationshipType && !edge.IsExternal);
                        Assert.Equal(expectedMain.Id, actualMain.Id);
                        Assert.Equal(expectedMain.TargetPartName, actualMain.TargetPartName);
                        Assert.Equal(expected.GetRelationships(null).Where(edge => edge.IsExternal).Select(edge => (edge.Id, edge.Type, edge.Target)),
                            actual.GetRelationships(null).Where(edge => edge.IsExternal).Select(edge => (edge.Id, edge.Type, edge.Target)));
                        if (carrierCase == "mixed relationship types") {
                            string slide = expected.GetRelatedPartName(expectedMain.TargetPartName, OfficeOpcPackage.SlideRelationshipType)!;
                            string part = OfficeOpcPartNames.GetRelationshipsPartName(slide);
                            Assert.Equal(ReadTemplateXml(expected, part).Elements().Select(edge => (string?)edge.Attribute("Type")),
                                ReadTemplateXml(actual, part).Elements().Select(edge => (string?)edge.Attribute("Type")));
                        }
                        if (carrierCase == "unbound content type") {
                            Assert.Contains(ReadTemplateXml(expected, "/[Content_Types].xml").Elements(), item => (string?)item.Attribute("PartName") == "/ppt/ghost.xml");
                            Assert.Contains(ReadTemplateXml(actual, "/[Content_Types].xml").Elements(), item => (string?)item.Attribute("PartName") == "/ppt/ghost.xml");
                        }
                    }
                    presentation.AddSlide();
                    presentation.Save();
                    Assert.Empty(presentation.ValidateDocument());
                }

                using PresentationDocument package = PresentationDocument.Open(outputPath, false);
                Assert.Equal(PresentationDocumentType.Presentation, package.DocumentType);
            } finally {
                Delete(sourcePptx, sourcePotx, outputPath, sdkCopy);
            }
        }

        private static void ChangeTemplateDocumentTypeThroughSdk(string path, PresentationDocumentType type) {
            using Stream stream = File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            using var document = PresentationDocument.Open(stream, true);
            document.ChangeDocumentType(type);
            document.Save();
        }

        [Theory]
        [InlineData("discarded Strict slide")]
        [InlineData("absolute internal hyperlink")]
        [InlineData("internal hyperlink to relationships")]
        [InlineData("internal hyperlink to directory")]
        public void CreateFromTemplate_ChecksRelationshipsBeforeDiscardingUnselectedSlides(string carrierCase) {
            string templatePath = CreateTemplatePresentation();
            string potx = TempPath(".potx");
            string reference = TempPath(".pptx");
            string outputPath = TempPath(".pptx");
            try {
                File.Copy(templatePath, potx);
                ChangeTemplateDocumentTypeThroughSdk(potx, PresentationDocumentType.Template);
                EditTemplateCarrierFixture(potx, carrierCase);
                File.Copy(potx, reference);
                Exception expected = Assert.ThrowsAny<Exception>(() =>
                    ChangeTemplateDocumentTypeThroughSdk(reference, PresentationDocumentType.Presentation));
                byte[] existingDestination = File.ReadAllBytes(templatePath);
                File.WriteAllBytes(outputPath, existingDestination);
                var options = new PowerPointTemplateCreationOptions {
                    SlideRetention = PowerPointTemplateSlideRetention.Selected,
                    Overwrite = true
                };
                options.SourceSlideIndexes.Add(1);
                Assert.IsType(expected.GetType(), Record.Exception(() => PowerPointTemplate.CreateDocument(potx, outputPath, options).Dispose()));
                Assert.Equal(existingDestination, File.ReadAllBytes(outputPath));
            } finally {
                Delete(templatePath, potx, reference, outputPath);
            }
        }

        private static XElement ReadTemplateXml(OfficeOpcPackage package, string part) {
            using var bytes = package.ReadPart(part);
            using var stream = new MemoryStream(bytes.Array, 0, bytes.Length, writable: false);
            return XElement.Load(stream);
        }

        private static void EditTemplateCarrierFixture(string path, string carrierCase) {
            byte[] data = File.ReadAllBytes(path);
            using var original = OfficeOpcPackage.Open(new MemoryStream(data, writable: false));
            string main = original.GetRelatedPartName(null, OfficeOpcPackage.OfficeDocumentRelationshipType)!;
            string slide = original.GetRelatedPartName(main, OfficeOpcPackage.SlideRelationshipType)!;
            using var editable = OfficeOpcEditablePackage.Open(new MemoryStream(data, writable: false));
            switch (carrierCase) {
                case "absolute internal hyperlink":
                case "internal hyperlink to relationships":
                case "internal hyperlink to directory":
                    string target = carrierCase switch {
                        "absolute internal hyperlink" => "https://example.invalid/template",
                        "internal hyperlink to directory" => ".",
                        _ => OfficeOpcPartNames.GetRelationshipsPartName(slide)
                    };
                    editable.AddRelationship(slide, PowerPointPartRef.HyperlinkRelationshipType, target);
                    break;
                case "visibility extension":
                    OfficeXmlTree slideTree = editable.GetXml(slide);
                    const int p = OfficeXmlNamespaces.PresentationML;
                    int extList = slideTree.FindChild(slideTree.Root, p, "extLst");
                    if (extList == OfficeXmlTree.NoNode) extList = PowerPointDocument.Append(slideTree, slideTree.Root, p, "extLst");
                    int extension = PowerPointDocument.Append(slideTree, extList, p, "ext");
                    slideTree.SetAttribute(extension, OfficeXmlNamespaces.None, "uri", "{A9E23632-C341-4775-A681-3E73E42A92EE}");
                    int decoy = PowerPointDocument.Append(slideTree, extension, p, "sld");
                    slideTree.SetAttribute(decoy, OfficeXmlNamespaces.None, "show", "1");
                    break;
                case "external relationship":
                    editable.AddRelationship(null, "urn:officeimo:template-reference", "https://example.invalid/template", external: true);
                    break;
                case "reserved main name":
                    string occupied = main.Insert(main.LastIndexOf('.'), "2");
                    editable.CreateBinaryPart(occupied, "application/octet-stream", new byte[] { 1 });
                    editable.AddRelationship(slide, OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "package", occupied);
                    break;
                case "discarded Strict slide":
                case "mixed relationship types":
                    OfficeXmlTree relationships = editable.GetRelationships(slide);
                    for (int edge = relationships.FindChild(relationships.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship");
                         edge != OfficeXmlTree.NoNode; edge = relationships.FindNextSibling(edge, OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                        if (relationships.GetAttribute(edge, OfficeXmlNamespaces.None, "Type") == OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "slideLayout") {
                            relationships.SetAttribute(edge, OfficeXmlNamespaces.None, "Type", "http://purl.oclc.org/ooxml/officeDocument/relationships/slideLayout");
                        }
                    }
                    if (carrierCase == "discarded Strict slide") {
                        byte[] malformed = Encoding.UTF8.GetBytes("<p:sld xmlns:p='http://schemas.openxmlformats.org/presentationml/2006/main'><p:cSld>");
                        editable.ReplacePart(slide, malformed, malformed.Length);
                    }
                    break;
                case "unbound content type":
                    editable.SetOverride("/ppt/ghost.xml", "application/xml");
                    break;
                case "invalid content type catalog":
                    editable.ContentTypes.SetAttribute(editable.ContentTypes.Root, OfficeXmlNamespaces.None, "unexpected", "1");
                    break;
                case "invalid media type":
                    editable.AddDefaultContentType("bogus", "invalidmime");
                    break;
            }
            using var modified = new MemoryStream();
            editable.Save(modified);
            File.WriteAllBytes(path, modified.ToArray());
        }

        [Fact]
        public void TemplateDesigner_RendersSemanticPlanIntoMappedNamedLayout() {
            string templatePath = CreateTemplatePresentation();
            string outputPath = TempPath(".pptx");
            try {
                PowerPointTemplateInventory inventory = PowerPointTemplate.Inspect(templatePath);
                PowerPointTemplateLayoutInfo namedLayout = inventory.Masters[0].Layouts[0];
                var map = new PowerPointTemplateLayoutMap()
                    .Map(PowerPointDeckPlanSlideKind.Process, namedLayout);
                var plan = new PowerPointDeckPlan().AddProcess("Mapped process", null, new[] {
                    new PowerPointProcessStep("Discover", "Collect constraints."),
                    new PowerPointProcessStep("Deliver", "Ship with evidence.")
                });

                using (PowerPointDocument presentation = PowerPointTemplate.CreateDocument(templatePath,
                           outputPath)) {
                    PowerPointDeckComposer deck = presentation.UseTemplateDesigner(inventory, map,
                        "template-render", "delivery plan");
                    PowerPointSlide rendered = Assert.Single(deck.AddSlides(plan));
                    Assert.Equal(namedLayout.LayoutIndex, rendered.LayoutIndex);
                    Assert.Equal("Corporate Test Theme", presentation.ThemeName);
                    presentation.Save();
                    Assert.Empty(presentation.ValidateDocument());
                }

                using PowerPointDocument reopened = PowerPointDocument.Load(outputPath);
                PowerPointSlide slide = Assert.Single(reopened.Slides);
                Assert.Equal(namedLayout.LayoutIndex, slide.LayoutIndex);
                Assert.Contains(slide.TextBoxes, box => box.Text == "Mapped process");
            } finally {
                Delete(templatePath, outputPath);
            }
        }

        [Fact]
        public void ExistingPptxFixture_PreservesMastersLayoutsAndSlidesThroughTemplateCopy() {
            string fixturePath = Path.Combine(RepositoryTestPaths.Find(),
                "Assets", "PowerPointTemplates", "PowerPointWithTitle.pptx");
            Assert.True(File.Exists(fixturePath), "Expected PowerPoint template fixture at " + fixturePath);
            string outputPath = TempPath(".pptx");
            try {
                PowerPointTemplateInventory source = PowerPointTemplate.Inspect(fixturePath);
                var options = new PowerPointTemplateCreationOptions {
                    SlideRetention = PowerPointTemplateSlideRetention.All
                };
                using (PowerPointDocument copied = PowerPointTemplate.CreateDocument(fixturePath,
                           outputPath, options)) {
                    PowerPointTemplateInventory result = PowerPointTemplate.Inspect(copied);
                    Assert.Equal(source.SourceSlideCount, result.SourceSlideCount);
                    Assert.Equal(source.Masters.Count, result.Masters.Count);
                    Assert.Equal(source.Masters.Sum(master => master.Layouts.Count),
                        result.Masters.Sum(master => master.Layouts.Count));
                    copied.Save();
                    Assert.Empty(copied.ValidateDocument());
                }
            } finally {
                Delete(outputPath);
            }
        }

        private static string CreateTemplatePresentation() {
            string path = TempPath(".pptx");
            using PowerPointDocument presentation = PowerPointDocument.Create(path);
            presentation.ThemeName = "Corporate Test Theme";
            presentation.SetThemeColor(PowerPointThemeColor.Accent1, "0B7FAB");
            presentation.SetThemeColor(PowerPointThemeColor.Accent2, "7C3AED");
            presentation.SetThemeColor(PowerPointThemeColor.Light2, "F5F7FA");
            presentation.SetThemeLatinFonts("Aptos Display", "Aptos");

            PowerPointTextBox primary = presentation.EnsureLayoutPlaceholderTextBox(0, 0,
                PowerPointPlaceholderType.Body, 20, new OfficeIMO.PowerPoint.PowerPointLayoutBox(OfficeIMO.OfficeLength.Centimeters(1.5),OfficeIMO.OfficeLength.Centimeters(3.0),OfficeIMO.OfficeLength.Centimeters(18),OfficeIMO.OfficeLength.Centimeters(7)),
                "Executive Summary Body");
            primary.Name = "Executive Summary Body";
            PowerPointTextBox secondary = presentation.EnsureLayoutPlaceholderTextBox(0, 0,
                PowerPointPlaceholderType.Body, 21, new OfficeIMO.PowerPoint.PowerPointLayoutBox(OfficeIMO.OfficeLength.Centimeters(20),OfficeIMO.OfficeLength.Centimeters(3.0),OfficeIMO.OfficeLength.Centimeters(5),OfficeIMO.OfficeLength.Centimeters(7)),
                "Supporting Body");
            secondary.Name = "Supporting Body";
            presentation.EnsureLayoutFooterPlaceholderTextBox(0, 0, "CONFIDENTIAL");

            for (int index = 1; index <= 3; index++) {
                PowerPointSlide slide = presentation.AddSlide();
                slide.AddTextBox("Source slide " + index,global::OfficeIMO.OfficeLength.Centimeters(2),global::OfficeIMO.OfficeLength.Centimeters(2),global::OfficeIMO.OfficeLength.Centimeters(10),global::OfficeIMO.OfficeLength.Centimeters(1));
            }
            presentation.Save();
            return path;
        }

        private static string TempPath(string extension) =>
            Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + extension);

        private static void Delete(params string[] paths) {
            foreach (string path in paths) if (File.Exists(path)) File.Delete(path);
        }
    }
}
