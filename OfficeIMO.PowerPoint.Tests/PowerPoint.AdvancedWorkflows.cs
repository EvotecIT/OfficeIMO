using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Tests.Pdf;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using Dgm = DocumentFormat.OpenXml.Drawing.Diagrams;
using P188 = DocumentFormat.OpenXml.Office2021.PowerPoint.Comment;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests {
    public class PowerPointAdvancedWorkflowTests {
        [Fact]
        public void ReviewAndAnimationInspectionProjectsClassicModernAndTimingMetadata() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream, new PowerPointCreateOptions());
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTextBox target = slide.AddTextBoxPoints("Animated review target", 40, 40, 240, 50);

            PresentationPart presentationPart = slide.SlidePart.GetParentParts().OfType<PresentationPart>().Single();
            CommentAuthorsPart classicAuthors = presentationPart.AddNewPart<CommentAuthorsPart>();
            classicAuthors.CommentAuthorList = new CommentAuthorList(
                new CommentAuthor {
                    Id = 0U, Name = "Classic Reviewer", Initials = "CR", LastIndex = 1U, ColorIndex = 0U
                });
            SlideCommentsPart classicPart = slide.SlidePart.AddNewPart<SlideCommentsPart>();
            classicPart.CommentList = new CommentList(
                new Comment(
                    new Position { X = 120, Y = 240 },
                    new DocumentFormat.OpenXml.Presentation.Text("Classic review")) {
                    AuthorId = 0U,
                    Index = 1U,
                    DateTime = new DateTime(2026, 7, 10, 8, 0, 0, DateTimeKind.Utc)
                });

            string modernAuthorId = "{11111111-1111-1111-1111-111111111111}";
            string modernCommentId = "{22222222-2222-2222-2222-222222222222}";
            string modernReplyId = "{33333333-3333-3333-3333-333333333333}";
            PowerPointAuthorsPart modernAuthors = presentationPart.AddNewPart<PowerPointAuthorsPart>();
            FeedXml(modernAuthors, $"""
                <p188:authorLst xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main">
                  <p188:author id="{modernAuthorId}" name="Modern Reviewer" initials="MR" userId="reviewer@example.test" providerId="OfficeIMO" />
                </p188:authorLst>
                """);
            PowerPointCommentPart modernPart = slide.SlidePart.AddNewPart<PowerPointCommentPart>();
            FeedXml(modernPart, $"""
                <p188:cmLst xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p188:cm id="{modernCommentId}" authorId="{modernAuthorId}" status="active" created="2026-07-10T08:05:00Z">
                    <p188:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Modern review</a:t></a:r></a:p></p188:txBody>
                    <p188:replyLst>
                      <p188:reply id="{modernReplyId}" authorId="{modernAuthorId}" status="active" created="2026-07-10T08:06:00Z">
                        <p188:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Reply review</a:t></a:r></a:p></p188:txBody>
                      </p188:reply>
                    </p188:replyLst>
                  </p188:cm>
                </p188:cmLst>
                """);

            slide.SlidePart.Slide!.Timing = new Timing(
                new TimeNodeList(
                    new ParallelTimeNode(
                        new CommonTimeNode(
                            new ChildTimeNodeList(
                                new Animate(
                                    new CommonBehavior(
                                        new CommonTimeNode { Id = 2U, Duration = "500" },
                                        new TargetElement(new ShapeTarget {
                                            ShapeId = target.Id!.Value.ToString(CultureInfo.InvariantCulture)
                                        }))))) {
                            Id = 1U,
                            Duration = "indefinite",
                            NodeType = TimeNodeValues.TmingRoot
                        })));

            PowerPointReviewReport review = presentation.InspectReviewComments();
            PowerPointAnimationReport animation = presentation.InspectAnimations();

            Assert.Equal(1, review.ClassicCount);
            Assert.Equal(2, review.ModernCount);
            Assert.Contains(review.Comments, comment => comment.AuthorName == "Classic Reviewer" &&
                comment.Text == "Classic review");
            Assert.True(review.Comments.Any(comment => comment.AuthorName == "Modern Reviewer" &&
                comment.Text == "Modern review" && string.Equals(comment.Status, "Active",
                    StringComparison.OrdinalIgnoreCase)), review.ToJson());
            PowerPointReviewComment reply = Assert.Single(review.Comments,
                comment => comment.Kind == PowerPointCommentKind.ModernReply);
            Assert.Equal(modernCommentId, reply.ParentId);
            Assert.Equal("Reply review", reply.Text);
            Assert.Contains("\"commentCount\":3", review.ToJson(), StringComparison.Ordinal);
            PowerPointAnimationNode animated = Assert.Single(animation.Nodes,
                node => node.Kind == PowerPointAnimationKind.Animate);
            PowerPointAnimationNode container = Assert.Single(animation.Nodes,
                node => node.Kind == PowerPointAnimationKind.Parallel);
            Assert.Equal("1", container.TimingId);
            Assert.Null(container.ShapeId);
            Assert.Null(container.ShapeName);
            Assert.Equal(target.Id, animated.ShapeId);
            Assert.Equal(target.Name, animated.ShapeName);
            Assert.Equal("500", animated.Duration);
        }

        [Fact]
        public void AnimationInspectionBoundsTraversalAndProjectedNodes() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            PowerPointSlide slide = presentation.AddSlide();
            slide.SlidePart.Slide!.Timing = new Timing(
                new TimeNodeList(
                    new ParallelTimeNode(new CommonTimeNode { Id = 1U }),
                    new ParallelTimeNode(new CommonTimeNode { Id = 2U })));

            InvalidDataException nodeException = Assert.Throws<InvalidDataException>(() =>
                presentation.InspectAnimations(new PowerPointAnimationInspectionOptions { MaxAnimationNodes = 1 }));
            InvalidDataException elementException = Assert.Throws<InvalidDataException>(() =>
                presentation.InspectAnimations(new PowerPointAnimationInspectionOptions { MaxXmlElements = 1 }));

            Assert.Contains(nameof(PowerPointAnimationInspectionOptions.MaxAnimationNodes), nodeException.Message, StringComparison.Ordinal);
            Assert.Contains(nameof(PowerPointAnimationInspectionOptions.MaxXmlElements), elementException.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void AnimationInspectionSkipsShapeTraversalForUntargetedTimingNodes() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(path)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    ShapeTree shapeTree = slide.SlidePart.Slide!.CommonSlideData!.ShapeTree!;
                    for (uint shapeId = 100U; shapeId < 10_101U; shapeId++) {
                        shapeTree.Append(new Shape(
                            new NonVisualShapeProperties(
                                new NonVisualDrawingProperties { Id = shapeId, Name = "Shape " + shapeId },
                                new NonVisualShapeDrawingProperties(),
                                new ApplicationNonVisualDrawingProperties()),
                            new ShapeProperties()));
                    }
                    slide.SlidePart.Slide.Timing = new Timing(
                        new TimeNodeList(new ParallelTimeNode(new CommonTimeNode { Id = 1U })));
                    presentation.Save();
                }

                using PowerPointPresentation reopened = PowerPointPresentation.Load(path);
                Assert.True(reopened.Slides[0].Shapes.Count > 10_000);
                PowerPointAnimationNode node = Assert.Single(reopened.InspectAnimations().Nodes);

                Assert.Equal(PowerPointAnimationKind.Parallel, node.Kind);
                Assert.Null(node.ShapeId);
                Assert.Null(node.ShapeName);
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void SemanticSmartArtWorkflowsRoundTripEditableNodeText() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(path)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointSmartArt process = slide.AddSmartArt(PowerPointSmartArtType.BasicProcess,
                        new[] { "Discover", "Design", "Deliver" }, 20, 20, 2600000, 1200000);
                    PowerPointSmartArt hierarchy = slide.AddSmartArt(PowerPointSmartArtType.BasicHierarchy,
                        new[] { "Executive", "Platform", "Delivery" }, 2800000, 20, 2600000, 1200000);
                    PowerPointSmartArt cycle = slide.AddSmartArt(PowerPointSmartArtType.BasicCycle,
                        new[] { "Plan", "Run", "Learn", "Improve" }, 20, 1400000, 5200000, 1400000);

                    Assert.Equal(new[] { "Discover", "Design", "Deliver" }, process.GetNodeTexts());
                    Assert.Equal(3, hierarchy.NodeCount);
                    Assert.Equal(4, cycle.NodeCount);
                    cycle.SetNodeText(3, "Adapt");
                    Assert.Equal("Adapt", cycle.GetNodeText(3));
                    presentation.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(path, new PowerPointLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                    PowerPointSmartArt[] diagrams = presentation.Slides[0].SmartArts.ToArray();
                    Assert.Equal(3, diagrams.Length);
                    Assert.Equal(new[] { "Discover", "Design", "Deliver" }, diagrams[0].GetNodeTexts());
                    Assert.Equal(new[] { "Plan", "Run", "Learn", "Adapt" }, diagrams[2].GetNodeTexts());
                    var validation = presentation.ValidateDocument();
                    Assert.True(validation.Count == 0, string.Join(Environment.NewLine,
                        validation.Select(error => error.Description + " | " + error.Path)));
                }
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void BasicCycleAuthorsEveryExportedTopologyEdge() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            slide.AddSmartArt(PowerPointSmartArtType.BasicCycle,
                new[] { "Plan", "Run", "Learn", "Improve" });

            Dgm.LayoutDefinition layout = Assert.Single(slide.SlidePart
                .DiagramLayoutDefinitionParts).LayoutDefinition!;
            Dgm.ForEach transitions = Assert.Single(
                layout.Descendants<Dgm.ForEach>(), iterator =>
                    iterator.Name?.Value == "cycleTransitions");
            Assert.Equal("followSib", transitions.Axis?.InnerText);
            Assert.Equal("sibTrans", transitions.PointType?.InnerText);
            Assert.Equal("0", transitions.HideLastTrans?.InnerText);
            Dgm.LayoutNode connector = Assert.Single(
                transitions.Elements<Dgm.LayoutNode>());
            Assert.Equal(Dgm.AlgorithmValues.Connector,
                connector.GetFirstChild<Dgm.Algorithm>()!.Type!.Value);
            Assert.Equal("conn",
                connector.GetFirstChild<Dgm.Shape>()!.Type!.Value);

            PowerPointSlideVisualSnapshot snapshot =
                slide.CreateVisualSnapshot();
            Assert.Equal(4, snapshot.Drawing.Shapes.Count(shape =>
                shape.Shape.Kind == OfficeShapeKind.Line));
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void BasicHierarchyAuthorsEveryParentTransition() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            slide.AddSmartArt(PowerPointSmartArtType.BasicHierarchy,
                new[] { "Executive", "Platform", "Delivery", "Operations" });

            Dgm.LayoutDefinition layout = Assert.Single(slide.SlidePart
                .DiagramLayoutDefinitionParts).LayoutDefinition!;
            Dgm.ForEach[] transitions = layout.Descendants<Dgm.ForEach>()
                .Where(iterator => iterator.Name?.Value?.StartsWith(
                    "hierarchyTransition", StringComparison.Ordinal) == true)
                .ToArray();
            Assert.Equal(3, transitions.Length);
            Assert.All(transitions, transition => {
                Assert.Equal("ch ch self", transition.Axis?.InnerText);
                Assert.Equal("node node parTrans",
                    transition.PointType?.InnerText);
                Dgm.LayoutNode connector = Assert.Single(
                    transition.Elements<Dgm.LayoutNode>());
                Assert.Equal(Dgm.AlgorithmValues.Connector,
                    connector.GetFirstChild<Dgm.Algorithm>()!.Type!.Value);
                Assert.Equal("conn",
                    connector.GetFirstChild<Dgm.Shape>()!.Type!.Value);
            });

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();
            Assert.Equal(3, snapshot.Drawing.Shapes.Count(shape =>
                shape.Shape.Kind == OfficeShapeKind.Line));
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void AddSmartArtRejectsUndefinedLayoutKind() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();

            ArgumentOutOfRangeException exception = Assert.Throws<ArgumentOutOfRangeException>(() =>
                slide.AddSmartArt((PowerPointSmartArtType)int.MaxValue,
                    new[] { "Discover", "Deliver" }));

            Assert.Equal("type", exception.ParamName);
            Assert.Empty(slide.SmartArts);
        }

        [Theory]
        [InlineData(27273042316901L, 3200400L, "width")]
        [InlineData(5486400L, 27273042316901L, "height")]
        public void AddSmartArtRejectsUnrepresentableExtentsBeforeMutation(
            long width, long height, string parameterName) {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();

            ArgumentOutOfRangeException exception =
                Assert.Throws<ArgumentOutOfRangeException>(() =>
                    slide.AddSmartArt(PowerPointSmartArtType.BasicProcess,
                        new[] { "Discover", "Deliver" }, width: width,
                        height: height));

            Assert.Equal(parameterName, exception.ParamName);
            Assert.Empty(slide.SmartArts);
            Assert.Empty(slide.SlidePart.DiagramDataParts);
            Assert.Empty(slide.SlidePart.DiagramLayoutDefinitionParts);
        }

        [Fact]
        public void ImportedSmartArtRejectsUnknownCategoryBeforeSemanticProjection() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            presentation.SlideSize.SetSizePoints(640, 360);
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointSmartArt smartArt = slide.AddSmartArt(
                PowerPointSmartArtType.BasicProcess,
                new[] { "Discover", "Deliver" });
            DiagramDataPart dataPart = Assert.Single(slide.SlidePart
                .DiagramDataParts);
            XDocument data;
            using (Stream stream = dataPart.GetStream(FileMode.Open,
                       FileAccess.Read)) {
                data = XDocument.Load(stream);
            }
            XNamespace dgm =
                "http://schemas.openxmlformats.org/drawingml/2006/diagram";
            XElement properties = data.Descendants(dgm + "prSet")
                .Single(element => element.Attribute("loCatId") != null);
            properties.SetAttributeValue("loCatId", "picture");
            using (Stream stream = dataPart.GetStream(FileMode.Create,
                       FileAccess.Write)) {
                data.Save(stream);
            }

            Assert.False(smartArt.TryGetOfficeDiagramSnapshot(out _));
            Assert.True(PowerPointDesktopReferenceRenderer
                .HasExpectedVisibleContent(slide));

            smartArt.LeftPoints = 700;
            Assert.False(PowerPointDesktopReferenceRenderer
                .HasExpectedVisibleContent(slide));
        }

        [Fact]
        public void ImportedSmartArtRejectsModifiedRecognizedLayoutBeforeProjection() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSmartArt smartArt = presentation.AddSlide().AddSmartArt(
                PowerPointSmartArtType.BasicProcess,
                new[] { "Discover", "Deliver" });
            Dgm.LayoutDefinition layout = Assert.Single(presentation.Slides[0]
                .SlidePart.DiagramLayoutDefinitionParts).LayoutDefinition!;
            Dgm.Shape shape = layout.Descendants<Dgm.Shape>().First();
            shape.Type = "ellipse";

            Assert.False(smartArt.TryGetOfficeDiagramSnapshot(out _));
        }

        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public void ImportedSmartArtRejectsModifiedRecognizedStyleParts(
            bool modifyQuickStyle) {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointSmartArt smartArt = slide.AddSmartArt(
                PowerPointSmartArtType.BasicProcess,
                new[] { "Discover", "Deliver" });
            OpenXmlElement definition = modifyQuickStyle
                ? Assert.Single(slide.SlidePart.DiagramStyleParts)
                    .StyleDefinition!
                : Assert.Single(slide.SlidePart.DiagramColorsParts)
                    .ColorsDefinition!;
            OpenXmlElement title = definition.ChildElements[0];
            title.SetAttribute(new OpenXmlAttribute("val", string.Empty,
                "Producer modified"));

            Assert.False(smartArt.TryGetOfficeDiagramSnapshot(out _));
        }

        [Fact]
        public void ImportedSmartArtRejectsUnrepresentableQuickAndColorStyles() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointSmartArt smartArt = slide.AddSmartArt(
                PowerPointSmartArtType.BasicProcess,
                new[] { "Discover", "Deliver" });
            DiagramDataPart dataPart = Assert.Single(slide.SlidePart
                .DiagramDataParts);
            XDocument data;
            using (Stream stream = dataPart.GetStream(FileMode.Open,
                       FileAccess.Read)) {
                data = XDocument.Load(stream);
            }
            XNamespace dgm =
                "http://schemas.openxmlformats.org/drawingml/2006/diagram";
            XElement properties = data.Descendants(dgm + "prSet")
                .Single(element => element.Attribute("loCatId") != null);
            properties.SetAttributeValue("qsTypeId",
                "urn:vendor:quickstyle:custom");
            properties.SetAttributeValue("csTypeId",
                "urn:vendor:colors:custom");
            using (Stream stream = dataPart.GetStream(FileMode.Create,
                       FileAccess.Write)) {
                data.Save(stream);
            }

            Assert.False(smartArt.TryGetOfficeDiagramSnapshot(out _));
        }

        [Fact]
        public void SemanticSmartArtCarriesThemeColorsAndMinorFont() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            presentation.SetThemeColor(PowerPointThemeColor.Accent1,
                "C00000");
            presentation.SetThemeColor(PowerPointThemeColor.Light1,
                "F5F5F5");
            presentation.SetThemeLatinFonts("Georgia", "Aptos");
            PowerPointSmartArt smartArt = presentation.AddSlide().AddSmartArt(
                PowerPointSmartArtType.BasicProcess,
                new[] { "Discover", "Deliver" });

            Assert.True(smartArt.TryGetOfficeDiagramSnapshot(
                out OfficeDiagramSnapshot snapshot));
            Assert.NotNull(snapshot.Style);
            Assert.Equal("Aptos", snapshot.Style!.FontFamily);
            Assert.Equal("C00000", snapshot.Style.NodeColors.Single()
                .ToRgbHex());
            Assert.Equal("F5F5F5", snapshot.Style.NodeTextColor.ToRgbHex());
        }

        [Theory]
        [InlineData(PowerPointSmartArtType.BasicHierarchy, OfficeDiagramKind.Hierarchy)]
        [InlineData(PowerPointSmartArtType.BasicList, OfficeDiagramKind.List)]
        [InlineData(PowerPointSmartArtType.BasicMatrix, OfficeDiagramKind.Matrix)]
        [InlineData(PowerPointSmartArtType.BasicPyramid, OfficeDiagramKind.Pyramid)]
        [InlineData(PowerPointSmartArtType.BasicRelationship, OfficeDiagramKind.Relationship)]
        public void BroaderSemanticSmartArtLayoutsRoundTripAndRender(
            PowerPointSmartArtType type, OfficeDiagramKind expectedKind) {
            using var stream = new MemoryStream();
            using (PowerPointPresentation presentation = PowerPointPresentation.Create(stream)) {
                presentation.SlideSize.SetSizePoints(360, 220);
                PowerPointSmartArt smartArt = presentation.AddSlide().AddSmartArt(type,
                    new[] { "Discover", "Build", "Validate", "Ship" },
                    PowerPointUnits.FromPoints(20),
                    PowerPointUnits.FromPoints(20),
                    PowerPointUnits.FromPoints(320),
                    PowerPointUnits.FromPoints(180));
                Assert.True(smartArt.TryGetOfficeDiagramSnapshot(
                    out OfficeDiagramSnapshot snapshot));
                Assert.Equal(expectedKind, snapshot.Kind);
                presentation.Save();
            }

            stream.Position = 0;
            using PowerPointPresentation reopened = PowerPointPresentation.Load(stream);
            PowerPointSmartArt authored = Assert.Single(reopened.Slides[0].SmartArts);
            Assert.True(authored.TryGetOfficeDiagramSnapshot(
                out OfficeDiagramSnapshot reopenedSnapshot));
            Assert.Equal(expectedKind, reopenedSnapshot.Kind);
            Assert.Equal(new[] { "Discover", "Build", "Validate", "Ship" },
                authored.GetNodeTexts());
            Dgm.LayoutDefinition layout = Assert.Single(reopened.Slides[0].SlidePart
                .DiagramLayoutDefinitionParts).LayoutDefinition!;
            AssertSmartArtNativeLayoutContract(type, layout);
            OfficeImageExportResult png = reopened.Slides[0].ExportImage(
                OfficeImageExportFormat.Png);
            Assert.DoesNotContain(png.Diagnostics,
                diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error
                    || (diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Warning
                        && diagnostic.Code != OfficeImageExportDiagnosticCodes.FontSubstituted));
            Assert.True(OfficePngReader.TryDecode(png.Bytes,
                out OfficeRasterImage? raster));
            Assert.Equal(360, raster!.Width);
            Assert.Equal(220, raster.Height);
            Assert.Empty(reopened.ValidateDocument());
        }

        [Theory]
        [InlineData(PowerPointSmartArtType.BasicList)]
        [InlineData(PowerPointSmartArtType.BasicMatrix)]
        [InlineData(PowerPointSmartArtType.BasicPyramid)]
        [InlineData(PowerPointSmartArtType.BasicRelationship)]
        public void BroaderSemanticSmartArtRejectsXmlInvalidNodeTextBeforeMutation(
            PowerPointSmartArtType type) {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();

            Assert.Throws<ArgumentException>(() => slide.AddSmartArt(type,
                new[] { "Valid", "Bad\u0001node" }));

            Assert.Empty(slide.SmartArts);
            Assert.Empty(slide.SlidePart.DiagramDataParts);
            Assert.Empty(slide.SlidePart.DiagramLayoutDefinitionParts);
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void BasicListSmartArtRemainsSemanticAfterFrameResize() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSmartArt smartArt = presentation.AddSlide().AddSmartArt(
                PowerPointSmartArtType.BasicList,
                new[] { "Discover", "Build", "Ship" },
                PowerPointUnits.FromPoints(20),
                PowerPointUnits.FromPoints(20),
                PowerPointUnits.FromPoints(320),
                PowerPointUnits.FromPoints(180));

            smartArt.Width = PowerPointUnits.FromPoints(240);
            smartArt.Height = PowerPointUnits.FromPoints(240);

            Assert.True(smartArt.TryGetOfficeDiagramSnapshot(
                out OfficeDiagramSnapshot snapshot));
            Assert.Equal(OfficeDiagramKind.List, snapshot.Kind);
            Assert.Equal(240D, snapshot.WidthPoints, 3);
            Assert.Equal(240D, snapshot.HeightPoints, 3);
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void SmartArtNodeMutationRejectsXmlInvalidTextWithoutChangingData() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSmartArt smartArt = presentation.AddSlide().AddSmartArt(
                PowerPointSmartArtType.BasicProcess, new[] { "Original" });

            Assert.Throws<ArgumentException>(() =>
                smartArt.SetNodeText(0, "Bad\u0001node"));
            Assert.Throws<ArgumentException>(() =>
                smartArt.SetNodeText(0, "   "));

            Assert.Equal("Original", smartArt.GetNodeText(0));
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void EmbeddedSmartArtRenderingKeepsSlideBackgroundVisible() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(360, 220);
            PowerPointSlide slide = presentation.AddSlide();
            slide.BackgroundColor = "123456";
            slide.AddSmartArt(PowerPointSmartArtType.BasicProcess,
                new[] { "Start", "Finish" },
                PowerPointUnits.FromPoints(20), PowerPointUnits.FromPoints(20),
                PowerPointUnits.FromPoints(320), PowerPointUnits.FromPoints(180));

            OfficeImageExportResult png = slide.ExportImage(
                OfficeImageExportFormat.Png);

            Assert.True(OfficePngReader.TryDecode(png.Bytes,
                out OfficeRasterImage? raster));
            Assert.Equal(OfficeColor.FromRgb(0x12, 0x34, 0x56),
                raster!.GetPixel(22, 22));
        }

        private static void AssertSmartArtNativeLayoutContract(
            PowerPointSmartArtType type, Dgm.LayoutDefinition layout) {
            Dgm.AlgorithmValues expectedAlgorithm = type switch {
                PowerPointSmartArtType.BasicHierarchy => Dgm.AlgorithmValues.Composite,
                PowerPointSmartArtType.BasicList => Dgm.AlgorithmValues.Composite,
                PowerPointSmartArtType.BasicMatrix => Dgm.AlgorithmValues.Composite,
                PowerPointSmartArtType.BasicPyramid => Dgm.AlgorithmValues.Composite,
                PowerPointSmartArtType.BasicRelationship => Dgm.AlgorithmValues.Cycle,
                _ => throw new ArgumentOutOfRangeException(nameof(type))
            };
            Dgm.Algorithm algorithm = Assert.Single(layout.Descendants<Dgm.Algorithm>(),
                candidate => candidate.Type?.Value == expectedAlgorithm);
            Assert.NotEmpty(layout.Descendants<Dgm.Constraints>());
            Assert.NotEmpty(layout.Descendants<Dgm.RuleList>());
            Assert.All(layout.Descendants<Dgm.Constraint>().Where(constraint =>
                    constraint.Type?.Value == Dgm.ConstraintValues.CenterWidth),
                constraint => Assert.Equal(Dgm.ConstraintValues.Width,
                    constraint.ReferenceType?.Value));
            Assert.All(layout.Descendants<Dgm.Constraint>().Where(constraint =>
                    constraint.Type?.Value == Dgm.ConstraintValues.CenterHeight),
                constraint => Assert.Equal(Dgm.ConstraintValues.Height,
                    constraint.ReferenceType?.Value));

            if (type == PowerPointSmartArtType.BasicHierarchy) {
                Assert.Equal(4, layout.Descendants<Dgm.LayoutNode>().Count(node =>
                    node.Name?.Value == "hierarchyRootNode"
                    || node.Name?.Value?.StartsWith("hierarchyChild",
                        StringComparison.Ordinal) == true));
                Assert.Equal(4, layout.Descendants<Dgm.Shape>().Count(shape =>
                    shape.Type?.Value == "roundRect"));
                Assert.Equal(3, layout.Descendants<Dgm.Shape>().Count(shape =>
                    shape.Type?.Value == "conn"));
            } else if (type == PowerPointSmartArtType.BasicList) {
                Dgm.Parameter aspectRatio = Assert.Single(
                    algorithm.Elements<Dgm.Parameter>(), parameter =>
                        parameter.Type?.Value ==
                        Dgm.ParameterIdValues.AspectRatio);
                Assert.Equal(16D / 9D, double.Parse(aspectRatio.Val!.Value!,
                    CultureInfo.InvariantCulture), 8);
                Dgm.LayoutNode[] nodes = layout.Descendants<Dgm.LayoutNode>()
                    .Where(node => node.Name?.Value?.StartsWith("listNode",
                        StringComparison.Ordinal) == true).ToArray();
                Assert.Equal(4, nodes.Length);
                Assert.All(nodes, node => Assert.Equal("rect",
                    node.GetFirstChild<Dgm.Shape>()?.Type?.Value));
                Assert.Equal(4, layout.Descendants<Dgm.Constraint>().Count(
                    constraint => constraint.Type?.Value ==
                        Dgm.ConstraintValues.CenterWidth
                        && constraint.ForName?.Value?.StartsWith("listNode",
                            StringComparison.Ordinal) == true));
                Dgm.Constraint firstHorizontalCenter = Assert.Single(
                    layout.Descendants<Dgm.Constraint>(), constraint =>
                        constraint.Type?.Value == Dgm.ConstraintValues.CenterWidth
                        && constraint.ForName?.Value == "listNode1");
                Dgm.Constraint firstVerticalCenter = Assert.Single(
                    layout.Descendants<Dgm.Constraint>(), constraint =>
                        constraint.Type?.Value == Dgm.ConstraintValues.CenterHeight
                        && constraint.ForName?.Value == "listNode1");
                Assert.Equal(0.5D,
                    firstHorizontalCenter.Fact?.Value ?? double.NaN, 8);
                Assert.Equal(0.125D,
                    firstVerticalCenter.Fact?.Value ?? double.NaN, 8);
            } else if (type == PowerPointSmartArtType.BasicMatrix) {
                Assert.Equal(4, layout.Descendants<Dgm.LayoutNode>().Count(node =>
                    node.Name?.Value?.StartsWith("matrixNode",
                        StringComparison.Ordinal) == true));
                Assert.Equal(4, layout.Descendants<Dgm.Shape>().Count(shape =>
                    shape.Type?.Value == "roundRect"));
            } else if (type == PowerPointSmartArtType.BasicPyramid) {
                Assert.Equal(4, layout.Descendants<Dgm.LayoutNode>().Count(node =>
                    node.Name?.Value?.StartsWith("level",
                        StringComparison.Ordinal) == true));
                Assert.Equal(4, layout.Descendants<Dgm.Shape>().Count(shape =>
                    shape.Type?.Value == "trapezoid"));
            } else if (type == PowerPointSmartArtType.BasicRelationship) {
                Assert.Contains(algorithm.Elements<Dgm.Parameter>(), parameter =>
                    parameter.Type?.Value == Dgm.ParameterIdValues.CenterShapeMapping
                    && parameter.Val?.Value == "fNode");
                Assert.DoesNotContain(layout.Descendants<Dgm.Shape>(), shape =>
                    shape.Type?.Value == "conn");
            }
        }

        [Theory]
        [InlineData(PowerPointSmartArtType.BasicHierarchy)]
        [InlineData(PowerPointSmartArtType.BasicMatrix)]
        [InlineData(PowerPointSmartArtType.BasicPyramid)]
        public void PositionedSmartArtLayoutsProjectEveryAuthoredNode(
            PowerPointSmartArtType type) {
            string[] nodeTexts = Enumerable.Range(1, 7)
                .Select(index => "Node " + index)
                .ToArray();
            using var stream = new MemoryStream();
            using (PowerPointPresentation presentation = PowerPointPresentation.Create(stream)) {
                presentation.AddSlide().AddSmartArt(type, nodeTexts);
                presentation.Save();
            }

            stream.Position = 0;
            using PowerPointPresentation reopened = PowerPointPresentation.Load(stream);
            PowerPointSmartArt smartArt = Assert.Single(reopened.Slides[0].SmartArts);
            Assert.Equal(nodeTexts, smartArt.GetNodeTexts());
            Dgm.LayoutDefinition layout = Assert.Single(reopened.Slides[0].SlidePart
                .DiagramLayoutDefinitionParts).LayoutDefinition!;
            Assert.Equal(nodeTexts.Length, layout.Descendants<Dgm.Shape>().Count(shape =>
                !string.IsNullOrWhiteSpace(shape.Type?.Value)
                && shape.Type?.Value != "conn"));
            if (type == PowerPointSmartArtType.BasicHierarchy) {
                Assert.Equal(nodeTexts.Length - 1,
                    layout.Descendants<Dgm.Shape>().Count(shape =>
                        shape.Type?.Value == "conn"));
            }
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void NotesPagesAndHandoutsExportExistingNotesWithoutCreatingNewNotesParts() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream, new PowerPointCreateOptions());
            for (int index = 0; index < 3; index++) {
                PowerPointSlide slide = presentation.AddSlide();
                slide.AddTitle("Slide " + (index + 1));
                if (index < 2) slide.Notes.Text = "Speaker note " + (index + 1);
            }
            Assert.Null(presentation.Slides[2].SlidePart.NotesSlidePart);

            byte[] notesPdf = presentation.ToPdf(new PowerPointPdfSaveOptions {
                PageLayout = PowerPointPdfPageLayout.NotesPages,
                IncludeSpeakerNotes = true
            });
            byte[] handoutPdf = presentation.ToPdf(new PowerPointPdfSaveOptions {
                PageLayout = PowerPointPdfPageLayout.Handouts,
                HandoutSlidesPerPage = 3,
                IncludeSpeakerNotes = true
            });

            using var notes = PdfPigDocument.Open(new MemoryStream(notesPdf));
            using var handout = PdfPigDocument.Open(new MemoryStream(handoutPdf));
            Assert.Equal(3, notes.NumberOfPages);
            Assert.Equal(1, handout.NumberOfPages);
            Assert.Contains("Speaker note 1", notes.GetPage(1).Text, StringComparison.Ordinal);
            Assert.Contains("Speaker note 2", handout.GetPage(1).Text, StringComparison.Ordinal);
            Assert.Null(presentation.Slides[2].SlidePart.NotesSlidePart);
            Assert.Throws<ArgumentOutOfRangeException>(() => new PowerPointPdfSaveOptions {
                HandoutSlidesPerPage = 5
            });
        }

        [Theory]
        [InlineData(PowerPointPdfPageLayout.NotesPages)]
        [InlineData(PowerPointPdfPageLayout.Handouts)]
        public void NotesAndHandoutThumbnailsHonorPdfContentFilters(PowerPointPdfPageLayout layout) {
            using var controlStream = new MemoryStream();
            using var pictureStream = new MemoryStream();
            using PowerPointPresentation control = PowerPointPresentation.Create(controlStream, new PowerPointCreateOptions());
            using PowerPointPresentation withPicture = PowerPointPresentation.Create(pictureStream, new PowerPointCreateOptions());
            control.AddSlide().AddTitle("Filtered thumbnail");
            withPicture.AddSlide().AddTitle("Filtered thumbnail");
            withPicture.Slides[0].AddPicture(new MemoryStream(PdfPngTestImages.CreateRgbPng(255, 0, 0)),
                OfficeIMO.PowerPoint.PowerPointImagePartType.Png, PowerPointUnits.FromPoints(72), PowerPointUnits.FromPoints(72),
                PowerPointUnits.FromPoints(180), PowerPointUnits.FromPoints(120));

            var controlOptions = new PowerPointPdfSaveOptions { PageLayout = layout };
            var pictureOptions = new PowerPointPdfSaveOptions { PageLayout = layout };
            controlOptions.UseProfile(PdfCore.PdfExportProfile.TextOnly);
            pictureOptions.UseProfile(PdfCore.PdfExportProfile.TextOnly);

            byte[] controlThumbnail = PdfCore.PdfPageImageRenderer.RenderPageAsPng(
                control.ToPdf(controlOptions));
            byte[] pictureThumbnail = PdfCore.PdfPageImageRenderer.RenderPageAsPng(
                withPicture.ToPdf(pictureOptions));
            VisualRasterComparison comparison = VisualBaselineTestSupport.CompareRasterImages(
                controlThumbnail, pictureThumbnail, channelTolerance: 0, allowedDifferentPixels: 0);

            Assert.True(comparison.Passed,
                $"Filtered {layout} thumbnail changed at {comparison.DifferentPixels} pixels.");
        }

        [Fact]
        public void SignedPresentationSaveIsBlockedUntilMutationPolicyIsExplicit() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(path)) {
                    presentation.AddSlide().AddTitle("Signed workflow");
                    presentation.Save();
                }
                AddSyntheticSignature(path);

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(path)) {
                    PowerPointSignatureReport inspection = presentation.InspectSignatures();
                    Assert.True(inspection.HasSignatureMetadata);
                    Assert.Equal(1, inspection.XmlSignaturePartCount);
                    PowerPointSignedPresentationMutationException blocked =
                        Assert.Throws<PowerPointSignedPresentationMutationException>(() => presentation.Save());
                    Assert.Equal(PowerPointSignatureMutationAction.Blocked, blocked.Report.Action);

                    presentation.SignatureMutationPolicy =
                        OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures;
                    presentation.Save();
                    Assert.Equal(PowerPointSignatureMutationAction.Removed,
                        presentation.LastSignatureReport!.Action);
                }

                using PresentationDocument reopened = PresentationDocument.Open(path, false);
                Assert.Null(reopened.DigitalSignatureOriginPart);
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void SignedPresentationDisposeCannotBypassMutationPolicy() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(path)) {
                    presentation.AddSlide().AddTitle("Signed workflow");
                    presentation.Save();
                }
                AddSyntheticSignature(path);

                PowerPointPresentation edited = PowerPointPresentation.Load(path,
                    new PowerPointLoadOptions { PersistenceMode = OfficeIMO.DocumentPersistenceMode.SaveOnDispose });
                edited.Slides[0].AddTextBox("Must not persist");
                PowerPointSignedPresentationMutationException blocked =
                    Assert.Throws<PowerPointSignedPresentationMutationException>(() => edited.Dispose());

                Assert.Equal(PowerPointSignatureMutationAction.Blocked, blocked.Report.Action);
                using (PresentationDocument signed = PresentationDocument.Open(path, false)) {
                    Assert.NotNull(signed.DigitalSignatureOriginPart);
                }
                using PowerPointPresentation reopened = PowerPointPresentation.Load(path, new PowerPointLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
                Assert.DoesNotContain(reopened.Slides[0].TextBoxes,
                    textBox => textBox.Text == "Must not persist");
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void UntouchedSignedPresentationCanBeInspectedThroughEditableOpen() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(path)) {
                    presentation.AddSlide().AddTitle("Signed inspection");
                    presentation.Save();
                }
                AddSyntheticSignature(path);

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(path, new PowerPointLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                    Assert.True(presentation.InspectSignatures().HasSignatureMetadata);
                    Assert.Equal("Signed inspection", presentation.Slides[0].TextBoxes.First().Text);
                }

                using PresentationDocument signed = PresentationDocument.Open(path, false);
                Assert.NotNull(signed.DigitalSignatureOriginPart);
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void SmartArtDeckCanUseOptInPowerPointDesktopReferenceLane() {
            if (!string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_POWERPOINT_DESKTOP_REFERENCE"),
                    "1", StringComparison.Ordinal)) return;
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string output = Path.Combine(Path.GetTempPath(), "OfficeIMO.SmartArtReference", Guid.NewGuid().ToString("N"));
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(path)) {
                    foreach (PowerPointSmartArtType type in
                             (PowerPointSmartArtType[])Enum.GetValues(typeof(PowerPointSmartArtType))) {
                        presentation.AddSlide().AddSmartArt(type,
                            new[] { "Inspect", "Build", "Validate", "Ship" });
                    }
                    presentation.Save();
                }
                PowerPointReferenceRenderResult result = PowerPointDesktopReferenceRenderer.TryRender(path, output,
                    enabled: true);
                Assert.True(result.IsSuccessful, result.Message);
                Assert.Equal(Enum.GetValues(typeof(PowerPointSmartArtType)).Length,
                    result.ImagePaths.Count);
            } finally {
                if (File.Exists(path)) File.Delete(path);
                if (Directory.Exists(output)) Directory.Delete(output, recursive: true);
            }
        }

        [Fact]
        public void DesktopReferenceLaneRemovesOnlyStalePowerPointSlideImages() {
            string output = Path.Combine(Path.GetTempPath(), "OfficeIMO.ReferenceCleanup",
                Guid.NewGuid().ToString("N"));
            try {
                Directory.CreateDirectory(output);
                File.WriteAllBytes(Path.Combine(output, "Slide1.png"), new byte[] { 1 });
                File.WriteAllBytes(Path.Combine(output, "slide12.PNG"), new byte[] { 2 });
                File.WriteAllBytes(Path.Combine(output, "comparison.png"), new byte[] { 3 });

                PowerPointDesktopReferenceRenderer.ClearExistingSlideImages(output);

                Assert.False(File.Exists(Path.Combine(output, "Slide1.png")));
                Assert.False(File.Exists(Path.Combine(output, "slide12.PNG")));
                Assert.True(File.Exists(Path.Combine(output, "comparison.png")));
            } finally {
                if (Directory.Exists(output)) Directory.Delete(output, recursive: true);
            }
        }

        [Fact]
        public void DesktopReferenceLaneReturnsOnlySlideImagesInNumericOrder() {
            string output = Path.Combine(Path.GetTempPath(), "OfficeIMO.ReferenceOrder",
                Guid.NewGuid().ToString("N"));
            try {
                Directory.CreateDirectory(output);
                File.WriteAllBytes(Path.Combine(output, "Slide10.png"), new byte[] { 10 });
                File.WriteAllBytes(Path.Combine(output, "Slide2.PNG"), new byte[] { 2 });
                File.WriteAllBytes(Path.Combine(output, "Slide1.png"), new byte[] { 1 });
                File.WriteAllBytes(Path.Combine(output, "comparison.png"), new byte[] { 3 });

                string[] images = PowerPointDesktopReferenceRenderer.GetSlideImagesInOrder(output);

                Assert.Equal(new[] { "Slide1.png", "Slide2.PNG", "Slide10.png" },
                    images.Select(Path.GetFileName));
            } finally {
                if (Directory.Exists(output)) Directory.Delete(output, recursive: true);
            }
        }

        [Fact]
        public void DesktopReferenceLaneRequiresACompleteValidPngSet() {
            string output = Path.Combine(Path.GetTempPath(),
                "OfficeIMO.PowerPointReferenceValidation", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(output);
            try {
                string first = Path.Combine(output, "Slide1.png");
                string second = Path.Combine(output, "Slide2.png");
                File.WriteAllBytes(first, VisualBaselineTestSupport.CreateRgbPng(
                    2, 1, new byte[] { 10, 20, 30, 40, 50, 60 }));
                File.WriteAllBytes(second, VisualBaselineTestSupport.CreateRgbPng(
                    1, 1, new byte[] { 70, 80, 90 }));

                Assert.True(PowerPointDesktopReferenceRenderer.ValidateSlideImages(
                    new[] { first, second }, 2, out string completeMessage),
                    completeMessage);
                Assert.False(PowerPointDesktopReferenceRenderer.ValidateSlideImages(
                    new[] { first }, 2, out string incompleteMessage));
                Assert.Contains("expected 2", incompleteMessage, StringComparison.Ordinal);
                Assert.False(PowerPointDesktopReferenceRenderer.ValidateSlideImages(
                    Array.Empty<string>(), 0, out string emptyMessage));
                Assert.Contains("no slide images", emptyMessage,
                    StringComparison.OrdinalIgnoreCase);

                string third = Path.Combine(output, "Slide3.png");
                File.Copy(second, third);
                Assert.False(PowerPointDesktopReferenceRenderer.ValidateSlideImages(
                    new[] { first, third }, 2, out string nonContiguousMessage));
                Assert.Contains("contiguous", nonContiguousMessage,
                    StringComparison.OrdinalIgnoreCase);

                File.WriteAllText(second, "not a PNG");
                Assert.False(PowerPointDesktopReferenceRenderer.ValidateSlideImages(
                    new[] { first, second }, 2, out string invalidMessage));
                Assert.Contains("invalid PNG", invalidMessage, StringComparison.Ordinal);

                string blank = Path.Combine(output, "Slide1.png");
                File.WriteAllBytes(blank, VisualBaselineTestSupport.CreateRgbPng(
                    2, 2, Enumerable.Repeat((byte)255, 12).ToArray()));
                Assert.False(PowerPointDesktopReferenceRenderer.ValidateSlideImages(
                    new[] { blank }, 1, new[] { true }, out string blankMessage));
                Assert.Contains("blank PNG", blankMessage,
                    StringComparison.Ordinal);

                string noisy = Path.Combine(output, "Slide1.png");
                File.WriteAllBytes(noisy, VisualBaselineTestSupport.CreateRgbPng(
                    2, 2, new byte[] {
                        0, 0, 0, 255, 255, 255,
                        255, 255, 255, 255, 255, 255
                    }));
                Assert.False(PowerPointDesktopReferenceRenderer.ValidateSlideImages(
                    new[] { noisy }, 1, new[] { true },
                    out string noisyMessage));
                Assert.Contains("blank PNG", noisyMessage,
                    StringComparison.Ordinal);

                File.WriteAllBytes(blank, VisualBaselineTestSupport.CreateRgbPng(
                    2, 2, Enumerable.Repeat((byte)255, 12).ToArray()));
                Assert.True(PowerPointDesktopReferenceRenderer.ValidateSlideImages(
                    new[] { blank }, 1, new[] { false },
                    out string intentionallyBlankMessage), intentionallyBlankMessage);
            } finally {
                if (Directory.Exists(output)) Directory.Delete(output, recursive: true);
            }
        }

        [Fact]
        public void DesktopReferenceLaneDerivesExpectedContentFromRenderedPaint() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            presentation.SlideSize.SetSizePoints(100, 100);

            PowerPointSlide offCanvas = presentation.AddSlide();
            PowerPointAutoShape offCanvasShape = offCanvas.AddRectanglePoints(
                120, 120, 20, 20);
            offCanvasShape.FillColor = "FF0000";
            Assert.False(PowerPointDesktopReferenceRenderer
                .HasExpectedVisibleContent(offCanvas));

            PowerPointSlide unpainted = presentation.AddSlide();
            PowerPointAutoShape unpaintedShape = unpainted.AddRectanglePoints(
                20, 20, 40, 40);
            ShapeProperties unpaintedProperties = ((Shape)unpaintedShape.Element)
                .ShapeProperties!;
            unpaintedProperties.RemoveAllChildren<A.SolidFill>();
            unpaintedProperties.InsertAfter(new A.NoFill(),
                unpaintedProperties.GetFirstChild<A.PresetGeometry>()!);
            unpaintedProperties.Append(new A.Outline(new A.NoFill()));
            Assert.False(PowerPointDesktopReferenceRenderer
                .HasExpectedVisibleContent(unpainted));

            PowerPointSlide transparent = presentation.AddSlide();
            PowerPointAutoShape transparentShape = transparent
                .AddRectanglePoints(20, 20, 40, 40);
            transparentShape.FillColor = "FF0000";
            transparentShape.FillTransparency = 100;
            Assert.False(PowerPointDesktopReferenceRenderer
                .HasExpectedVisibleContent(transparent));

            PowerPointSlide visible = presentation.AddSlide();
            PowerPointAutoShape visibleShape = visible.AddRectanglePoints(
                20, 20, 40, 40);
            visibleShape.FillColor = "FF0000";
            Assert.True(PowerPointDesktopReferenceRenderer
                .HasExpectedVisibleContent(visible));
        }

        [Fact]
        public void DesktopReferenceLaneTreatsSkippedVisibleChartAsExpectedContent() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            presentation.SlideSize.SetSizePoints(640, 360);
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointChart chart = slide.AddChartPoints(
                OfficeChartKind.ColumnClustered,
                new OfficeChartData(new[] { "Q1", "Q2" }, new[] {
                    new OfficeChartSeries("Actual", new[] { 10D, 20D })
                }), 40, 40, 560, 280);
            chart.SetLegendTextStyle(fontName: "Arial")
                .SetCategoryAxisLabelTextStyle(fontName: "Georgia")
                .SetValueAxisLabelTextStyle(fontName: "Arial");

            OfficeImageExportResult rendered = slide.ExportImage(
                OfficeImageExportFormat.Png);
            Assert.Contains(rendered.Diagnostics, diagnostic =>
                diagnostic.Code
                    == PowerPointImageExportDiagnosticCodes.UnsupportedShape
                && diagnostic.Message.Contains("chart",
                    StringComparison.OrdinalIgnoreCase));
            Assert.True(PowerPointDesktopReferenceRenderer
                .HasExpectedVisibleContent(slide));

            chart.LeftPoints = 700;
            Assert.False(PowerPointDesktopReferenceRenderer
                .HasExpectedVisibleContent(slide));
        }

        [Fact]
        public void DesktopReferenceLaneFailsClosedWhenMacroSecurityCannotBeSet() {
            Assert.False(PowerPointDesktopReferenceRenderer
                .TryConfigureApplicationSecurity(new object(),
                    out string message));
            Assert.Contains("force-disable macros", message,
                StringComparison.OrdinalIgnoreCase);
        }

        private static void FeedXml(OpenXmlPart part, string xml) {
            using var data = new MemoryStream(Encoding.UTF8.GetBytes(xml));
            part.FeedData(data);
        }

        private static void AddSyntheticSignature(string path) {
            using PresentationDocument document = PresentationDocument.Open(path, true);
            DigitalSignatureOriginPart origin = document.AddDigitalSignatureOriginPart();
            XmlSignaturePart signature = origin.AddNewPart<XmlSignaturePart>();
            FeedXml(signature,
                "<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\"><SignedInfo/><SignatureValue>AA==</SignatureValue></Signature>");
        }
    }
}
