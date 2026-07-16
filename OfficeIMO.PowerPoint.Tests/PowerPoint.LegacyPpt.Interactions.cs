using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Capabilities;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests {
    public partial class PowerPointLegacyPptInteractionTests {
        [Fact]
        public void NativeWriter_AuthorsAndProjectsShapeTextAndJumpInteractions() {
            var shapeUri = new Uri("https://example.com/shape");
            var textUri = new Uri("https://example.com/text?q=1");
            byte[] bytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                PowerPointSlide slide = source.AddSlide(P.SlideLayoutValues.Blank);
                PowerPointTextBox textBox = slide.AddTextBox("Visit OfficeIMO",
                    100000, 100000, 3000000, 500000);
                P.Shape textShape = (P.Shape)textBox.Element;
                HyperlinkRelationship textRelationship = slide.SlidePart
                    .AddHyperlinkRelationship(textUri, true);
                textShape.TextBody = new P.TextBody(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(
                        new A.Run(new A.Text("Visit ")),
                        new A.Run(
                            new A.RunProperties(
                                new A.HyperlinkOnClick { Id = textRelationship.Id }),
                            new A.Text("OfficeIMO"))));

                PowerPointAutoShape linkedShape = slide.AddRectangle(
                    100000, 800000, 1200000, 500000);
                AddShapeHyperlink(slide, linkedShape, shapeUri,
                    mouseOver: true);
                PowerPointAutoShape actionShape = slide.AddEllipse(
                    1600000, 800000, 1200000, 500000);
                GetDrawingProperties(actionShape).Append(new A.HyperlinkOnClick {
                    Id = string.Empty,
                    Action = "ppaction://hlinkshowjump?jump=nextslide"
                });

                LegacyPptWritePreflightReport report = source.AnalyzeLegacyPptWrite();
                Assert.True(report.CanWrite, string.Join(Environment.NewLine, report.Findings));
                bytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation legacy = LegacyPptPresentation.Load(bytes);
            Assert.Equal(2, legacy.Hyperlinks.Count);
            Assert.Contains(legacy.Hyperlinks, hyperlink => hyperlink.Uri == textUri);
            Assert.Contains(legacy.Hyperlinks, hyperlink => hyperlink.Uri == shapeUri);
            LegacyPptSlide legacySlide = Assert.Single(legacy.Slides);
            LegacyPptShape legacyText = Assert.Single(legacySlide.Shapes,
                shape => shape.Text == "Visit OfficeIMO");
            LegacyPptTextInteraction textInteraction = Assert.Single(
                legacyText.TextBody.Interactions);
            Assert.Equal(6, textInteraction.Start);
            Assert.Equal(9, textInteraction.Length);
            Assert.Equal(LegacyPptInteractionTrigger.MouseClick,
                textInteraction.Interaction.Trigger);
            Assert.Equal(textUri, textInteraction.Interaction.Hyperlink!.Uri);
            LegacyPptInteraction hover = Assert.Single(legacySlide.Shapes
                .SelectMany(shape => shape.Interactions), interaction =>
                    interaction.Trigger == LegacyPptInteractionTrigger.MouseOver);
            Assert.Equal(shapeUri, hover.Hyperlink!.Uri);
            LegacyPptInteraction jump = Assert.Single(legacySlide.Shapes
                .SelectMany(shape => shape.Interactions), interaction =>
                    interaction.Action == LegacyPptInteractionAction.Jump);
            Assert.Equal(LegacyPptInteractionJump.NextSlide, jump.Jump);
            LegacyPptImportReport inventory = legacy.CreateImportReport();
            Assert.Equal(2, inventory.HyperlinkTargetCount);
            Assert.Equal(2, inventory.ShapeInteractionCount);
            Assert.Equal(1, inventory.TextInteractionCount);

            using var input = new MemoryStream(bytes, writable: false);
            using PowerPointPresentation projected = PowerPointPresentation.Load(input);
            PowerPointSlide projectedSlide = Assert.Single(projected.Slides);
            OpenXmlElement projectedHover = Assert.Single(projectedSlide.SlidePart
                .Slide!.Descendants(), element => element.LocalName == "hlinkHover");
            string hoverRelationshipId = projectedHover.GetAttribute("id",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships").Value;
            Assert.Equal(shapeUri, projectedSlide.SlidePart.HyperlinkRelationships
                .Single(relationship => relationship.Id == hoverRelationshipId).Uri);
            P.Shape projectedText = projectedSlide.SlidePart.Slide!.CommonSlideData!
                .ShapeTree!.Elements<P.Shape>().Single(shape =>
                    shape.TextBody?.InnerText == "Visit OfficeIMO");
            A.Run projectedRun = projectedText.TextBody!.Descendants<A.Run>()
                .Single(run => run.InnerText == "OfficeIMO");
            string textRelationshipId = projectedRun.RunProperties!
                .GetFirstChild<A.HyperlinkOnClick>()!.Id!.Value!;
            Assert.Equal(textUri, projectedSlide.SlidePart.HyperlinkRelationships
                .Single(relationship => relationship.Id == textRelationshipId).Uri);
            Assert.Contains(projectedSlide.SlidePart.Slide.Descendants<A.HyperlinkOnClick>(),
                action => action.Action?.Value
                    == "ppaction://hlinkshowjump?jump=nextslide");
            Assert.Empty(projected.ValidateDocument());
        }

        [Fact]
        public void NativeWriter_BlocksTargetFrameWithoutBinaryBaseRepresentation() {
            using PowerPointPresentation source = PowerPointPresentation.Create();
            PowerPointSlide slide = source.AddSlide(P.SlideLayoutValues.Blank);
            PowerPointTextBox textBox = slide.AddTextBox("Details");
            HyperlinkRelationship relationship = slide.SlidePart.AddHyperlinkRelationship(
                new Uri("https://example.com/details"), true);
            P.Shape shape = (P.Shape)textBox.Element;
            A.Run run = shape.TextBody!.Descendants<A.Run>().First();
            run.RunProperties ??= new A.RunProperties();
            run.RunProperties.Append(new A.HyperlinkOnClick {
                Id = relationship.Id,
                TargetFrame = "_blank"
            });

            LegacyPptWritePreflightReport report = source.AnalyzeLegacyPptWrite();

            LegacyPptWriteFinding finding = Assert.Single(report.Findings,
                item => item.Code == "PPT-WRITE-INTERACTION");
            Assert.Equal(LegacyPptFeature.Hyperlinks, finding.Feature);
            Assert.Contains("target frames", finding.Description,
                StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void NativeWriter_AuthorsAndProjectsPpt9ScreenTip() {
            var target = new Uri("https://example.com/screen-tip");
            byte[] bytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                PowerPointSlide slide = source.AddSlide(P.SlideLayoutValues.Blank);
                PowerPointTextBox textBox = slide.AddTextBox("Details");
                textBox.Paragraphs.Single().Runs.Single().SetHyperlink(target,
                    "Open the detailed report");

                LegacyPptWritePreflightReport report = source.AnalyzeLegacyPptWrite();
                Assert.True(report.CanWrite, string.Join(Environment.NewLine, report.Findings));
                bytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation legacy = LegacyPptPresentation.Load(bytes);
            LegacyPptHyperlink hyperlink = Assert.Single(legacy.Hyperlinks);
            Assert.Equal(target, hyperlink.Uri);
            Assert.Equal("Open the detailed report", hyperlink.ScreenTip);
            Assert.Equal(0U, hyperlink.ExtensionFlags);
            LegacyPptImportReport inventory = legacy.CreateImportReport();
            Assert.Equal(1, inventory.HyperlinkScreenTipCount);
            Assert.Equal(0, inventory.HyperlinkExtensionFlagCount);

            using var input = new MemoryStream(bytes, writable: false);
            using PowerPointPresentation projected = PowerPointPresentation.Load(input);
            A.HyperlinkOnClick projectedLink = projected.Slides[0].SlidePart.Slide!
                .Descendants<A.HyperlinkOnClick>().Single();
            Assert.Equal("Open the detailed report", projectedLink.Tooltip?.Value);
            Assert.Empty(projected.ValidateDocument());
        }

        [Fact]
        public void NativeWriter_KeepsDistinctScreenTipsForTheSameTarget() {
            var target = new Uri("https://example.com/shared-target");
            byte[] bytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                PowerPointSlide slide = source.AddSlide(P.SlideLayoutValues.Blank);
                AddShapeHyperlink(slide, slide.AddRectangle(
                    100000, 100000, 1000000, 500000), target,
                    mouseOver: false, screenTip: "First tip");
                AddShapeHyperlink(slide, slide.AddRectangle(
                    1200000, 100000, 1000000, 500000), target,
                    mouseOver: false, screenTip: "Second tip");
                Assert.True(source.AnalyzeLegacyPptWrite().CanWrite);
                bytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation legacy = LegacyPptPresentation.Load(bytes);
            Assert.Equal(2, legacy.Hyperlinks.Count);
            Assert.Equal(new[] { "First tip", "Second tip" }, legacy.Hyperlinks
                .Select(link => link.ScreenTip).OrderBy(value => value).ToArray());
        }

        [Fact]
        public void NativeWriter_BlocksDuplicateShapeInteractionTriggers() {
            using PowerPointPresentation source = PowerPointPresentation.Create();
            PowerPointSlide slide = source.AddSlide(P.SlideLayoutValues.Blank);
            PowerPointAutoShape shape = slide.AddRectangle(
                100000, 100000, 1000000, 500000);
            P.NonVisualDrawingProperties properties = GetDrawingProperties(shape);
            properties.Append(new A.HyperlinkOnClick {
                Id = slide.SlidePart.AddHyperlinkRelationship(
                    new Uri("https://example.com/first"), true).Id
            });
            properties.Append(new A.HyperlinkOnClick {
                Id = slide.SlidePart.AddHyperlinkRelationship(
                    new Uri("https://example.com/second"), true).Id
            });

            LegacyPptWriteFinding finding = Assert.Single(
                source.AnalyzeLegacyPptWrite().Findings,
                item => item.Code == "PPT-WRITE-INTERACTION");
            Assert.Contains("multiple", finding.Description,
                StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void NativeWriter_BlocksDuplicateTextInteractionTriggers() {
            using PowerPointPresentation source = PowerPointPresentation.Create();
            PowerPointSlide slide = source.AddSlide(P.SlideLayoutValues.Blank);
            P.Shape shape = (P.Shape)slide.AddTextBox("Duplicate").Element;
            A.Run run = shape.TextBody!.Descendants<A.Run>().Single();
            run.RunProperties = new A.RunProperties(
                new A.HyperlinkOnClick {
                    Id = slide.SlidePart.AddHyperlinkRelationship(
                        new Uri("https://example.com/first"), true).Id
                },
                new A.HyperlinkOnClick {
                    Id = slide.SlidePart.AddHyperlinkRelationship(
                        new Uri("https://example.com/second"), true).Id
                });

            LegacyPptWriteFinding finding = Assert.Single(
                source.AnalyzeLegacyPptWrite().Findings,
                item => item.Code == "PPT-WRITE-INTERACTION");
            Assert.Contains("multiple", finding.Description,
                StringComparison.OrdinalIgnoreCase);
        }

        [Theory]
        [InlineData("nextslide", LegacyPptInteractionJump.NextSlide)]
        [InlineData("previousslide", LegacyPptInteractionJump.PreviousSlide)]
        [InlineData("firstslide", LegacyPptInteractionJump.FirstSlide)]
        [InlineData("lastslide", LegacyPptInteractionJump.LastSlide)]
        [InlineData("lastslideviewed", LegacyPptInteractionJump.LastViewedSlide)]
        [InlineData("endshow", LegacyPptInteractionJump.EndShow)]
        public void NativeWriter_AuthorsEveryBuiltInSlideShowJump(string actionName,
            LegacyPptInteractionJump expected) {
            byte[] bytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                PowerPointAutoShape shape = source.AddSlide(P.SlideLayoutValues.Blank).AddRectangle(
                    100000, 100000, 1000000, 500000);
                GetDrawingProperties(shape).Append(new A.HyperlinkOnClick {
                    Id = string.Empty,
                    Action = "ppaction://hlinkshowjump?jump=" + actionName
                });
                Assert.True(source.AnalyzeLegacyPptWrite().CanWrite);
                bytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptInteraction interaction = Assert.Single(
                LegacyPptPresentation.Load(bytes).Slides[0].Shapes[0].Interactions);
            Assert.Equal(LegacyPptInteractionAction.Jump, interaction.Action);
            Assert.Equal(expected, interaction.Jump);
        }

        [Fact]
        public void NativeWriter_AuthorsTextMouseOverHyperlink() {
            var target = new Uri("https://example.com/text-hover");
            byte[] bytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                PowerPointSlide slide = source.AddSlide(P.SlideLayoutValues.Blank);
                PowerPointTextBox textBox = slide.AddTextBox("Hover");
                HyperlinkRelationship relationship = slide.SlidePart
                    .AddHyperlinkRelationship(target, true);
                A.Run run = ((P.Shape)textBox.Element).TextBody!
                    .Descendants<A.Run>().Single();
                run.RunProperties ??= new A.RunProperties();
                run.RunProperties.Append(new A.HyperlinkOnMouseOver {
                    Id = relationship.Id
                });
                Assert.True(source.AnalyzeLegacyPptWrite().CanWrite);
                bytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptTextInteraction interaction = Assert.Single(
                LegacyPptPresentation.Load(bytes).Slides[0].Shapes[0]
                    .TextBody.Interactions);
            Assert.Equal(LegacyPptInteractionTrigger.MouseOver,
                interaction.Interaction.Trigger);
            Assert.Equal(target, interaction.Interaction.Hyperlink!.Uri);
        }

        [Fact]
        public void ImportedHyperlink_UnrelatedGeometryEditPreservesRecordGraph() {
            byte[] sourceBytes;
            var target = new Uri("https://example.com/preserved");
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                PowerPointSlide slide = source.AddSlide(P.SlideLayoutValues.Blank);
                PowerPointAutoShape rectangle = slide.AddRectangle(
                    100000, 100000, 1000000, 500000);
                AddShapeHyperlink(slide, rectangle, target, mouseOver: false);
                sourceBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original = LegacyPptPresentation.Load(sourceBytes);

            byte[] savedBytes;
            using (var input = new MemoryStream(sourceBytes, writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation.Load(input)) {
                PowerPointShape linked = Assert.Single(imported.Slides[0].Shapes,
                    shape => shape.Hyperlink == target);
                linked.Left += 250000;

                LegacyPptWritePreflightReport report = imported.AnalyzeLegacyPptWrite();
                Assert.True(report.CanWrite, string.Join(Environment.NewLine, report.Findings));
                savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation saved = LegacyPptPresentation.Load(savedBytes);
            LegacyPptInteraction interaction = Assert.Single(
                Assert.Single(saved.Slides).Shapes.SelectMany(shape => shape.Interactions));
            Assert.Equal(target, interaction.Hyperlink!.Uri);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));
        }

        [Fact]
        public void ImportedHyperlink_NoOpSaveIsByteExact() {
            byte[] sourceBytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                PowerPointSlide slide = source.AddSlide(P.SlideLayoutValues.Blank);
                PowerPointAutoShape shape = slide.AddRectangle(
                    100000, 100000, 1000000, 500000);
                AddShapeHyperlink(slide, shape,
                    new Uri("https://example.com/no-op"), mouseOver: true);
                sourceBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            byte[] savedBytes;
            using (var input = new MemoryStream(sourceBytes, writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation.Load(input)) {
                savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }

            Assert.Equal(sourceBytes, savedBytes);
        }

        [Fact]
        public void ImportedHyperlink_EditAndRemoveScreenTipAppendsPreservingRecords() {
            byte[] sourceBytes;
            var target = new Uri("https://example.com/edit-screen-tip");
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                PowerPointSlide slide = source.AddSlide(P.SlideLayoutValues.Blank);
                AddShapeHyperlink(slide, slide.AddRectangle(
                    100000, 100000, 1000000, 500000), target,
                    mouseOver: false, screenTip: "Original tip");
                sourceBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original = LegacyPptPresentation.Load(sourceBytes);

            byte[] editedBytes;
            using (var input = new MemoryStream(sourceBytes, writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation.Load(input)) {
                A.HyperlinkOnClick link = imported.Slides[0].SlidePart.Slide!
                    .Descendants<A.HyperlinkOnClick>().Single();
                link.Tooltip = "Updated tip";
                LegacyPptWritePreflightReport report = imported.AnalyzeLegacyPptWrite();
                Assert.True(report.CanWrite,
                    string.Join(Environment.NewLine, report.Findings));
                editedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation edited = LegacyPptPresentation.Load(editedBytes);
            Assert.Equal("Updated tip", Assert.Single(edited.Slides[0].Shapes
                .SelectMany(shape => shape.Interactions)).Hyperlink!.ScreenTip);
            Assert.True(edited.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));

            byte[] removedBytes;
            using (var input = new MemoryStream(editedBytes, writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation.Load(input)) {
                A.HyperlinkOnClick link = imported.Slides[0].SlidePart.Slide!
                    .Descendants<A.HyperlinkOnClick>().Single();
                link.Tooltip = null;
                Assert.True(imported.AnalyzeLegacyPptWrite().CanWrite);
                removedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation removed = LegacyPptPresentation.Load(removedBytes);
            Assert.Null(Assert.Single(removed.Slides[0].Shapes
                .SelectMany(shape => shape.Interactions)).Hyperlink!.ScreenTip);
            Assert.True(removed.Package.DocumentStream.AsSpan(0,
                    edited.Package.DocumentStream.Length)
                .SequenceEqual(edited.Package.DocumentStream));
        }

        [Fact]
        public void ImportedShapeHyperlink_AddEditAndRemoveAppendPreservingRecords() {
            byte[] sourceBytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                source.AddSlide(P.SlideLayoutValues.Blank)
                    .AddRectangle(100000, 100000, 1000000, 500000);
                sourceBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original = LegacyPptPresentation.Load(sourceBytes);
            var firstTarget = new Uri("https://example.com/first");

            byte[] addedBytes;
            using (var input = new MemoryStream(sourceBytes, writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation.Load(input)) {
                PowerPointSlide slide = imported.Slides[0];
                AddShapeHyperlink(slide, Assert.Single(slide.Shapes), firstTarget,
                    mouseOver: false);
                LegacyPptWritePreflightReport report = imported.AnalyzeLegacyPptWrite();
                Assert.True(report.CanWrite, string.Join(Environment.NewLine, report.Findings));
                addedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation added = LegacyPptPresentation.Load(addedBytes);
            Assert.Equal(firstTarget, Assert.Single(added.Slides[0].Shapes
                .SelectMany(shape => shape.Interactions)).Hyperlink!.Uri);
            Assert.True(added.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));

            var secondTarget = new Uri("https://example.com/second");
            byte[] editedBytes;
            using (var input = new MemoryStream(addedBytes, writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation.Load(input)) {
                PowerPointSlide slide = imported.Slides[0];
                PowerPointShape shape = Assert.Single(slide.Shapes);
                P.NonVisualDrawingProperties properties = GetDrawingProperties(shape);
                properties.RemoveAllChildren<A.HyperlinkOnClick>();
                AddShapeHyperlink(slide, shape, secondTarget, mouseOver: false);
                Assert.True(imported.AnalyzeLegacyPptWrite().CanWrite);
                editedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation edited = LegacyPptPresentation.Load(editedBytes);
            Assert.Equal(secondTarget, Assert.Single(edited.Slides[0].Shapes
                .SelectMany(shape => shape.Interactions)).Hyperlink!.Uri);
            Assert.True(edited.Package.DocumentStream.AsSpan(0,
                    added.Package.DocumentStream.Length)
                .SequenceEqual(added.Package.DocumentStream));

            byte[] removedBytes;
            using (var input = new MemoryStream(editedBytes, writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation.Load(input)) {
                GetDrawingProperties(Assert.Single(imported.Slides[0].Shapes))
                    .RemoveAllChildren<A.HyperlinkOnClick>();
                Assert.True(imported.AnalyzeLegacyPptWrite().CanWrite);
                removedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation removed = LegacyPptPresentation.Load(removedBytes);
            Assert.Empty(removed.Slides[0].Shapes.SelectMany(shape => shape.Interactions));
            Assert.True(removed.Package.DocumentStream.AsSpan(0,
                    edited.Package.DocumentStream.Length)
                .SequenceEqual(edited.Package.DocumentStream));
        }

        [Fact]
        public void ImportedTextHyperlink_AddMoveAndRemoveAppendPreservingRecords() {
            byte[] sourceBytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                source.AddSlide(P.SlideLayoutValues.Blank)
                    .AddTextBox("Visit OfficeIMO");
                sourceBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original = LegacyPptPresentation.Load(sourceBytes);

            byte[] addedBytes;
            using (var input = new MemoryStream(sourceBytes, writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation.Load(input)) {
                PowerPointSlide slide = imported.Slides[0];
                P.Shape shape = GetOnlyTextShape(slide);
                HyperlinkRelationship relationship = slide.SlidePart.AddHyperlinkRelationship(
                    new Uri("https://example.com/text-range"), true);
                SetLinkedText(shape, relationship.Id, linkFirstRun: false);
                LegacyPptWritePreflightReport report = imported.AnalyzeLegacyPptWrite();
                Assert.True(report.CanWrite, string.Join(Environment.NewLine, report.Findings));
                addedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation added = LegacyPptPresentation.Load(addedBytes);
            LegacyPptTextInteraction addedRange = Assert.Single(
                added.Slides[0].Shapes[0].TextBody.Interactions);
            Assert.Equal(6, addedRange.Start);
            Assert.Equal(9, addedRange.Length);
            Assert.True(added.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));

            byte[] movedBytes;
            using (var input = new MemoryStream(addedBytes, writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation.Load(input)) {
                P.Shape shape = GetOnlyTextShape(imported.Slides[0]);
                string relationshipId = shape.TextBody!.Descendants<A.HyperlinkOnClick>()
                    .Single().Id!.Value!;
                SetLinkedText(shape, relationshipId, linkFirstRun: true);
                Assert.True(imported.AnalyzeLegacyPptWrite().CanWrite);
                movedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation moved = LegacyPptPresentation.Load(movedBytes);
            LegacyPptTextInteraction movedRange = Assert.Single(
                moved.Slides[0].Shapes[0].TextBody.Interactions);
            Assert.Equal(0, movedRange.Start);
            Assert.Equal(6, movedRange.Length);
            Assert.True(moved.Package.DocumentStream.AsSpan(0,
                    added.Package.DocumentStream.Length)
                .SequenceEqual(added.Package.DocumentStream));

            byte[] removedBytes;
            using (var input = new MemoryStream(movedBytes, writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation.Load(input)) {
                foreach (A.RunProperties properties in GetOnlyTextShape(imported.Slides[0])
                             .TextBody!.Descendants<A.RunProperties>()) {
                    properties.RemoveAllChildren<A.HyperlinkOnClick>();
                }
                Assert.True(imported.AnalyzeLegacyPptWrite().CanWrite);
                removedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation removed = LegacyPptPresentation.Load(removedBytes);
            Assert.Empty(removed.Slides[0].Shapes[0].TextBody.Interactions);
            Assert.True(removed.Package.DocumentStream.AsSpan(0,
                    moved.Package.DocumentStream.Length)
                .SequenceEqual(moved.Package.DocumentStream));
        }

        [Fact]
        public void ImportedPresentation_AppendedSlideCanCarryNativeHyperlink() {
            byte[] sourceBytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                source.AddSlide(P.SlideLayoutValues.Blank).AddTextBox("Existing");
                sourceBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original = LegacyPptPresentation.Load(sourceBytes);
            var target = new Uri("https://example.com/appended-slide");

            byte[] savedBytes;
            using (var input = new MemoryStream(sourceBytes, writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation.Load(input)) {
                PowerPointSlide added = imported.AddSlide(P.SlideLayoutValues.Blank);
                PowerPointAutoShape rectangle = added.AddRectangle(
                    100000, 100000, 1000000, 500000);
                AddShapeHyperlink(added, rectangle, target, mouseOver: false);
                LegacyPptWritePreflightReport report = imported.AnalyzeLegacyPptWrite();
                Assert.True(report.CanWrite, string.Join(Environment.NewLine, report.Findings));
                savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation saved = LegacyPptPresentation.Load(savedBytes);
            Assert.Equal(2, saved.Slides.Count);
            LegacyPptInteraction interaction = Assert.Single(saved.Slides[1].Shapes
                .SelectMany(shape => shape.Interactions));
            Assert.Equal(target, interaction.Hyperlink!.Uri);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));
        }

        private static void AddShapeHyperlink(PowerPointSlide slide,
            PowerPointShape shape, Uri uri, bool mouseOver,
            string? screenTip = null) {
            HyperlinkRelationship relationship = slide.SlidePart
                .AddHyperlinkRelationship(uri, true);
            P.NonVisualDrawingProperties properties = GetDrawingProperties(shape);
            properties.Append(mouseOver
                ? new A.HyperlinkOnHover {
                    Id = relationship.Id,
                    Tooltip = screenTip
                }
                : new A.HyperlinkOnClick {
                    Id = relationship.Id,
                    Tooltip = screenTip
                });
        }

        private static P.NonVisualDrawingProperties GetDrawingProperties(
            PowerPointShape shape) => shape.Element switch {
                P.Shape item => item.NonVisualShapeProperties!.NonVisualDrawingProperties!,
                P.ConnectionShape item => item.NonVisualConnectionShapeProperties!
                    .NonVisualDrawingProperties!,
                P.Picture item => item.NonVisualPictureProperties!
                    .NonVisualDrawingProperties!,
                _ => throw new NotSupportedException()
            };

        private static P.Shape GetOnlyTextShape(PowerPointSlide slide) =>
            slide.SlidePart.Slide!.CommonSlideData!.ShapeTree!.Elements<P.Shape>()
                .Single(shape => shape.TextBody?.InnerText == "Visit OfficeIMO");

        private static void SetLinkedText(P.Shape shape, string relationshipId,
            bool linkFirstRun) {
            A.Run first = new(new A.Text("Visit "));
            A.Run second = new(new A.Text("OfficeIMO"));
            (linkFirstRun ? first : second).RunProperties = new A.RunProperties(
                new A.HyperlinkOnClick { Id = relationshipId });
            shape.TextBody!.RemoveAllChildren<A.Paragraph>();
            shape.TextBody.Append(new A.Paragraph(first, second));
        }
    }
}
