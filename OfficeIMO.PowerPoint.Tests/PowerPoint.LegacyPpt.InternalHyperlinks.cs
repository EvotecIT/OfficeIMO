using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
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
        public void NativeWriter_AuthorsAndProjectsInternalSlideHyperlink() {
            byte[] bytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                PowerPointSlide first = source.AddSlide(P.SlideLayoutValues.Blank);
                PowerPointSlide destination = source.AddSlide(P.SlideLayoutValues.Blank);
                destination.AddTextBox("Destination");
                AddInternalSlideHyperlink(first, destination, first.AddRectangle(
                    100000, 100000, 1000000, 500000), mouseOver: false,
                    screenTip: "Go to destination");

                LegacyPptWritePreflightReport report = source.AnalyzeLegacyPptWrite();
                Assert.True(report.CanWrite,
                    string.Join(Environment.NewLine, report.Findings));
                bytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation legacy = LegacyPptPresentation.Load(bytes);
            LegacyPptInteraction interaction = Assert.Single(legacy.Slides[0].Shapes
                .SelectMany(shape => shape.Interactions));
            Assert.Equal(LegacyPptInteractionAction.Hyperlink, interaction.Action);
            Assert.Equal(LegacyPptHyperlinkType.SlideNumber,
                interaction.HyperlinkType);
            LegacyPptHyperlink hyperlink = Assert.IsType<LegacyPptHyperlink>(
                interaction.Hyperlink);
            Assert.True(hyperlink.IsInternalSlideTarget);
            Assert.Equal(legacy.Slides[1].SlideId, hyperlink.TargetSlideId);
            Assert.Equal(2, hyperlink.TargetSlideNumber);
            Assert.Null(hyperlink.TargetSlideName);
            Assert.Null(hyperlink.Uri);
            Assert.Equal("Go to destination", hyperlink.ScreenTip);

            var security = OfficePackageSecurityOptions.SecureDefaults;
            security.ExternalRelationships =
                OfficePackageContentPolicy.Reject;
            using (var secureInput = new MemoryStream(bytes,
                       writable: false))
            using (PowerPointPresentation secure =
                   PowerPointPresentation.Load(secureInput,
                       new PowerPointLoadOptions {
                           PackageSecurity = security
                       })) {
                Assert.Equal(2, secure.Slides.Count);
            }

            using var input = new MemoryStream(bytes, writable: false);
            using PowerPointPresentation projected = PowerPointPresentation.Load(input);
            A.HyperlinkOnClick projectedLink = projected.Slides[0].SlidePart.Slide!
                .Descendants<A.HyperlinkOnClick>().Single();
            Assert.Equal("ppaction://hlinksldjump", projectedLink.Action?.Value);
            Assert.Equal("Go to destination", projectedLink.Tooltip?.Value);
            Assert.True(projected.Slides[0].SlidePart.TryGetPartById(
                projectedLink.Id!.Value!, out OpenXmlPart? targetPart));
            Assert.Same(projected.Slides[1].SlidePart, targetPart);
            Assert.Empty(projected.ValidateDocument());
            Assert.Equal(bytes, projected.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Fact]
        public void NativeWriter_AuthorsTextInternalSlideHyperlink() {
            byte[] bytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                PowerPointSlide first = source.AddSlide(P.SlideLayoutValues.Blank);
                PowerPointTextBox textBox = first.AddTextBox("Jump");
                PowerPointSlide destination = source.AddSlide(P.SlideLayoutValues.Blank);
                destination.AddTextBox("Destination");
                first.SlidePart.AddPart(destination.SlidePart);
                string relationshipId = first.SlidePart.GetIdOfPart(
                    destination.SlidePart);
                A.Run run = ((P.Shape)textBox.Element).TextBody!
                    .Descendants<A.Run>().Single();
                run.RunProperties ??= new A.RunProperties();
                run.RunProperties.Append(new A.HyperlinkOnClick {
                    Id = relationshipId,
                    Action = "ppaction://hlinksldjump",
                    Tooltip = "Jump to the destination"
                });
                Assert.True(source.AnalyzeLegacyPptWrite().CanWrite);
                bytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation legacy = LegacyPptPresentation.Load(bytes);
            LegacyPptTextInteraction interaction = Assert.Single(
                legacy.Slides[0].Shapes[0].TextBody.Interactions);
            Assert.Equal(legacy.Slides[1].SlideId,
                interaction.Interaction.Hyperlink!.TargetSlideId);
            Assert.Equal("Jump to the destination",
                interaction.Interaction.Hyperlink.ScreenTip);

            using var input = new MemoryStream(bytes, writable: false);
            using PowerPointPresentation projected = PowerPointPresentation.Load(input);
            A.HyperlinkOnClick projectedLink = projected.Slides[0].SlidePart.Slide!
                .Descendants<A.HyperlinkOnClick>().Single();
            Assert.Equal("ppaction://hlinksldjump", projectedLink.Action?.Value);
            Assert.True(projected.Slides[0].SlidePart.TryGetPartById(
                projectedLink.Id!.Value!, out OpenXmlPart? targetPart));
            Assert.Same(projected.Slides[1].SlidePart, targetPart);
        }

        [Fact]
        public void BinaryImport_ProjectsMasterInternalSlideHyperlinkAfterSlides() {
            byte[] bytes;
            using (PowerPointPresentation source =
                   PowerPointPresentation.Create()) {
                source.AddSlide(
                    P.SlideLayoutValues.Blank);
                SlideMasterPart masterPart = source.OpenXmlDocument
                    .PresentationPart!.SlideMasterParts.Single();
                masterPart.SlideMaster!.CommonSlideData!.ShapeTree!.Append(
                    new P.Shape(
                        new P.NonVisualShapeProperties(
                            new P.NonVisualDrawingProperties {
                                Id = 2U,
                                Name = "Master internal link"
                            },
                            new P.NonVisualShapeDrawingProperties(),
                            new P.ApplicationNonVisualDrawingProperties()),
                        new P.ShapeProperties(
                            new A.Transform2D(
                                new A.Offset { X = 100000, Y = 100000 },
                                new A.Extents {
                                    Cx = 1000000,
                                    Cy = 500000
                                }),
                            new A.PresetGeometry(
                                new A.AdjustValueList()) {
                                Preset = A.ShapeTypeValues.Rectangle
                            }),
                        new P.TextBody(
                            new A.BodyProperties(),
                            new A.ListStyle(),
                            new A.Paragraph(
                                new A.Run(new A.Text("Master link"))))));
                masterPart.SlideMaster.Save();
                bytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation legacy = LegacyPptPresentation.Load(bytes);
            LegacyPptMaster master = Assert.Single(legacy.Masters);
            LegacyPptShape sourceShape = Assert.Single(master.Shapes);
            LegacyPptSlide target = Assert.Single(legacy.Slides);
            var hyperlink = new LegacyPptHyperlink(1U, null, null,
                target.SlideId.ToString(
                    System.Globalization.CultureInfo.InvariantCulture)
                + ",1,Master target");
            hyperlink.ApplyExtension("Open target", 0U);
            var interaction = new LegacyPptInteraction(
                LegacyPptInteractionTrigger.MouseClick,
                LegacyPptInteractionAction.Hyperlink,
                LegacyPptInteractionJump.None,
                LegacyPptHyperlinkType.SlideNumber,
                soundIdReference: 0U, hyperlinkIdReference: 1U,
                oleVerb: 0, flags: 0, name: null, hyperlink,
                customShow: null);
            var linkedShape = new LegacyPptShape(sourceShape.Kind,
                sourceShape.OfficeArtShapeType, sourceShape.ShapeId,
                sourceShape.RecordOffset, sourceShape.Bounds,
                sourceShape.Text, sourceShape.Placeholder,
                sourceShape.Style, sourceShape.FillColor,
                sourceShape.LineColor, transform: sourceShape.Transform,
                textBody: sourceShape.TextBody.WithInteractions(new[] {
                    new LegacyPptTextInteraction(0,
                        sourceShape.Text.Length, interaction)
                }), interactions: new[] { interaction });
            List<LegacyPptShape> masterShapes = Assert.IsType<
                List<LegacyPptShape>>(master.Shapes);
            masterShapes[0] = linkedShape;

            using PowerPointPresentation projected =
                PowerPointPresentation.ProjectLoadedLegacyPpt(legacy,
                    sourcePath: null, PowerPointFileFormat.Ppt,
                    new PowerPointLoadOptions());
            SlideMasterPart projectedMaster = projected.OpenXmlDocument
                .PresentationPart!.SlideMasterParts.Single();
            A.HyperlinkOnClick[] links = projectedMaster.SlideMaster!
                .Descendants<A.HyperlinkOnClick>().ToArray();
            Assert.Equal(2, links.Length);
            Assert.All(links, link => {
                Assert.Equal("Open target", link.Tooltip?.Value);
                Assert.True(projectedMaster.TryGetPartById(
                    link.Id!.Value!, out OpenXmlPart? projectedTarget));
                Assert.Same(projected.Slides[0].SlidePart,
                    projectedTarget);
            });
            Assert.Empty(projected.ValidateDocument());
        }

        [Fact]
        public void ImportedInternalSlideHyperlink_RetargetsAndRemovesIncrementally() {
            byte[] sourceBytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                PowerPointSlide first = source.AddSlide(P.SlideLayoutValues.Blank);
                PowerPointSlide second = source.AddSlide(P.SlideLayoutValues.Blank);
                PowerPointSlide third = source.AddSlide(P.SlideLayoutValues.Blank);
                second.AddTextBox("Second destination");
                third.AddTextBox("Third destination");
                AddInternalSlideHyperlink(first, second, first.AddRectangle(
                    100000, 100000, 1000000, 500000), mouseOver: false);
                sourceBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original = LegacyPptPresentation.Load(sourceBytes);

            byte[] retargetedBytes;
            using (var input = new MemoryStream(sourceBytes, writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation.Load(input)) {
                PowerPointSlide first = imported.Slides[0];
                A.HyperlinkOnClick link = first.SlidePart.Slide!
                    .Descendants<A.HyperlinkOnClick>().Single();
                if (!first.SlidePart.Parts.Any(pair => ReferenceEquals(
                        pair.OpenXmlPart, imported.Slides[2].SlidePart))) {
                    first.SlidePart.AddPart(imported.Slides[2].SlidePart);
                }
                link.Id = first.SlidePart.GetIdOfPart(imported.Slides[2].SlidePart);
                LegacyPptWritePreflightReport report = imported.AnalyzeLegacyPptWrite();
                Assert.True(report.CanWrite,
                    string.Join(Environment.NewLine, report.Findings));
                retargetedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation retargeted = LegacyPptPresentation.Load(retargetedBytes);
            LegacyPptHyperlink retargetedLink = Assert.Single(retargeted.Slides[0].Shapes
                .SelectMany(shape => shape.Interactions)).Hyperlink!;
            Assert.Equal(retargeted.Slides[2].SlideId,
                retargetedLink.TargetSlideId);
            Assert.True(retargeted.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));

            byte[] removedBytes;
            using (var input = new MemoryStream(retargetedBytes, writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation.Load(input)) {
                GetDrawingProperties(Assert.Single(imported.Slides[0].Shapes))
                    .RemoveAllChildren<A.HyperlinkOnClick>();
                Assert.True(imported.AnalyzeLegacyPptWrite().CanWrite);
                removedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation removed = LegacyPptPresentation.Load(removedBytes);
            Assert.Empty(removed.Slides[0].Shapes
                .SelectMany(shape => shape.Interactions));
            Assert.True(removed.Package.DocumentStream.AsSpan(0,
                    retargeted.Package.DocumentStream.Length)
                .SequenceEqual(retargeted.Package.DocumentStream));
        }

        [Fact]
        public void ImportedPresentation_AppendedSlideLinksWorkInBothDirections() {
            byte[] sourceBytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                source.AddSlide(P.SlideLayoutValues.Blank).AddRectangle(
                    100000, 100000, 1000000, 500000);
                sourceBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original = LegacyPptPresentation.Load(sourceBytes);

            byte[] savedBytes;
            using (var input = new MemoryStream(sourceBytes, writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation.Load(input)) {
                PowerPointSlide existing = imported.Slides[0];
                PowerPointSlide appended = imported.AddSlide(P.SlideLayoutValues.Blank);
                PowerPointAutoShape appendedShape = appended.AddRectangle(
                    100000, 100000, 1000000, 500000);
                AddInternalSlideHyperlink(existing, appended,
                    Assert.Single(existing.Shapes), mouseOver: false,
                    screenTip: "Forward");
                AddInternalSlideHyperlink(appended, existing, appendedShape,
                    mouseOver: true, screenTip: "Back");
                LegacyPptWritePreflightReport report = imported.AnalyzeLegacyPptWrite();
                Assert.True(report.CanWrite,
                    string.Join(Environment.NewLine, report.Findings));
                savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation saved = LegacyPptPresentation.Load(savedBytes);
            Assert.Equal(2, saved.Slides.Count);
            LegacyPptHyperlink forward = Assert.Single(saved.Slides[0].Shapes
                .SelectMany(shape => shape.Interactions)).Hyperlink!;
            LegacyPptInteraction backInteraction = Assert.Single(saved.Slides[1].Shapes
                .SelectMany(shape => shape.Interactions));
            Assert.Equal(LegacyPptInteractionTrigger.MouseOver,
                backInteraction.Trigger);
            Assert.Equal(saved.Slides[1].SlideId, forward.TargetSlideId);
            Assert.Equal(saved.Slides[0].SlideId,
                backInteraction.Hyperlink!.TargetSlideId);
            Assert.Equal("Forward", forward.ScreenTip);
            Assert.Equal("Back", backInteraction.Hyperlink.ScreenTip);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));
        }

        [Fact]
        public void ImportedInternalSlideHyperlink_FollowsTargetAcrossSlideReorder() {
            byte[] sourceBytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                PowerPointSlide first = source.AddSlide(P.SlideLayoutValues.Blank);
                source.AddSlide(P.SlideLayoutValues.Blank).AddTextBox("Middle");
                PowerPointSlide destination = source.AddSlide(P.SlideLayoutValues.Blank);
                destination.AddTextBox("Destination");
                AddInternalSlideHyperlink(first, destination, first.AddRectangle(
                    100000, 100000, 1000000, 500000), mouseOver: false);
                sourceBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            byte[] savedBytes;
            using (var input = new MemoryStream(sourceBytes, writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation.Load(input)) {
                imported.MoveSlide(2, 1);
                Assert.True(imported.AnalyzeLegacyPptWrite().CanWrite);
                savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation saved = LegacyPptPresentation.Load(savedBytes);
            LegacyPptHyperlink link = Assert.Single(saved.Slides[0].Shapes
                .SelectMany(shape => shape.Interactions)).Hyperlink!;
            Assert.Equal("Destination", saved.Slides[1].Shapes.Single().Text);
            Assert.Equal(saved.Slides[1].SlideId, link.TargetSlideId);

            using var roundTripStream = new MemoryStream(savedBytes, writable: false);
            using PowerPointPresentation roundTrip = PowerPointPresentation.Load(
                roundTripStream);
            A.HyperlinkOnClick projectedLink = roundTrip.Slides[0].SlidePart.Slide!
                .Descendants<A.HyperlinkOnClick>().Single();
            Assert.True(roundTrip.Slides[0].SlidePart.TryGetPartById(
                projectedLink.Id!.Value!, out OpenXmlPart? targetPart));
            Assert.Same(roundTrip.Slides[1].SlidePart, targetPart);
        }

        private static void AddInternalSlideHyperlink(PowerPointSlide source,
            PowerPointSlide target, PowerPointShape shape, bool mouseOver,
            string? screenTip = null) {
            if (!source.SlidePart.Parts.Any(pair => ReferenceEquals(
                    pair.OpenXmlPart, target.SlidePart))) {
                source.SlidePart.AddPart(target.SlidePart);
            }
            string relationshipId = source.SlidePart.GetIdOfPart(target.SlidePart);
            P.NonVisualDrawingProperties properties = GetDrawingProperties(shape);
            properties.Append(mouseOver
                ? new A.HyperlinkOnHover {
                    Id = relationshipId,
                    Action = "ppaction://hlinksldjump",
                    Tooltip = screenTip
                }
                : new A.HyperlinkOnClick {
                    Id = relationshipId,
                    Action = "ppaction://hlinksldjump",
                    Tooltip = screenTip
                });
        }
    }
}
