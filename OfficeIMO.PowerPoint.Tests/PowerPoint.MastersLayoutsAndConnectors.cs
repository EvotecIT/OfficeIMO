using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public class PowerPointMastersLayoutsAndConnectorsTests {
        private static readonly OfficeColor Band = OfficeColor.FromRgb(200, 30, 40);

        [Fact]
        public void MasterAndLayoutAuthoring_SurvivesSaveWithUniqueIdsAndRemapsRemovedLayouts() {
            using var stream = new MemoryStream();
            using (PowerPointDocument presentation = PowerPointDocument.Create(stream)) {
                PowerPointSlideMaster master = presentation.SlideMasters[0];
                master.AddShape(OfficePresetShapeType.Rectangle, OfficeLength.Points(0), OfficeLength.Points(0),
                    OfficeLength.Points(720), OfficeLength.Points(18), "Brand band").Fill(Band);

                PowerPointSlideLayout agenda = master.AddLayout("Agenda");
                agenda.AddPlaceholder(PowerPointPlaceholderType.Title,
                    new PowerPointLayoutBox(OfficeLength.Points(40), OfficeLength.Points(30), OfficeLength.Points(640), OfficeLength.Points(60)));
                agenda.AddPlaceholder(PowerPointPlaceholderType.Picture,
                    new PowerPointLayoutBox(OfficeLength.Points(400), OfficeLength.Points(110), OfficeLength.Points(280), OfficeLength.Points(240)));
                agenda.AddPlaceholder(PowerPointPlaceholderType.SlideNumber,
                    new PowerPointLayoutBox(OfficeLength.Points(620), OfficeLength.Points(500), OfficeLength.Points(60), OfficeLength.Points(20)));
                agenda.BackgroundColor = OfficeColor.FromRgb(245, 245, 245);

                PowerPointSlideLayout copy = agenda.Duplicate("Agenda (dark)");
                copy.BackgroundColor = OfficeColor.FromRgb(20, 20, 20);
                presentation.AddSlide(copy);
                presentation.AddSlide(copy);

                PowerPointSlideMaster second = presentation.AddSlideMaster("Appendix");
                Assert.True(second.Layouts.Count > 1);

                copy.Remove(agenda);
                Assert.All(presentation.Slides, slide => Assert.Same(agenda, slide.Layout));
                Assert.Throws<InvalidOperationException>(() => agenda.Remove());
                Assert.Throws<InvalidOperationException>(() => copy.Name = "removed");

                Assert.Empty(presentation.ValidateDocument());
                presentation.Save();
            }

            stream.Position = 0;
            using (PresentationDocument package = PresentationDocument.Open(stream, false)) {
                PresentationPart presentationPart = package.PresentationPart!;
                uint[] ids = presentationPart.Presentation.SlideMasterIdList!.Elements<SlideMasterId>().Select(id => id.Id!.Value)
                    .Concat(presentationPart.SlideMasterParts.SelectMany(part =>
                        part.SlideMaster.SlideLayoutIdList!.Elements<SlideLayoutId>().Select(id => id.Id!.Value)))
                    .ToArray();
                Assert.Equal(ids.Length, ids.Distinct().Count());
                Assert.All(presentationPart.SlideMasterParts, master => Assert.Equal(
                    master.SlideLayoutParts.Count(), master.SlideMaster.SlideLayoutIdList!.Elements<SlideLayoutId>().Count()));
            }

            stream.Position = 0;
            using (PowerPointDocument reopened = PowerPointDocument.Load(stream)) {
                Assert.Equal(new[] { "Appendix" }, reopened.SlideMasters.Skip(1).Select(master => master.Name));
                PowerPointSlideLayout layout = reopened.Slides[0].Layout!;
                Assert.Equal("Agenda", layout.Name);
                Assert.True(layout.Index == reopened.Slides[0].LayoutIndex);
                Assert.DoesNotContain(reopened.SlideMasters[0].Layouts, candidate => candidate.Name == "Agenda (dark)");
                Assert.Equal(new[] { PowerPointPlaceholderType.Title, PowerPointPlaceholderType.Picture, PowerPointPlaceholderType.SlideNumber },
                    layout.Placeholders.Select(placeholder => placeholder.PlaceholderType!.Value));
                Assert.Contains(reopened.SlideMasters[0].Shapes, shape => shape.Name == "Brand band");
            }
        }

        [Fact]
        public void MasterAndLayoutContent_IsInheritedBySlidesAndRendered() {
            using var stream = new MemoryStream();
            using PowerPointDocument presentation = PowerPointDocument.Create(stream);
            presentation.SlideSize.SetSize(OfficeLength.Points(400), OfficeLength.Points(300));
            PowerPointSlideMaster master = presentation.SlideMasters[0];
            master.AddShape(OfficePresetShapeType.Rectangle, OfficeLength.Points(0), OfficeLength.Points(0),
                OfficeLength.Points(400), OfficeLength.Points(20)).Fill(Band);
            PowerPointSlideLayout layout = master.AddLayout("Plain");
            PowerPointTextBox body = layout.AddPlaceholder(PowerPointPlaceholderType.Body,
                new PowerPointLayoutBox(OfficeLength.Points(20), OfficeLength.Points(60), OfficeLength.Points(200), OfficeLength.Points(100)));
            layout.AddShape(OfficePresetShapeType.Ellipse, OfficeLength.Points(330), OfficeLength.Points(230),
                OfficeLength.Points(50), OfficeLength.Points(50)).Fill(OfficeColor.FromRgb(20, 120, 40));
            PowerPointSlide slide = presentation.AddSlide(layout);

            body.Bounds = new PowerPointLayoutBox(OfficeLength.Points(40), OfficeLength.Points(80), OfficeLength.Points(300), OfficeLength.Points(120));
            Assert.Equal(OfficeLength.Points(300).ToEmus(), slide.GetLayoutPlaceholderBounds(PowerPointPlaceholderType.Body, body.PlaceholderIndex)!.Value.Width.ToEmus());

            OfficeImageExportResult png = slide.ToImage(OfficeImageExportFormat.Png, new PowerPointImageExportOptions());
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            double scale = image!.Width / 400D;
            Assert.Equal(Band, image.GetPixel((int)(200 * scale), (int)(10 * scale)));
            Assert.Equal(OfficeColor.FromRgb(20, 120, 40), image.GetPixel((int)(355 * scale), (int)(255 * scale)));
        }

        [Fact]
        public void BindToPlaceholder_ReplacesEmptySlidePlaceholderAndKeepsTheLayoutBinding() {
            using var stream = new MemoryStream();
            string imagePath = Path.Combine(AppContext.BaseDirectory, "Images", "BackgroundImage.png");
            uint pictureIndex;
            using (PowerPointDocument presentation = PowerPointDocument.Create(stream)) {
                PowerPointSlideLayout layout = presentation.SlideMasters[0].AddLayout("Picture and table");
                PowerPointTextBox layoutTitle = layout.AddPlaceholder(PowerPointPlaceholderType.Title,
                    new PowerPointLayoutBox(OfficeLength.Points(20), OfficeLength.Points(330), OfficeLength.Points(600), OfficeLength.Points(40)));
                PowerPointLayoutBox pictureBox = new(OfficeLength.Points(20), OfficeLength.Points(20), OfficeLength.Points(200), OfficeLength.Points(300));
                pictureIndex = layout.AddPlaceholder(PowerPointPlaceholderType.Picture, pictureBox).PlaceholderIndex!.Value;
                layout.AddPlaceholder(PowerPointPlaceholderType.Object,
                    new PowerPointLayoutBox(OfficeLength.Points(260), OfficeLength.Points(20), OfficeLength.Points(400), OfficeLength.Points(200)));
                PowerPointSlide slide = presentation.AddSlide(layout);
                int placeholdersBefore = slide.Shapes.Count(shape => shape.ShapePlaceholderIndex == pictureIndex);

                PowerPointShape picture = slide.AddImage(imagePath).BindToPlaceholder(PowerPointPlaceholderType.Picture);
                slide.AddTable(2, 2).BindToPlaceholder(PowerPointPlaceholderType.Object);
                PowerPointShape title = slide.AddTextBox("Review").BindToPlaceholder(PowerPointPlaceholderType.Title);
                Assert.Empty(title.ConnectionSites);
                Assert.Throws<NotSupportedException>(() => slide.AddConnector(title, picture, PowerPointConnectorKind.Straight));
                // The bound title has no local geometry, so moving the layout placeholder moves it.
                layoutTitle.Top = OfficeLength.Points(350);
                Assert.True(slide.GetPlaceholder(PowerPointPlaceholderType.Title)!.TryGetExportBoundsPoints(out _, out double titleTop, out _, out _));
                Assert.Equal(350D, titleTop, 3);

                Assert.Equal(1, slide.Shapes.Count(shape => shape.ShapePlaceholderIndex == pictureIndex));
                Assert.True(placeholdersBefore <= 1);
                Assert.Equal(pictureBox.Width.ToEmus(), picture.Width.ToEmus());
                Assert.Equal(pictureBox.Left.ToEmus(), picture.Left.ToEmus());
                Assert.Empty(presentation.ValidateDocument());
                presentation.Save();
            }

            stream.Position = 0;
            using (PowerPointDocument reopened = PowerPointDocument.Load(stream)) {
                PowerPointSlide slide = reopened.Slides[0];
                PowerPointImage image = Assert.Single(slide.Images);
                Assert.Equal(PowerPointPlaceholderType.Picture, image.ShapePlaceholderType);
                Assert.Equal(pictureIndex, image.ShapePlaceholderIndex);
                PowerPointTable table = Assert.Single(slide.Tables);
                Assert.NotNull(table.ShapePlaceholderIndex);
                Assert.Equal("Review", slide.GetPlaceholder(PowerPointPlaceholderType.Title)!.Text);
            }
        }

        [Fact]
        public void GluedConnector_FollowsShapeMovesAndKeepsGlueAfterReopen() {
            using var stream = new MemoryStream();
            using (PowerPointDocument presentation = PowerPointDocument.Create(stream)) {
                presentation.SlideSize.SetSize(OfficeLength.Points(400), OfficeLength.Points(300));
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape left = slide.AddRectangle(OfficeLength.Points(20), OfficeLength.Points(40), OfficeLength.Points(80), OfficeLength.Points(40));
                PowerPointAutoShape right = slide.AddShape(OfficePresetShapeType.Ellipse, OfficeLength.Points(260), OfficeLength.Points(180),
                    OfficeLength.Points(80), OfficeLength.Points(60));
                Assert.Equal(4, left.ConnectionSites.Count);
                Assert.Equal(8, right.ConnectionSites.Count);

                PowerPointConnectionShape elbow = slide.AddConnector(left, 3, right, 0, PowerPointConnectorKind.Elbow);
                elbow.SetLineEnds(null, OfficeLineMarkerKind.Triangle);
                AssertEnds(elbow, left.ConnectionSites[3], right.ConnectionSites[0]);
                PowerPointConnectionShape auto = slide.AddConnector(left, right, PowerPointConnectorKind.Curved);

                right.MoveBy(OfficeLength.Points(-40), OfficeLength.Points(20));
                left.Rotation = 90;
                AssertEnds(elbow, left.ConnectionSites[3], right.ConnectionSites[0]);
                Assert.Equal(PowerPointConnectorKind.Curved, auto.Kind);
                Assert.Same(right, auto.EndShape);

                OfficeImageExportResult png = slide.ToImage(OfficeImageExportFormat.Png, new PowerPointImageExportOptions());
                Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Message.Contains("connector", StringComparison.OrdinalIgnoreCase));
                Assert.Empty(presentation.ValidateDocument());
                presentation.Save();
            }

            stream.Position = 0;
            using (PowerPointDocument reopened = PowerPointDocument.Load(stream)) {
                PowerPointSlide slide = reopened.Slides[0];
                PowerPointConnectionShape elbow = slide.Shapes.OfType<PowerPointConnectionShape>().First();
                Assert.Equal(PowerPointConnectorKind.Elbow, elbow.Kind);
                Assert.Equal(3, elbow.StartSite);
                Assert.Equal(0, elbow.EndSite);
                PowerPointShape right = elbow.EndShape!;
                right.Left = OfficeLength.Points(300);
                AssertEnds(elbow, elbow.StartShape!.ConnectionSites[3], right.ConnectionSites[0]);
            }
        }

        [Fact]
        public void MovingAShapeGluedWithoutKnownSites_LeavesTheConnectorAsStored() {
            using var stream = new MemoryStream();
            using (PowerPointDocument presentation = PowerPointDocument.Create(stream)) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape left = slide.AddRectangle(OfficeLength.Points(20), OfficeLength.Points(40), OfficeLength.Points(80), OfficeLength.Points(40));
                PowerPointAutoShape right = slide.AddRectangle(OfficeLength.Points(260), OfficeLength.Points(180), OfficeLength.Points(80), OfficeLength.Points(60));
                slide.AddConnector(left, 3, right, 1, PowerPointConnectorKind.Elbow);
                // Another producer glued the connector to a preset whose sites this library does not know.
                ((Shape)right.Element).ShapeProperties!.GetFirstChild<A.PresetGeometry>()!.Preset = A.ShapeTypeValues.Hexagon;
                presentation.Save();
            }

            stream.Position = 0;
            using PowerPointDocument reopened = PowerPointDocument.Load(stream);
            PowerPointSlide reloaded = reopened.Slides[0];
            PowerPointConnectionShape connector = reloaded.Shapes.OfType<PowerPointConnectionShape>().Single();
            string before = connector.Element.OuterXml;
            PowerPointShape hexagon = connector.EndShape!;

            hexagon.Bounds = new PowerPointLayoutBox(OfficeLength.Points(300), OfficeLength.Points(200), OfficeLength.Points(90), OfficeLength.Points(70));

            Assert.Equal(OfficeLength.Points(90).ToEmus(), hexagon.Width.ToEmus());
            Assert.Equal(before, connector.Element.OuterXml);
            Assert.Throws<NotSupportedException>(() => connector.Reroute());
        }

        [Fact]
        public void RemovingTheLastLayoutOfAMaster_FailsWithoutMovingItsSlides() {
            using var stream = new MemoryStream();
            using PowerPointDocument presentation = PowerPointDocument.Create(stream);
            PowerPointSlideMaster second = presentation.AddSlideMaster("Second");
            foreach (PowerPointSlideLayout extra in second.Layouts.Skip(1).ToList()) extra.Remove();
            PowerPointSlideLayout only = Assert.Single(second.Layouts);
            PowerPointSlide slide = presentation.AddSlide(only);

            Assert.Throws<InvalidOperationException>(() => only.Remove(presentation.SlideMasters[0].Layouts[0]));
            Assert.Same(only, slide.Layout);
        }

        private static void AssertEnds(PowerPointConnectionShape connector, PowerPointConnectionSite start, PowerPointConnectionSite end) {
            var element = (ConnectionShape)connector.Element;
            A.Transform2D transform = element.ShapeProperties!.Transform2D!;
            double x = transform.Offset!.X!.Value, y = transform.Offset.Y!.Value;
            double cx = transform.Extents!.Cx!.Value, cy = transform.Extents.Cy!.Value;
            bool rotated = transform.Rotation?.Value == 5400000;
            double localX = (transform.HorizontalFlip?.Value == true ? 1 : -1) * cx / 2, localY = (transform.VerticalFlip?.Value == true ? 1 : -1) * cy / 2;
            double sx = rotated ? -localY : localX, sy = rotated ? localX : localY;
            double centerX = x + cx / 2, centerY = y + cy / 2;
            Assert.InRange(centerX + sx, start.X.ToEmus() - 2, start.X.ToEmus() + 2);
            Assert.InRange(centerY + sy, start.Y.ToEmus() - 2, start.Y.ToEmus() + 2);
            Assert.InRange(centerX - sx, end.X.ToEmus() - 2, end.X.ToEmus() + 2);
            Assert.InRange(centerY - sy, end.Y.ToEmus() - 2, end.Y.ToEmus() + 2);
        }
    }
}
