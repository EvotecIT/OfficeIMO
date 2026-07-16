using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class PowerPointLegacyPptMasterTests {
        [Fact]
        public void NativeWriter_WritesNotesMasterDrawingsAndPlaceholders() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            NotesMasterPart notesMasterPart = presentation.OpenXmlDocument
                .PresentationPart!.NotesMasterPart!;
            P.ShapeTree tree = notesMasterPart.NotesMaster!.CommonSlideData!
                .ShapeTree!;
            foreach (var child in tree.ChildElements.Where(child =>
                         child is not P.NonVisualGroupShapeProperties
                             and not P.GroupShapeProperties).ToArray()) {
                child.Remove();
            }
            var slideBounds = new PowerPointLayoutBox(400000, 300000,
                4200000, 2500000);
            var bodyBounds = new PowerPointLayoutBox(500000, 3300000,
                8100000, 2800000);
            tree.Append(
                CreateNotesMasterShape(2U, "Notes slide image", slideBounds,
                    P.PlaceholderValues.SlideImage),
                CreateNotesMasterShape(3U, "Notes body", bodyBounds,
                    P.PlaceholderValues.Body, "Notes body default"),
                CreateNotesMasterShape(4U, "Notes decoration",
                    new PowerPointLayoutBox(8200000, 400000, 400000, 400000),
                    placeholder: null, text: null,
                    shapeType: A.ShapeTypeValues.Ellipse));
            presentation.AddSlide(P.SlideLayoutValues.Blank);

            LegacyPptWritePreflightReport preflight = presentation
                .AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));
            byte[] bytes = presentation.ToBytes(PowerPointFileFormat.Ppt);
            LegacyPptSpecialMaster notesMaster = Assert.IsType<
                LegacyPptSpecialMaster>(LegacyPptPresentation.Load(bytes)
                .NotesMaster);

            Assert.Equal(3, notesMaster.Shapes.Count);
            LegacyPptShape slideImage = Assert.Single(notesMaster.Shapes,
                shape => shape.Placeholder?.Kind
                    == LegacyPptPlaceholderKind.MasterNotesSlideImage);
            Assert.Equal(LegacyPptShapeKind.Rectangle, slideImage.Kind);
            Assert.Equal(ToEmus(ToMasterUnits(slideBounds.Left)),
                ToEmus(slideImage.Bounds.Left));
            LegacyPptShape body = Assert.Single(notesMaster.Shapes,
                shape => shape.Placeholder?.Kind
                    == LegacyPptPlaceholderKind.MasterNotesBody);
            Assert.Equal("Notes body default", body.Text);
            Assert.Equal(LegacyPptTextType.Notes, body.TextBody.TextType);
            LegacyPptShape decoration = Assert.Single(notesMaster.Shapes,
                shape => shape.Placeholder == null);
            Assert.Equal(LegacyPptShapeKind.Ellipse, decoration.Kind);
            Assert.All(notesMaster.Shapes, shape => Assert.Equal(
                slideImage.ShapeId >> 10, shape.ShapeId >> 10));

            using var stream = new MemoryStream(bytes);
            using PowerPointPresentation reopened = PowerPointPresentation.Load(
                stream);
            P.Shape[] projected = reopened.OpenXmlDocument.PresentationPart!
                .NotesMasterPart!.NotesMaster!.CommonSlideData!.ShapeTree!
                .Elements<P.Shape>().ToArray();
            Assert.Equal(3, projected.Length);
            Assert.Contains(projected, shape => shape.NonVisualShapeProperties?
                .ApplicationNonVisualDrawingProperties?.PlaceholderShape?
                .Type?.Value == P.PlaceholderValues.SlideImage);
            Assert.Contains(projected, shape => shape.TextBody?.InnerText
                == "Notes body default");
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void NativeWriter_RejectsUnsupportedNotesMasterDrawing() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            presentation.OpenXmlDocument.PresentationPart!.NotesMasterPart!
                .NotesMaster!.CommonSlideData!.ShapeTree!
                .Append(new P.GraphicFrame());
            presentation.AddSlide(P.SlideLayoutValues.Blank);

            LegacyPptWritePreflightReport preflight = presentation
                .AnalyzeLegacyPptWrite();

            Assert.False(preflight.CanWrite);
            Assert.Contains(preflight.Findings, finding =>
                finding.Code == "PPT-WRITE-NOTES-MASTER-SHAPE");
        }

        [Fact]
        public void NativeWriter_WritesHandoutMasterTopologyDrawingsAndBackground() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            HandoutMasterPart handoutMasterPart = CreateHandoutMaster(presentation);
            P.CommonSlideData commonSlideData = handoutMasterPart.HandoutMaster!
                .CommonSlideData!;
            P.ShapeTree tree = commonSlideData.ShapeTree!;
            var dateBounds = new PowerPointLayoutBox(350000, 300000,
                2600000, 450000);
            var headerBounds = new PowerPointLayoutBox(3200000, 300000,
                5700000, 450000);
            tree.Append(
                CreateNotesMasterShape(2U, "Handout date", dateBounds,
                    P.PlaceholderValues.DateAndTime, "July 2026"),
                CreateNotesMasterShape(3U, "Handout header", headerBounds,
                    P.PlaceholderValues.Header, "Quarterly review"),
                CreateNotesMasterShape(4U, "Handout footer",
                    new PowerPointLayoutBox(350000, 6200000, 6000000, 450000),
                    P.PlaceholderValues.Footer, "Confidential"),
                CreateNotesMasterShape(5U, "Handout slide number",
                    new PowerPointLayoutBox(8300000, 6200000, 600000, 450000),
                    P.PlaceholderValues.SlideNumber, "1"),
                CreateNotesMasterShape(6U, "Handout decoration",
                    new PowerPointLayoutBox(8450000, 950000, 400000, 400000),
                    placeholder: null, text: null,
                    shapeType: A.ShapeTypeValues.Ellipse));
            commonSlideData.Background = new P.Background(
                new P.BackgroundProperties(
                    new A.SolidFill(
                        new A.RgbColorModelHex { Val = "334455" })));
            PowerPointSlide slide = presentation.AddSlide(
                P.SlideLayoutValues.Blank);
            slide.Notes.Text = "Handout topology note";

            LegacyPptWritePreflightReport preflight = presentation
                .AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));
            byte[] bytes = presentation.ToBytes(PowerPointFileFormat.Ppt);
            LegacyPptPresentation binary = LegacyPptPresentation.Load(bytes);
            Assert.Equal(16U, binary.DocumentSettings!.HandoutMasterPersistId);
            LegacyPptSpecialMaster handoutMaster = Assert.IsType<
                LegacyPptSpecialMaster>(binary.HandoutMaster);

            Assert.Equal(LegacyPptSpecialMasterKind.Handout,
                handoutMaster.Kind);
            Assert.Equal(16U, handoutMaster.PersistId);
            Assert.Equal(5, handoutMaster.Shapes.Count);
            Assert.Equal("334455", Assert.IsType<LegacyPptBackground>(
                handoutMaster.Background).ForegroundColor);
            LegacyPptShape date = Assert.Single(handoutMaster.Shapes,
                shape => shape.Placeholder?.Kind
                    == LegacyPptPlaceholderKind.MasterDate);
            Assert.Equal("July 2026", date.Text);
            Assert.Equal(ToEmus(ToMasterUnits(dateBounds.Left)),
                ToEmus(date.Bounds.Left));
            Assert.Equal("Quarterly review", Assert.Single(
                handoutMaster.Shapes, shape => shape.Placeholder?.Kind
                    == LegacyPptPlaceholderKind.MasterHeader).Text);
            Assert.Equal("Confidential", Assert.Single(
                handoutMaster.Shapes, shape => shape.Placeholder?.Kind
                    == LegacyPptPlaceholderKind.MasterFooter).Text);
            Assert.Equal("1", Assert.Single(handoutMaster.Shapes,
                shape => shape.Placeholder?.Kind
                    == LegacyPptPlaceholderKind.MasterSlideNumber).Text);
            Assert.Equal(LegacyPptShapeKind.Ellipse, Assert.Single(
                handoutMaster.Shapes, shape => shape.Placeholder == null).Kind);
            Assert.All(handoutMaster.Shapes, shape => Assert.Equal(
                date.ShapeId >> 10, shape.ShapeId >> 10));

            using var stream = new MemoryStream(bytes);
            using PowerPointPresentation reopened = PowerPointPresentation.Load(
                stream);
            HandoutMasterPart projectedPart = Assert.IsType<HandoutMasterPart>(
                reopened.OpenXmlDocument.PresentationPart!.HandoutMasterPart);
            P.Shape[] projected = projectedPart.HandoutMaster!.CommonSlideData!
                .ShapeTree!.Elements<P.Shape>().ToArray();
            Assert.Equal(5, projected.Length);
            Assert.Contains(projected, shape => shape.NonVisualShapeProperties?
                .ApplicationNonVisualDrawingProperties?.PlaceholderShape?
                .Type?.Value == P.PlaceholderValues.Header);
            Assert.Contains(projected, shape => shape.TextBody?.InnerText
                == "Quarterly review");
            A.SolidFill projectedBackground = Assert.IsType<A.SolidFill>(
                projectedPart.HandoutMaster.CommonSlideData.Background!
                    .BackgroundProperties!.GetFirstChild<A.SolidFill>());
            Assert.Equal("334455",
                projectedBackground.RgbColorModelHex!.Val!.Value);
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void NativeWriter_RejectsUnsupportedHandoutMasterDrawing() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            HandoutMasterPart handoutMasterPart = CreateHandoutMaster(presentation);
            handoutMasterPart.HandoutMaster!.CommonSlideData!.ShapeTree!
                .Append(new P.GraphicFrame());
            presentation.AddSlide(P.SlideLayoutValues.Blank);

            LegacyPptWritePreflightReport preflight = presentation
                .AnalyzeLegacyPptWrite();

            Assert.False(preflight.CanWrite);
            Assert.Contains(preflight.Findings, finding =>
                finding.Code == "PPT-WRITE-HANDOUT-MASTER-SHAPE");
        }

        private static HandoutMasterPart CreateHandoutMaster(
            PowerPointPresentation presentation) {
            PresentationPart presentationPart = presentation.OpenXmlDocument
                .PresentationPart!;
            HandoutMasterPart handoutMasterPart = presentationPart
                .AddNewPart<HandoutMasterPart>();
            P.ShapeTree notesTree = presentationPart.NotesMasterPart!
                .NotesMaster!.CommonSlideData!.ShapeTree!;
            var tree = new P.ShapeTree(
                notesTree.GetFirstChild<P.NonVisualGroupShapeProperties>()!
                    .CloneNode(true),
                notesTree.GetFirstChild<P.GroupShapeProperties>()!
                    .CloneNode(true));
            P.ColorMap colorMap = (P.ColorMap)presentationPart
                .SlideMasterParts.First().SlideMaster!.ColorMap!
                .CloneNode(true);
            handoutMasterPart.HandoutMaster = new P.HandoutMaster(
                new P.CommonSlideData(tree), colorMap);
            ThemePart themePart = handoutMasterPart.AddNewPart<ThemePart>();
            themePart.Theme = (A.Theme)presentationPart.SlideMasterParts.First()
                .ThemePart!.Theme!.CloneNode(true);

            P.Presentation root = presentationPart.Presentation!;
            P.HandoutMasterIdList list = root.HandoutMasterIdList
                ??= new P.HandoutMasterIdList();
            var id = new P.HandoutMasterId();
            PowerPointUtils.SetRelationshipIdValue(id,
                presentationPart.GetIdOfPart(handoutMasterPart));
            list.Append(id);
            return handoutMasterPart;
        }

        private static P.Shape CreateNotesMasterShape(uint id, string name,
            PowerPointLayoutBox bounds, P.PlaceholderValues? placeholder,
            string? text = null,
            A.ShapeTypeValues? shapeType = null) {
            var applicationProperties =
                new P.ApplicationNonVisualDrawingProperties();
            if (placeholder.HasValue) {
                applicationProperties.Append(new P.PlaceholderShape {
                    Type = placeholder.Value,
                    Index = id - 2U
                });
            }
            var shape = new P.Shape(
                new P.NonVisualShapeProperties(
                    new P.NonVisualDrawingProperties { Id = id, Name = name },
                    new P.NonVisualShapeDrawingProperties(),
                    applicationProperties),
                new P.ShapeProperties(
                    new A.Transform2D(
                        new A.Offset { X = bounds.Left, Y = bounds.Top },
                        new A.Extents {
                            Cx = bounds.Width,
                            Cy = bounds.Height
                        }),
                    new A.PresetGeometry(new A.AdjustValueList()) {
                        Preset = shapeType ?? A.ShapeTypeValues.Rectangle
                    }));
            if (text != null) {
                shape.Append(new P.TextBody(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(
                        new A.Run(new A.Text(text)),
                        new A.EndParagraphRunProperties())));
            }
            return shape;
        }
    }
}
