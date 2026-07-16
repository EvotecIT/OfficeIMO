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
