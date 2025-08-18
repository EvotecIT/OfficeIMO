using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Represents notes for a slide.
    /// </summary>
    public class PPNotes {
        private readonly SlidePart _slidePart;

        internal PPNotes(SlidePart slidePart) {
            _slidePart = slidePart;
        }

        private NotesSlide NotesSlide {
            get {
                if (_slidePart.NotesSlidePart == null) {
                    NotesSlidePart notesPart = _slidePart.AddNewPart<NotesSlidePart>();
                    notesPart.NotesSlide = new NotesSlide(
                        new CommonSlideData(new ShapeTree(
                            new Shape(
                                new NonVisualShapeProperties(
                                    new NonVisualDrawingProperties { Id = 1U, Name = "Notes Placeholder" },
                                    new NonVisualShapeDrawingProperties(),
                                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape())
                                ),
                                new ShapeProperties(),
                                new TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(new A.Run(new A.Text())))
                            )
                        )),
                        new ColorMapOverride(new A.MasterColorMapping()));
                }

                return _slidePart.NotesSlidePart!.NotesSlide;
            }
        }

        /// <summary>
        /// Gets or sets the notes text.
        /// </summary>
        public string Text {
            get {
                A.Run run = NotesSlide.CommonSlideData!.ShapeTree.GetFirstChild<Shape>()!
                    .TextBody!.GetFirstChild<A.Paragraph>()!
                    .GetFirstChild<A.Run>()!;
                A.Text text = run.GetFirstChild<A.Text>()!;
                return text.Text ?? string.Empty;
            }
            set {
                A.Run run = NotesSlide.CommonSlideData!.ShapeTree.GetFirstChild<Shape>()!
                    .TextBody!.GetFirstChild<A.Paragraph>()!
                    .GetFirstChild<A.Run>()!;
                A.Text text = run.GetFirstChild<A.Text>()!;
                text.Text = value;
            }
        }

        internal void Save() {
            _slidePart.NotesSlidePart?.NotesSlide?.Save();
        }
    }
}

