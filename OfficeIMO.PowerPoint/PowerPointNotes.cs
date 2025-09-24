using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents notes for a slide.
    /// </summary>
    public class PowerPointNotes {
        private readonly SlidePart _slidePart;
        private readonly Action _onChanged;

        internal PowerPointNotes(SlidePart slidePart, Action onChanged) {
            _slidePart = slidePart;
            _onChanged = onChanged;
        }

        private NotesSlide NotesSlide {
            get {
                if (_slidePart.NotesSlidePart == null) {
                    // Generate a unique relationship ID for the notes part
                    var slideRelationships = new HashSet<string>(
                        _slidePart.Parts.Select(p => p.RelationshipId)
                        .Union(_slidePart.ExternalRelationships.Select(r => r.Id))
                        .Union(_slidePart.HyperlinkRelationships.Select(r => r.Id))
                        .Where(id => !string.IsNullOrEmpty(id))
                    );
                    
                    int notesIdNum = 1;
                    string notesRelId;
                    do {
                        notesRelId = "rId" + notesIdNum;
                        notesIdNum++;
                    } while (slideRelationships.Contains(notesRelId));
                    
                    NotesSlidePart notesPart = _slidePart.AddNewPart<NotesSlidePart>(notesRelId);
                    PresentationPart? presentationPart = _slidePart.GetParentParts().OfType<PresentationPart>().FirstOrDefault();
                    if (presentationPart?.NotesMasterPart != null) {
                        notesPart.AddPart(presentationPart.NotesMasterPart);
                    }
                    notesPart.NotesSlide = new NotesSlide(
                        new CommonSlideData(new ShapeTree(
                            new Shape(
                                new NonVisualShapeProperties(
                                    new NonVisualDrawingProperties { Id = 1U, Name = "Notes Placeholder" },
                                    new NonVisualShapeDrawingProperties(),
                                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape())
                                ),
                                new ShapeProperties(),
                                new TextBody(new A.BodyProperties(), new A.ListStyle(),
                                    new A.Paragraph(new A.Run(new A.Text())))
                            )
                        )),
                        new ColorMapOverride(new A.MasterColorMapping()));
                    _onChanged();
                }

                return _slidePart.NotesSlidePart!.NotesSlide;
            }
        }

        /// <summary>
        ///     Gets or sets the notes text.
        /// </summary>
        public string Text {
            get {
                Shape? shape = NotesSlide.CommonSlideData?.ShapeTree?.GetFirstChild<Shape>();
                A.Paragraph? paragraph = shape?.TextBody?.GetFirstChild<A.Paragraph>();
                A.Run? run = paragraph?.GetFirstChild<A.Run>();
                A.Text? text = run?.GetFirstChild<A.Text>();
                return text?.Text ?? string.Empty;
            }
            set {
                NotesSlide notesSlide = NotesSlide;
                CommonSlideData common = notesSlide.CommonSlideData ??= new CommonSlideData(new ShapeTree());
                ShapeTree tree = common.ShapeTree ??= new ShapeTree();
                Shape? shape = tree.GetFirstChild<Shape>();
                if (shape is null) {
                    shape = new Shape(
                        new NonVisualShapeProperties(
                            new NonVisualDrawingProperties { Id = 1U, Name = "Notes Placeholder" },
                            new NonVisualShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties(new PlaceholderShape())
                        ),
                        new ShapeProperties(),
                        new TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph())
                    );
                    tree.AppendChild(shape);
                }

                A.Paragraph paragraph = shape.TextBody!.GetFirstChild<A.Paragraph>() ?? shape.TextBody.AppendChild(new A.Paragraph());
                A.Run run = paragraph.GetFirstChild<A.Run>() ?? paragraph.AppendChild(new A.Run());
                A.Text text = run.GetFirstChild<A.Text>() ?? run.AppendChild(new A.Text());
                text.Text = value;
                _onChanged();
            }
        }

        internal void Save() {
            _slidePart.NotesSlidePart?.NotesSlide?.Save();
        }
    }
}