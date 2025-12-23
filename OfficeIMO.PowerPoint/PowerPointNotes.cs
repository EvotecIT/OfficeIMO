using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    internal interface INotesMasterPartFactory {
        NotesMasterPart EnsureNotesMasterPart(PresentationPart presentationPart);
    }

    internal sealed class DefaultNotesMasterPartFactory : INotesMasterPartFactory {
        internal static DefaultNotesMasterPartFactory Instance { get; } = new();

        private DefaultNotesMasterPartFactory() {
        }

        public NotesMasterPart EnsureNotesMasterPart(PresentationPart presentationPart) {
            return PowerPointUtils.EnsureNotesMasterPart(presentationPart);
        }
    }

    /// <summary>
    ///     Represents notes for a slide.
    /// </summary>
    public class PowerPointNotes {
        private readonly SlidePart _slidePart;
        private readonly INotesMasterPartFactory _notesMasterPartFactory;
        private HashSet<string>? _cachedRelationshipIds;

        internal PowerPointNotes(SlidePart slidePart, INotesMasterPartFactory? notesMasterPartFactory = null) {
            _slidePart = slidePart;
            _notesMasterPartFactory = notesMasterPartFactory ?? DefaultNotesMasterPartFactory.Instance;
        }

        private NotesSlide NotesSlide {
            get {
                NotesSlidePart? notesPart = _slidePart.NotesSlidePart;
                if (notesPart == null) {
                    // Generate a unique relationship ID for the notes part
                    HashSet<string> slideRelationships = GetRelationshipIds();

                    int notesIdNum = 1;
                    string notesRelId;
                    do {
                        notesRelId = "rId" + notesIdNum;
                        notesIdNum++;
                    } while (!slideRelationships.Add(notesRelId));

                    notesPart = _slidePart.AddNewPart<NotesSlidePart>(notesRelId);

                    ShapeTree shapeTree = CreateEmptyShapeTree();
                    uint placeholderId = GetNextShapeId(shapeTree);
                    shapeTree.Append(CreateNotesPlaceholderShape(placeholderId));

                    notesPart.NotesSlide = new NotesSlide(
                        new CommonSlideData(shapeTree),
                        new ColorMapOverride(new A.MasterColorMapping()));
                }

                EnsureNotesMasterRelationship(notesPart);

                if (notesPart.NotesSlide == null) {
                    ShapeTree shapeTree = CreateEmptyShapeTree();
                    uint placeholderId = GetNextShapeId(shapeTree);
                    shapeTree.Append(CreateNotesPlaceholderShape(placeholderId));

                    notesPart.NotesSlide = new NotesSlide(
                        new CommonSlideData(shapeTree),
                        new ColorMapOverride(new A.MasterColorMapping()));
                }

                return notesPart.NotesSlide!;
            }
        }

        private void EnsureNotesMasterRelationship(NotesSlidePart notesPart) {
            PresentationPart? presentationPart = _slidePart
                .GetParentParts()
                .OfType<PresentationPart>()
                .FirstOrDefault();

            if (presentationPart == null) {
                return;
            }

            NotesMasterPart notesMasterPart = _notesMasterPartFactory.EnsureNotesMasterPart(presentationPart);

            bool hasNotesMasterRelationship = notesPart.Parts
                .Any(pair => ReferenceEquals(pair.OpenXmlPart, notesMasterPart));

            if (!hasNotesMasterRelationship) {
                notesPart.AddPart(notesMasterPart);
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
                CommonSlideData common = notesSlide.CommonSlideData ??= new CommonSlideData(CreateEmptyShapeTree());
                ShapeTree tree = EnsureShapeTree(common);
                Shape? shape = tree.GetFirstChild<Shape>();
                if (shape is null) {
                    uint placeholderId = GetNextShapeId(tree);
                    shape = CreateNotesPlaceholderShape(placeholderId);
                    tree.AppendChild(shape);
                }

                A.Paragraph paragraph = shape.TextBody!.GetFirstChild<A.Paragraph>() ?? shape.TextBody.AppendChild(new A.Paragraph());
                A.Run run = paragraph.GetFirstChild<A.Run>() ?? paragraph.AppendChild(new A.Run());
                A.Text text = run.GetFirstChild<A.Text>() ?? run.AppendChild(new A.Text());
                text.Text = value;
            }
        }

        internal void Save() {
            _slidePart.NotesSlidePart?.NotesSlide?.Save();
        }

        private static ShapeTree CreateEmptyShapeTree() {
            return new ShapeTree(
                new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties { Id = 1U, Name = "Notes Group Shape" },
                    new NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                PowerPointUtils.CreateDefaultGroupShapeProperties());
        }

        private static ShapeTree EnsureShapeTree(CommonSlideData commonSlideData) {
            ShapeTree tree = commonSlideData.ShapeTree ??= new ShapeTree();

            if (tree.GetFirstChild<NonVisualGroupShapeProperties>() == null) {
                tree.PrependChild(new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties { Id = 1U, Name = "Notes Group Shape" },
                    new NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()));
            }

            if (tree.GetFirstChild<GroupShapeProperties>() == null) {
                tree.AppendChild(PowerPointUtils.CreateDefaultGroupShapeProperties());
            }

            return tree;
        }

        private static Shape CreateNotesPlaceholderShape(uint id) {
            return new Shape(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties { Id = id, Name = "Notes Placeholder" },
                    new NonVisualShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape())
                ),
                new ShapeProperties(),
                new TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(new A.Run(new A.Text())))
            );
        }

        private static uint GetNextShapeId(ShapeTree shapeTree) {
            uint maxId = shapeTree
                .Descendants<NonVisualDrawingProperties>()
                .Select(properties => properties.Id?.Value ?? 0U)
                .DefaultIfEmpty(0U)
                .Max();

            return maxId + 1U;
        }

        private HashSet<string> GetRelationshipIds() {
            if (_cachedRelationshipIds == null) {
                _cachedRelationshipIds = new HashSet<string>(
                    _slidePart.Parts.Select(p => p.RelationshipId)
                        .Concat(_slidePart.ExternalRelationships.Select(r => r.Id))
                        .Concat(_slidePart.HyperlinkRelationships.Select(r => r.Id))
                        .Where(id => !string.IsNullOrEmpty(id)),
                    StringComparer.Ordinal);
            }

            return _cachedRelationshipIds;
        }
    }
}
