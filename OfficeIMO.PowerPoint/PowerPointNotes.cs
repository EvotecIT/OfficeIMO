using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
                Shape? shape = GetNotesTextShape(NotesSlide.CommonSlideData?.ShapeTree);
                if (shape?.TextBody == null) {
                    return string.Empty;
                }

                List<string> paragraphs = shape.TextBody.Elements<A.Paragraph>()
                    .Select(ReadParagraphText)
                    .ToList();
                return paragraphs.Count == 0 ? string.Empty : string.Join(Environment.NewLine, paragraphs);
            }
            set {
                NotesSlide notesSlide = NotesSlide;
                CommonSlideData common = notesSlide.CommonSlideData ??= new CommonSlideData(CreateEmptyShapeTree());
                ShapeTree tree = EnsureShapeTree(common);
                Shape shape = GetOrCreateNotesTextShape(tree);
                TextBody textBody = shape.TextBody ?? new TextBody(new A.BodyProperties(), new A.ListStyle());
                shape.TextBody ??= textBody;

                A.Paragraph? templateParagraph = textBody.Elements<A.Paragraph>().FirstOrDefault();
                A.ParagraphProperties? templateParagraphProperties = templateParagraph?.GetFirstChild<A.ParagraphProperties>();
                A.EndParagraphRunProperties? templateEndParagraphRunProperties = templateParagraph?.GetFirstChild<A.EndParagraphRunProperties>();
                A.RunProperties? templateRunProperties = templateParagraph?
                    .Elements<A.Run>()
                    .Select(run => run.RunProperties)
                    .FirstOrDefault(properties => properties != null);

                textBody.RemoveAllChildren<A.Paragraph>();

                string[] lines = (value ?? string.Empty).Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
                foreach (string line in lines) {
                    A.Paragraph paragraph = new();
                    if (templateParagraphProperties != null) {
                        paragraph.Append((A.ParagraphProperties)templateParagraphProperties.CloneNode(true));
                    }

                    A.Run run = new();
                    if (templateRunProperties != null) {
                        run.RunProperties = (A.RunProperties)templateRunProperties.CloneNode(true);
                    }

                    run.Append(new A.Text(line));
                    paragraph.Append(run);

                    if (templateEndParagraphRunProperties != null) {
                        paragraph.Append((A.EndParagraphRunProperties)templateEndParagraphRunProperties.CloneNode(true));
                    }

                    textBody.Append(paragraph);
                }
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
                    new NonVisualDrawingProperties { Id = id, Name = "Notes Placeholder 1" },
                    new NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties(
                        new PlaceholderShape { Type = PlaceholderValues.Body, Index = 1U })
                ),
                new ShapeProperties(),
                new TextBody(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(
                        new A.Run(
                            new A.RunProperties { Language = "en-US" },
                            new A.Text()),
                        new A.EndParagraphRunProperties { Language = "en-US" }))
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

        private static Shape? GetNotesTextShape(ShapeTree? shapeTree) {
            if (shapeTree == null) {
                return null;
            }

            return shapeTree.Elements<Shape>().FirstOrDefault(IsNotesTextShape)
                ?? shapeTree.Elements<Shape>().FirstOrDefault(shape => shape.TextBody != null);
        }

        private static Shape GetOrCreateNotesTextShape(ShapeTree shapeTree) {
            Shape? shape = GetNotesTextShape(shapeTree);
            if (shape != null) {
                return shape;
            }

            uint placeholderId = GetNextShapeId(shapeTree);
            shape = CreateNotesPlaceholderShape(placeholderId);
            shapeTree.AppendChild(shape);
            return shape;
        }

        private static bool IsNotesTextShape(Shape shape) {
            PlaceholderShape? placeholder = shape.NonVisualShapeProperties?
                .ApplicationNonVisualDrawingProperties?
                .GetFirstChild<PlaceholderShape>();
            return placeholder?.Type?.Value == PlaceholderValues.Body;
        }

        private static string ReadParagraphText(A.Paragraph paragraph) {
            StringBuilder builder = new();
            foreach (DocumentFormat.OpenXml.OpenXmlElement child in paragraph.ChildElements) {
                switch (child) {
                    case A.Run run:
                        builder.Append(run.Text?.Text ?? string.Empty);
                        break;
                    case A.Break:
                        builder.AppendLine();
                        break;
                    case A.Field field:
                        builder.Append(field.Text?.Text ?? string.Empty);
                        break;
                }
            }

            return builder.ToString();
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
