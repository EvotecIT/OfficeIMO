using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointUtils {

        /// <summary>
        /// The notes slide part of a slide the SDK holds, created as the notes model first uses it: an empty body
        /// placeholder, with relationships back to the slide and to the presentation's notes master, which is created
        /// too when the presentation has none. Takes and returns SDK parts as <see cref="object"/>.
        /// </summary>
        internal static object EnsureNotesSlidePart(object slide) {
            SlidePart slidePart = (SlidePart)slide;
            NotesSlidePart? notesPart = slidePart.NotesSlidePart;
            if (notesPart == null) {
                var used = new HashSet<string>(
                    slidePart.Parts.Select(part => part.RelationshipId)
                        .Concat(slidePart.DataPartReferenceRelationships.Select(relationship => relationship.Id))
                        .Concat(slidePart.ExternalRelationships.Select(relationship => relationship.Id))
                        .Concat(slidePart.HyperlinkRelationships.Select(relationship => relationship.Id))
                        .Where(id => !string.IsNullOrEmpty(id)),
                    StringComparer.Ordinal);
                int number = 1;
                string relationshipId;
                do {
                    relationshipId = "rId" + number++;
                } while (!used.Add(relationshipId));
                notesPart = slidePart.AddNewPart<NotesSlidePart>(relationshipId);
            }

            if (!notesPart.Parts.Any(pair => ReferenceEquals(pair.OpenXmlPart, slidePart))) notesPart.AddPart(slidePart);
            if (slidePart.GetParentPresentationPart() is { } presentationPart) {
                NotesMasterPart notesMasterPart = EnsureNotesMasterPart(presentationPart);
                if (!notesPart.Parts.Any(pair => ReferenceEquals(pair.OpenXmlPart, notesMasterPart))) notesPart.AddPart(notesMasterPart);
            }

            if (notesPart.NotesSlide == null) {
                ShapeTree shapeTree = new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1U, Name = "Notes Group Shape" },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    CreateDefaultGroupShapeProperties());
                shapeTree.Append(new P.Shape(
                    new P.NonVisualShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 2U, Name = "Notes Placeholder 1" },
                        new P.NonVisualShapeDrawingProperties(new D.ShapeLocks { NoGrouping = true }),
                        new ApplicationNonVisualDrawingProperties(new PlaceholderShape { Type = PlaceholderValues.Body, Index = 1U })),
                    new P.ShapeProperties(),
                    new P.TextBody(
                        new D.BodyProperties(),
                        new D.ListStyle(),
                        new D.Paragraph(
                            new D.Run(new D.RunProperties { Language = "en-US" }, new D.Text()),
                            new D.EndParagraphRunProperties { Language = "en-US" }))));
                notesPart.NotesSlide = new NotesSlide(new CommonSlideData(shapeTree), new ColorMapOverride(new D.MasterColorMapping()));
            }

            return notesPart;
        }

        /// <summary>Writes the notes slide of a slide the SDK holds, when it has one loaded.</summary>
        internal static void SaveNotesSlide(object slide) => ((SlidePart)slide).NotesSlidePart?.NotesSlide?.Save();

        internal static NotesMasterPart EnsureNotesMasterPart(PresentationPart presentationPart) {
            NotesMasterPart notesMasterPart = presentationPart.NotesMasterPart ?? presentationPart.AddNewPart<NotesMasterPart>();

            if (notesMasterPart.NotesMaster == null) {
                notesMasterPart.NotesMaster = CreateDefaultNotesMaster();
            }

            if (notesMasterPart.ThemePart == null) {
                ThemePart notesThemePart = notesMasterPart.AddNewPart<ThemePart>();
                if (presentationPart.ThemePart?.Theme != null) {
                    notesThemePart.Theme = (D.Theme)presentationPart.ThemePart.Theme.CloneNode(true);
                } else {
                    ThemePart fallback = CreateTheme(presentationPart);
                    D.Theme fallbackTheme = fallback.Theme ?? throw new InvalidOperationException("Theme part did not contain a theme.");
                    notesThemePart.Theme = (D.Theme)fallbackTheme.CloneNode(true);
                }
            }

            Presentation presentation = presentationPart.Presentation ??= new Presentation();
            NotesMasterIdList notesMasterIdList = presentation.NotesMasterIdList ??= new NotesMasterIdList();

            string relationshipId = presentationPart.GetIdOfPart(notesMasterPart);
            bool hasEntry = notesMasterIdList
                .Elements<NotesMasterId>()
                .Any(existing => GetRelationshipId(existing) == relationshipId);
            if (!hasEntry) {
                NotesMasterId notesMasterId = new NotesMasterId();
                SetRelationshipId(notesMasterId, relationshipId);
                notesMasterIdList.AppendChild(notesMasterId);
            }

            return notesMasterPart;
        }

        private static NotesMaster CreateDefaultNotesMaster() {
            ShapeTree shapeTree = new ShapeTree();
            shapeTree.Append(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 1U, Name = "Notes Group Shape" },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                CreateDefaultGroupShapeProperties());

            shapeTree.Append(
                CreatePlaceholderShape(2U, "Notes Placeholder", PlaceholderValues.Body, 1U, includeEndParagraph: true),
                CreatePlaceholderShape(3U, "Slide Image Placeholder", PlaceholderValues.SlideImage, 2U, includeEndParagraph: false),
                CreatePlaceholderShape(4U, "Date Placeholder", PlaceholderValues.DateAndTime, 3U, includeEndParagraph: true),
                CreatePlaceholderShape(5U, "Slide Number Placeholder", PlaceholderValues.SlideNumber, 4U, includeEndParagraph: true),
                CreatePlaceholderShape(6U, "Footer Placeholder", PlaceholderValues.Footer, 5U, includeEndParagraph: true));

            Background background = new Background(new BackgroundProperties(new D.NoFill()));

            var master = new NotesMaster(
                new CommonSlideData(background, shapeTree),
                new P.ColorMap {
                    Background1 = D.ColorSchemeIndexValues.Light1,
                    Text1 = D.ColorSchemeIndexValues.Dark1,
                    Background2 = D.ColorSchemeIndexValues.Light2,
                    Text2 = D.ColorSchemeIndexValues.Dark2,
                    Accent1 = D.ColorSchemeIndexValues.Accent1,
                    Accent2 = D.ColorSchemeIndexValues.Accent2,
                    Accent3 = D.ColorSchemeIndexValues.Accent3,
                    Accent4 = D.ColorSchemeIndexValues.Accent4,
                    Accent5 = D.ColorSchemeIndexValues.Accent5,
                    Accent6 = D.ColorSchemeIndexValues.Accent6,
                    Hyperlink = D.ColorSchemeIndexValues.Hyperlink,
                    FollowedHyperlink = D.ColorSchemeIndexValues.FollowedHyperlink
                },
                new NotesStyle(
                    new D.Level1ParagraphProperties(new D.DefaultRunProperties { Language = "en-US" }),
                    new D.Level2ParagraphProperties(new D.DefaultRunProperties { Language = "en-US" }),
                    new D.Level3ParagraphProperties(new D.DefaultRunProperties { Language = "en-US" }),
                    new D.Level4ParagraphProperties(new D.DefaultRunProperties { Language = "en-US" }),
                    new D.Level5ParagraphProperties(new D.DefaultRunProperties { Language = "en-US" }),
                    new D.Level6ParagraphProperties(new D.DefaultRunProperties { Language = "en-US" }),
                    new D.Level7ParagraphProperties(new D.DefaultRunProperties { Language = "en-US" }),
                    new D.Level8ParagraphProperties(new D.DefaultRunProperties { Language = "en-US" }),
                    new D.Level9ParagraphProperties(new D.DefaultRunProperties { Language = "en-US" }))
            );
            // Loading the serialized master promotes this namespace to the root; emit it now so an unchanged notes part stays byte-stable.
            master.AddNamespaceDeclaration("a", DrawingNamespace);
            return master;
        }

        private static P.Shape CreatePlaceholderShape(uint id, string name, PlaceholderValues type, uint index, bool includeEndParagraph) {
            P.Shape shape = new P.Shape(
                new P.NonVisualShapeProperties(
                    new P.NonVisualDrawingProperties { Id = id, Name = name },
                    new P.NonVisualShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape { Type = type, Index = index })),
                new P.ShapeProperties(),
                new P.TextBody(
                    new D.BodyProperties(),
                    new D.ListStyle()));

            D.Paragraph paragraph = new D.Paragraph();
            if (includeEndParagraph) {
                paragraph.Append(new D.EndParagraphRunProperties { Language = "en-US" });
            }

            shape.TextBody!.Append(paragraph);
            return shape;
        }

        private static string? GetRelationshipId(NotesMasterId notesMasterId) {
            OpenXmlAttribute attribute = notesMasterId.GetAttribute("id", RelationshipNamespace);
            return string.IsNullOrEmpty(attribute.Value) ? null : attribute.Value;
        }

        private static void SetRelationshipId(NotesMasterId notesMasterId, string relationshipId) {
            notesMasterId.SetAttribute(new OpenXmlAttribute("r", "id", RelationshipNamespace, relationshipId));
        }

    }
}
