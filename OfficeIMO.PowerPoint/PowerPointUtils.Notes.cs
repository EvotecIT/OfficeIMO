using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointUtils {

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
                    notesThemePart.Theme = (D.Theme)fallback.Theme.CloneNode(true);
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

            return new NotesMaster(
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
