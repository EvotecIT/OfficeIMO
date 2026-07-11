using System.Text;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>Whether the slide already contains a speaker-notes part.</summary>
        public bool HasSpeakerNotes => _slidePart.NotesSlidePart?.NotesSlide != null;

        /// <summary>Reads existing speaker-note text without creating a notes part as a side effect.</summary>
        public string GetSpeakerNotesText() {
            NotesSlide? notesSlide = _slidePart.NotesSlidePart?.NotesSlide;
            if (notesSlide == null) return string.Empty;
            List<string> blocks = notesSlide.CommonSlideData?.ShapeTree?
                .Elements<Shape>()
                .Select(ReadNotesShapeText)
                .Where(text => !string.IsNullOrWhiteSpace(text))
                .ToList() ?? new List<string>();
            return string.Join("\n\n", blocks);
        }

        private static string ReadNotesShapeText(Shape shape) {
            List<string> paragraphs = shape.TextBody?
                .Elements<A.Paragraph>()
                .Select(ReadNotesParagraphText)
                .Where(text => !string.IsNullOrWhiteSpace(text))
                .ToList() ?? new List<string>();
            return string.Join("\n", paragraphs);
        }

        private static string ReadNotesParagraphText(A.Paragraph paragraph) {
            var builder = new StringBuilder();
            foreach (OpenXmlElement child in paragraph.ChildElements) {
                if (child is A.Run run) builder.Append(run.Text?.Text ?? string.Empty);
                else if (child is A.Break) builder.AppendLine();
                else if (child is A.Field field) builder.Append(field.Text?.Text ?? string.Empty);
            }
            return builder.ToString().Trim();
        }
    }
}
