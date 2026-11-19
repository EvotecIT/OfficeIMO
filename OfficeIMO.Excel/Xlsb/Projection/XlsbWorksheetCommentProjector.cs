using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Model;
using OfficeIMO.Excel.Xlsb.Read;

namespace OfficeIMO.Excel.Xlsb.Projection {
    /// <summary>
    /// Projects a worksheet's native legacy comments into a WorksheetCommentsPart, its VML drawing into a VmlDrawingPart
    /// and BrtLegacyDrawing into the LegacyDrawing element, under the native relationship ids.
    /// </summary>
    internal static class XlsbWorksheetCommentProjector {
        internal static void Apply(ExcelSheet sheet, XlsbSheet source, XlsbStylesheet? stylesheet, XlsbRelatedPartProjector relatedParts) {
            XlsbComments? comments = source.Comments;
            if (comments == null) return;
            WorksheetPart worksheetPart = sheet.WorksheetPart;
            WorksheetCommentsPart commentsPart = worksheetPart.AddNewPart<WorksheetCommentsPart>(comments.RelationshipId);
            commentsPart.Comments = Create(comments, stylesheet);
            commentsPart.Comments.Save();
            if (comments.VmlPartName == null || comments.VmlRelationshipId == null) return;
            relatedParts.ProjectWorksheetPart(worksheetPart, comments.VmlPartName, XlsbCommentRelationships.VmlDrawing, comments.VmlRelationshipId);
            Worksheet worksheet = worksheetPart.Worksheet
                ?? throw new InvalidDataException($"Worksheet '{sheet.Name}' has no worksheet root.");
            worksheet.Append(new LegacyDrawing { Id = comments.VmlRelationshipId });
        }

        internal static Comments Create(XlsbComments source, XlsbStylesheet? stylesheet) {
            var authors = new Authors();
            foreach (string author in source.Authors) authors.Append(new Author(author));
            var list = new CommentList();
            foreach (XlsbComment item in source.Items) list.Append(Create(item, stylesheet));
            return new Comments(authors, list);
        }

        internal static Comment Create(XlsbComment source, XlsbStylesheet? stylesheet) {
            var comment = new Comment {
                Reference = ExcelA1.CellReference(source.Row, source.Column),
                AuthorId = (uint)source.AuthorIndex
            };
            comment.Append(CreateText(source.Text, source.Runs, stylesheet));
            return comment;
        }

        // A StrRun covers its text up to the next run; text before the first run, or all text without runs, has no
        // run properties.
        private static CommentText CreateText(string text, XlsbTextRun[]? runs, XlsbStylesheet? stylesheet) {
            var result = new CommentText();
            if (runs == null || runs.Length == 0) {
                result.Append(CreateRun(text, null));
                return result;
            }
            if (runs[0].Start > 0) result.Append(CreateRun(text.Substring(0, runs[0].Start), null));
            for (int index = 0; index < runs.Length; index++) {
                int start = runs[index].Start;
                int end = index + 1 < runs.Length ? runs[index + 1].Start : text.Length;
                ushort font = runs[index].Font;
                XlsbFont? fontSource = stylesheet != null && font < stylesheet.Fonts.Count ? stylesheet.Fonts[font] : null;
                result.Append(CreateRun(text.Substring(start, end - start),
                    fontSource == null ? null : XlsbStylesheetProjector.CreateRunProperties(fontSource)));
            }
            return result;
        }

        private static Run CreateRun(string text, RunProperties? properties) {
            var run = new Run();
            if (properties != null) run.Append(properties);
            run.Append(new Text(Utilities.ExcelSanitizer.EncodeXString(text)) { Space = SpaceProcessingModeValues.Preserve });
            return run;
        }
    }
}
