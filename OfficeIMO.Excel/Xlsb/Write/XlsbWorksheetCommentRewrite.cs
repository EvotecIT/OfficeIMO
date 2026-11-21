using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>
    /// Carries a loaded worksheet's legacy comments through a native rewrite. Unchanged comment and VML parts are copied;
    /// edited ones are re-encoded into the same package parts, so the worksheet record, relationships and content types
    /// stay valid. Adding comments to a worksheet without them, or removing all of them, is refused.
    /// </summary>
    internal static class XlsbWorksheetCommentRewrite {
        internal static void Apply(ExcelDocument document, ExcelSheet sheet, XlsbSheet source, XlsbWorkbook sourceWorkbook,
            IDictionary<string, byte[]> replacements) {
            XlsbWorksheetCommentModel? model = XlsbWorksheetCommentModel.Read(sheet, "Native XLSB rewriting");
            XlsbComments? native = source.Comments;
            if (model == null && native == null) return;
            if (model == null) {
                throw new NotSupportedException($"Native XLSB rewriting cannot remove every comment from worksheet '{sheet.Name}'; save as .xlsx to retain that change.");
            }
            if (native == null) {
                throw new NotSupportedException($"Native XLSB rewriting cannot add a comments part to worksheet '{sheet.Name}', which had no comments; save as .xlsx to retain the comments.");
            }

            XlsbProjectionPreservationGuard guard = document.NativeXlsbPreservation;
            bool commentsChanged = !guard.MatchesCaptured(model.CommentsPart);
            bool vmlChanged = model.VmlPart == null ? native.VmlPartName != null : !guard.MatchesCaptured(model.VmlPart);
            if (!commentsChanged && !vmlChanged) return;
            if (model.VmlPart == null || native.VmlPartName == null) {
                throw new NotSupportedException($"Native XLSB rewriting cannot add or remove the drawing that displays the comments on worksheet '{sheet.Name}'; save as .xlsx to retain that change.");
            }

            if (commentsChanged) {
                Stylesheet stylesheet = document.WorkbookPartRoot.WorkbookStylesPart?.Stylesheet
                    ?? throw new NotSupportedException($"Native XLSB rewriting cannot encode the comments on worksheet '{sheet.Name}' without a style table.");
                var fonts = new XlsbRichTextFontTable(stylesheet, new XlsbCellTextRuns(null));
                byte[] part = XlsbWorksheetCommentWriter.CreatePart(model.Comments, fonts, native, sourceWorkbook.Stylesheet, sheet.Name);
                if (fonts.AddedFonts.Count != 0) {
                    throw new NotSupportedException($"Native XLSB rewriting cannot add the fonts that the edited comment formatting on worksheet '{sheet.Name}' needs; save as .xlsx to retain that formatting.");
                }
                replacements[native.PartName] = part;
            }
            if (vmlChanged) {
                if (model.VmlPart.Parts.Any() || model.VmlPart.ExternalRelationships.Any() || model.VmlPart.HyperlinkRelationships.Any()) {
                    throw new NotSupportedException($"Native XLSB rewriting cannot change a comment drawing with related parts on worksheet '{sheet.Name}'.");
                }
                replacements[native.VmlPartName] = model.ReadVml();
            }
        }
    }
}
