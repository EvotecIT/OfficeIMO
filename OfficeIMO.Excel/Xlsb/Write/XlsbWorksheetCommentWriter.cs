using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Model;
using OfficeIMO.Excel.Xlsb.Projection;
using OfficeIMO.Excel.Xlsb.Read;
using Reader = OfficeIMO.Excel.Xlsb.Read.XlsbWorksheetCommentReader;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>
    /// The legacy comments of a worksheet as the native writers accept them: one WorksheetCommentsPart, the VmlDrawingPart
    /// its LegacyDrawing element names, and nothing else related to the worksheet.
    /// </summary>
    internal sealed class XlsbWorksheetCommentModel {
        private XlsbWorksheetCommentModel(WorksheetCommentsPart commentsPart, Comments comments, VmlDrawingPart? vmlPart) {
            CommentsPart = commentsPart;
            Comments = comments;
            VmlPart = vmlPart;
        }

        private const string RelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        internal WorksheetCommentsPart CommentsPart { get; }

        internal Comments Comments { get; }

        internal VmlDrawingPart? VmlPart { get; }

        internal byte[] ReadVml() {
            if (VmlPart == null) return Array.Empty<byte>();
            using Stream input = VmlPart.GetStream(FileMode.Open, FileAccess.Read);
            using var output = new MemoryStream();
            input.CopyTo(output);
            return output.ToArray();
        }

        /// <summary>Reads and validates the worksheet's comments; null when the worksheet has none.</summary>
        internal static XlsbWorksheetCommentModel? Read(ExcelSheet sheet, string operation) {
            WorksheetPart worksheetPart = sheet.WorksheetPart;
            Worksheet worksheet = worksheetPart.Worksheet
                ?? throw new InvalidDataException($"Worksheet '{sheet.Name}' has no worksheet root.");
            LegacyDrawing[] legacy = worksheet.Elements<LegacyDrawing>().ToArray();
            WorksheetCommentsPart? commentsPart = worksheetPart.WorksheetCommentsPart;
            if (commentsPart == null) {
                if (legacy.Length != 0) throw new NotSupportedException($"{operation} supports a legacy VML drawing only for the comments of worksheet '{sheet.Name}'.");
                return null;
            }
            if (legacy.Length > 1) throw new NotSupportedException($"{operation} requires at most one legacy drawing on worksheet '{sheet.Name}'.");
            VmlDrawingPart? vmlPart = null;
            if (legacy.Length == 1) {
                if (legacy[0].HasChildren || XlsbOpenXmlAttributeValidator.FindUnsupported(legacy[0], Array.Empty<string>(),
                        attribute => attribute.LocalName == "id" && attribute.NamespaceUri == RelationshipNamespace) is OpenXmlAttribute
                    || legacy[0].Id?.Value is not string id || worksheetPart.GetPartById(id) is not VmlDrawingPart vml) {
                    throw new NotSupportedException($"{operation} found an unsupported legacy drawing on worksheet '{sheet.Name}'.");
                }
                vmlPart = vml;
            }
            if (commentsPart.Parts.Any() || commentsPart.ExternalRelationships.Any()) {
                throw new NotSupportedException($"{operation} found relationships from the comments part of worksheet '{sheet.Name}'.");
            }
            Comments comments = commentsPart.Comments
                ?? throw new NotSupportedException($"{operation} found an empty comments part on worksheet '{sheet.Name}'.");
            Validate(comments, sheet.Name, operation);
            return new XlsbWorksheetCommentModel(commentsPart, comments, vmlPart);
        }

        private static void Validate(Comments comments, string sheetName, string operation) {
            string message = $"{operation} found legacy comment content it cannot encode on worksheet '{sheetName}'";
            if (comments.GetAttributes().Count != 0 || comments.ChildElements.Any(child => child is not Authors and not CommentList)
                || comments.Elements<Authors>().Count() > 1 || comments.Elements<CommentList>().Count() > 1) {
                throw new NotSupportedException(message + ".");
            }
            string[] authors = (comments.Authors?.Elements<Author>() ?? Enumerable.Empty<Author>()).Select(author => author.Text ?? string.Empty).ToArray();
            if (comments.Authors != null && (comments.Authors.ChildElements.Any(child => child is not Author) || comments.Authors.GetAttributes().Count != 0)) {
                throw new NotSupportedException(message + ".");
            }
            foreach (string author in authors) {
                if (author.Length < 1 || author.Length > Reader.MaximumAuthorLength) {
                    throw new NotSupportedException($"{operation} requires comment author names of 1 to {Reader.MaximumAuthorLength} characters on worksheet '{sheetName}'; '{author}' does not fit.");
                }
            }
            var cells = new HashSet<(int Row, int Column)>();
            foreach (OpenXmlElement child in comments.CommentList?.ChildElements ?? Enumerable.Empty<OpenXmlElement>()) {
                if (child is not Comment comment || comment.ChildElements.Any(item => item is not CommentText) || comment.Elements<CommentText>().Count() != 1
                    || XlsbOpenXmlAttributeValidator.FindUnsupported(comment, new[] { "ref", "authorId", "guid" }) is OpenXmlAttribute) {
                    throw new NotSupportedException(message + ".");
                }
                if (!TryParseCell(comment.Reference?.Value, out int row, out int column) || !cells.Add((row, column))) {
                    throw new NotSupportedException($"{message}: comment reference '{comment.Reference?.Value}' is not a single, unique cell.");
                }
                if ((comment.AuthorId?.Value ?? 0U) >= (uint)authors.Length) {
                    throw new NotSupportedException($"{message}: the comment on {comment.Reference?.Value} has no author.");
                }
                CommentText text = comment.Elements<CommentText>().Single();
                if (text.GetAttributes().Count != 0 || text.ChildElements.Any(item => item is not Run and not Text)
                    || (text.Elements<Text>().Any() && text.ChildElements.Count != 1)
                    || text.Elements<Run>().Any(run => run.ChildElements.Any(item => item is not RunProperties and not Text))) {
                    throw new NotSupportedException($"{message}: the comment on {comment.Reference?.Value} has text structure other than plain or formatted runs.");
                }
                if (GetText(text).Length > Reader.MaximumTextLength) {
                    throw new NotSupportedException($"{message}: the comment on {comment.Reference?.Value} is longer than {Reader.MaximumTextLength} characters.");
                }
            }
        }

        internal static string GetText(CommentText text) {
            Text? plain = text.GetFirstChild<Text>();
            if (plain != null) return Utilities.ExcelSanitizer.DecodeXString(plain.Text ?? string.Empty);
            return string.Concat(text.Elements<Run>().Select(run => Utilities.ExcelSanitizer.DecodeXString(run.Text?.Text ?? string.Empty)));
        }

        internal static bool TryParseCell(string? reference, out int row, out int column) {
            row = 0;
            column = 0;
            return !string.IsNullOrWhiteSpace(reference)
                && XlsbPayloadWriter.TryParseRange(reference!, out row, out column, out int lastRow, out int lastColumn)
                && row == lastRow && column == lastColumn;
        }
    }

    /// <summary>Encodes legacy comments as a BIFF12 Comments part (MS-XLSB 2.1.7.8) with its VML drawing and BrtLegacyDrawing.</summary>
    internal static class XlsbWorksheetCommentWriter {
        private const int BrtLegacyDrawing = 551;

        internal static void Append(XlsbWorksheetFeatureRecords records, XlsbWorksheetFeatureContext context) {
            XlsbWorksheetCommentModel? model = XlsbWorksheetCommentModel.Read(context.Sheet, "Native XLSB generation");
            if (model == null) return;
            if (context.Parts == null || context.RichText == null) {
                throw new NotSupportedException($"Native XLSB rewriting cannot add the comments part of new worksheet '{context.Sheet.Name}'; save as .xlsx to retain the comments.");
            }
            if (model.VmlPart == null) {
                throw new NotSupportedException($"Native XLSB generation requires the VML drawing that displays the comments on worksheet '{context.Sheet.Name}'.");
            }
            if (model.VmlPart.Parts.Any() || model.VmlPart.ExternalRelationships.Any() || model.VmlPart.HyperlinkRelationships.Any()) {
                throw new NotSupportedException($"Native XLSB generation cannot carry the related parts of the comment drawing on worksheet '{context.Sheet.Name}'.");
            }

            byte[] comments = CreatePart(model.Comments, context.RichText, native: null, nativeStyles: null, context.Sheet.Name);
            string commentsPartName = context.Parts.Names.Next("xl/comments", ".bin");
            context.Parts.AddPart(new XlsbGeneratedPart(commentsPartName, XlsbCommentRelationships.CommentsContentType, comments));
            context.Parts.AddRelationship(XlsbCommentRelationships.Comments, commentsPartName);
            string vmlPartName = context.Parts.Names.Next("xl/drawings/vmlDrawing", ".vml");
            context.Parts.AddPart(new XlsbGeneratedPart(vmlPartName, XlsbCommentRelationships.VmlDrawingContentType, model.ReadVml()));
            string vmlId = context.Parts.AddRelationship(XlsbCommentRelationships.VmlDrawing, vmlPartName);
            using var payload = new MemoryStream();
            XlsbPayloadWriter.WriteWideString(payload, vmlId);
            records.Drawings.Add(new XlsbGeneratedRecord(BrtLegacyDrawing, payload.ToArray()));
        }

        /// <summary>
        /// Encodes COMMENTS for <paramref name="comments"/>. A comment equal to the projection of the native comment on the
        /// same cell keeps that comment's records, with only its author index updated.
        /// </summary>
        internal static byte[] CreatePart(Comments comments, XlsbRichTextFontTable fonts, XlsbComments? native, XlsbStylesheet? nativeStyles,
            string sheetName) {
            string[] authors = (comments.Authors?.Elements<Author>() ?? Enumerable.Empty<Author>()).Select(author => author.Text ?? string.Empty).ToArray();
            var nativeByCell = native?.Items.ToDictionary(item => (item.Row, item.Column));
            using var output = new MemoryStream();
            XlsbRecordWriter.Write(output, Reader.BrtBeginComments);
            XlsbRecordWriter.Write(output, Reader.BrtBeginCommentAuthors);
            foreach (string author in authors) {
                using var payload = new MemoryStream();
                XlsbPayloadWriter.WriteWideString(payload, author);
                XlsbRecordWriter.Write(output, Reader.BrtCommentAuthor, payload.ToArray());
            }
            XlsbRecordWriter.Write(output, Reader.BrtEndCommentAuthors);
            XlsbRecordWriter.Write(output, Reader.BrtBeginCommentList);
            foreach (Comment comment in comments.CommentList?.Elements<Comment>() ?? Enumerable.Empty<Comment>()) {
                XlsbWorksheetCommentModel.TryParseCell(comment.Reference?.Value, out int row, out int column);
                int authorIndex = checked((int)(comment.AuthorId?.Value ?? 0U));
                CommentText text = comment.Elements<CommentText>().Single();
                if (nativeByCell != null && nativeByCell.TryGetValue((row, column), out XlsbComment? source)
                    && comment.Guid == null
                    && string.Equals(native!.Authors[source.AuthorIndex], authors[authorIndex], StringComparison.Ordinal)
                    && string.Equals(XlsbWorksheetCommentProjector.Create(source, nativeStyles).CommentText?.OuterXml, text.OuterXml, StringComparison.Ordinal)) {
                    WriteNativeComment(output, source.Records, authorIndex);
                    continue;
                }
                WriteComment(output, comment, row, column, authorIndex, text, fonts, sheetName);
            }
            XlsbRecordWriter.Write(output, Reader.BrtEndCommentList);
            if (native != null && native.TrailingRecords.Length != 0) output.Write(native.TrailingRecords, 0, native.TrailingRecords.Length);
            XlsbRecordWriter.Write(output, Reader.BrtEndComments);
            return output.ToArray();
        }

        private static void WriteComment(Stream output, Comment comment, int row, int column, int authorIndex, CommentText text,
            XlsbRichTextFontTable fonts, string sheetName) {
            using (var begin = new MemoryStream(36)) {
                XlsbPayloadWriter.WriteInt32(begin, authorIndex);
                XlsbPayloadWriter.WriteRfX(begin, row, column, row, column);
                // The GUID (MS-DTYP 2.3.4, the little-endian layout Guid.ToByteArray produces) matters only to shared workbooks.
                byte[] guid = new byte[16];
                if (comment.Guid?.Value is string guidText) {
                    if (!System.Guid.TryParse(guidText, out Guid parsed)) {
                        throw new NotSupportedException($"Native XLSB generation found the invalid comment GUID '{guidText}' on worksheet '{sheetName}'.");
                    }
                    guid = parsed.ToByteArray();
                }
                begin.Write(guid, 0, guid.Length);
                XlsbRecordWriter.Write(output, Reader.BrtBeginComment, begin.ToArray());
            }

            // BrtCommentText: RichStr with fRichStr = 1 and fExtStr = 0; dwSizeStrRun is 0 when no run has its own font.
            string value = XlsbWorksheetCommentModel.GetText(text);
            XlsbTextRun[]? runs = fonts.CreateRuns(text.Elements<Run>().ToArray(), 0, value, $"the comment text on {sheetName}!{comment.Reference?.Value}");
            using (var payload = new MemoryStream()) {
                payload.WriteByte(0x01);
                XlsbPayloadWriter.WriteWideString(payload, value);
                XlsbPayloadWriter.WriteUInt32(payload, (uint)(runs?.Length ?? 0));
                foreach (XlsbTextRun run in runs ?? Array.Empty<XlsbTextRun>()) {
                    XlsbPayloadWriter.WriteUInt16(payload, run.Start);
                    XlsbPayloadWriter.WriteUInt16(payload, run.Font);
                }
                XlsbRecordWriter.Write(output, Reader.BrtCommentText, payload.ToArray());
            }
            XlsbRecordWriter.Write(output, Reader.BrtEndComment);
        }

        private static void WriteNativeComment(Stream output, byte[] records, int authorIndex) {
            using var input = new MemoryStream(records, writable: false);
            foreach (XlsbRecord record in XlsbRecordReader.ReadAll(input)) {
                byte[] payload = record.Data;
                if (record.Type == Reader.BrtBeginComment) {
                    payload = (byte[])payload.Clone();
                    for (int index = 0; index < 4; index++) payload[index] = (byte)(authorIndex >> (8 * index));
                }
                XlsbRecordWriter.Write(output, record.Type, payload);
            }
        }
    }
}
