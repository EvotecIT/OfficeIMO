using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Threaded = DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments;

namespace OfficeIMO.Excel.Utilities {
    internal static class ExcelWorksheetCommentResolver {
        internal static Dictionary<string, ExcelCommentSnapshot> BuildLegacyCommentMap(WorksheetPart worksheetPart) {
            if (worksheetPart == null) {
                throw new ArgumentNullException(nameof(worksheetPart));
            }

            var commentsPart = worksheetPart.WorksheetCommentsPart;
            var comments = commentsPart?.Comments;
            if (comments?.CommentList == null) {
                return new Dictionary<string, ExcelCommentSnapshot>(StringComparer.OrdinalIgnoreCase);
            }

            var authorNames = comments.Authors?
                .Elements<Author>()
                .Select(author => author.Text ?? string.Empty)
                .ToList()
                ?? new List<string>();

            var map = new Dictionary<string, ExcelCommentSnapshot>(StringComparer.OrdinalIgnoreCase);
            WorkbookPart? workbookPart = worksheetPart.GetParentParts().OfType<WorkbookPart>().FirstOrDefault();
            foreach (var comment in comments.CommentList.Elements<Comment>()) {
                string? reference = comment.Reference?.Value;
                if (string.IsNullOrWhiteSpace(reference)) {
                    continue;
                }

                string? author = null;
                uint? authorId = comment.AuthorId?.Value;
                if (authorId.HasValue && authorId.Value < authorNames.Count) {
                    author = authorNames[checked((int)authorId.Value)];
                }

                map[reference!] = new ExcelCommentSnapshot {
                    Author = string.IsNullOrWhiteSpace(author) ? null : author,
                    Text = ExtractCommentText(comment.CommentText),
                    RichTextRuns = ExtractCommentRuns(comment.CommentText, workbookPart),
                };
            }

            return map;
        }

        internal static Dictionary<string, string> BuildThreadedCommentPersonMap(WorkbookPart workbookPart) {
            if (workbookPart == null) {
                throw new ArgumentNullException(nameof(workbookPart));
            }

            var people = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var personPart in workbookPart.WorkbookPersonParts) {
                var personList = personPart.PersonList;
                if (personList == null) {
                    continue;
                }

                foreach (var person in personList.Elements<Threaded.Person>()) {
                    string? id = person.Id?.Value;
                    if (string.IsNullOrWhiteSpace(id)) {
                        continue;
                    }

                    string? displayName = person.DisplayName?.Value;
                    if (!string.IsNullOrWhiteSpace(displayName)) {
                        people[id!] = displayName!;
                    }
                }
            }

            return people;
        }

        internal static Dictionary<string, List<ExcelThreadedCommentSnapshot>> BuildThreadedCommentMap(
            WorksheetPart worksheetPart,
            IReadOnlyDictionary<string, string> people,
            string? sheetName = null) {
            if (worksheetPart == null) {
                throw new ArgumentNullException(nameof(worksheetPart));
            }

            if (people == null) {
                throw new ArgumentNullException(nameof(people));
            }

            var map = new Dictionary<string, List<ExcelThreadedCommentSnapshot>>(StringComparer.OrdinalIgnoreCase);
            foreach (var commentsPart in worksheetPart.WorksheetThreadedCommentsParts) {
                var threadedComments = commentsPart.ThreadedComments;
                if (threadedComments == null) {
                    continue;
                }

                foreach (var comment in threadedComments.Elements<Threaded.ThreadedComment>()) {
                    string? reference = comment.Ref?.Value;
                    if (string.IsNullOrWhiteSpace(reference)) {
                        continue;
                    }

                    string? personId = comment.PersonId?.Value;
                    string? author = null;
                    if (!string.IsNullOrWhiteSpace(personId) && people.TryGetValue(personId!, out string? displayName)) {
                        author = displayName;
                    }

                    var snapshot = new ExcelThreadedCommentSnapshot {
                        SheetName = sheetName ?? string.Empty,
                        CellReference = reference!,
                        Id = NullIfWhiteSpace(comment.Id?.Value),
                        ParentId = NullIfWhiteSpace(comment.ParentId?.Value),
                        PersonId = NullIfWhiteSpace(personId),
                        Author = author,
                        Text = NormalizeMultilineText(comment.ThreadedCommentText?.InnerText ?? string.Empty),
                        Date = comment.DT?.Value,
                        Done = comment.Done?.Value == true,
                    };

                    if (!map.TryGetValue(reference!, out List<ExcelThreadedCommentSnapshot>? comments)) {
                        comments = new List<ExcelThreadedCommentSnapshot>();
                        map[reference!] = comments;
                    }

                    comments.Add(snapshot);
                }
            }

            return map;
        }

        private static string ExtractCommentText(CommentText? commentText) {
            if (commentText == null) {
                return string.Empty;
            }

            var builder = new System.Text.StringBuilder();
            foreach (var element in commentText.Descendants<OpenXmlElement>()) {
                if (element is Text text) {
                    builder.Append(text.Text);
                } else if (IsLineBreak(element)) {
                    builder.Append('\n');
                }
            }

            return NormalizeMultilineText(builder.ToString());
        }

        private static IReadOnlyList<ExcelRichTextRun> ExtractCommentRuns(CommentText? commentText, WorkbookPart? workbookPart) {
            if (commentText == null) {
                return Array.Empty<ExcelRichTextRun>();
            }

            var runs = new List<ExcelRichTextRun>();
            foreach (OpenXmlElement element in commentText.ChildElements) {
                if (element is Run run) {
                    AppendRunElements(runs, run, run.RunProperties, workbookPart);
                } else if (element is Text text) {
                    AppendCommentRun(runs, text.Text ?? string.Empty, null, workbookPart);
                } else if (IsLineBreak(element)) {
                    AppendCommentRun(runs, "\n", null, workbookPart);
                }
            }

            if (runs.Count == 0) {
                string plainText = ExtractCommentText(commentText);
                if (!string.IsNullOrEmpty(plainText)) {
                    runs.Add(new ExcelRichTextRun(plainText));
                }
            }

            string extracted = NormalizeMultilineText(string.Concat(runs.Select(run => run.Text)));
            string expected = ExtractCommentText(commentText);
            if (!string.Equals(extracted, expected, StringComparison.Ordinal)) {
                return string.IsNullOrEmpty(expected)
                    ? Array.Empty<ExcelRichTextRun>()
                    : new[] { new ExcelRichTextRun(expected) };
            }

            return runs.Count == 0 ? Array.Empty<ExcelRichTextRun>() : runs.AsReadOnly();
        }

        private static void AppendRunElements(List<ExcelRichTextRun> runs, Run run, RunProperties? properties, WorkbookPart? workbookPart) {
            foreach (OpenXmlElement child in run.ChildElements) {
                if (child is Text text) {
                    AppendCommentRun(runs, text.Text ?? string.Empty, properties, workbookPart);
                } else if (IsLineBreak(child)) {
                    AppendCommentRun(runs, "\n", properties, workbookPart);
                }
            }
        }

        private static void AppendCommentRun(List<ExcelRichTextRun> runs, string text, RunProperties? properties, WorkbookPart? workbookPart) {
            runs.Add(new ExcelRichTextRun(text) {
                Bold = properties?.GetFirstChild<Bold>() != null,
                Italic = properties?.GetFirstChild<Italic>() != null,
                Underline = properties?.GetFirstChild<Underline>() != null,
                Strikethrough = properties?.GetFirstChild<Strike>() != null,
                UnderlineStyle = ExcelRichTextRun.GetUnderlineStyle(properties),
                FontColor = ExcelThemeColorResolver.Resolve(properties?.GetFirstChild<Color>(), workbookPart),
                FontName = properties?.GetFirstChild<RunFont>()?.Val?.Value,
                FontSize = properties?.GetFirstChild<FontSize>()?.Val?.Value,
                VerticalTextAlignment = ExcelRichTextRun.GetVerticalTextAlignment(properties),
                Outline = properties?.GetFirstChild<Outline>() != null,
                Shadow = properties?.GetFirstChild<Shadow>() != null,
                Condense = properties?.GetFirstChild<Condense>() != null,
                Extend = properties?.GetFirstChild<Extend>() != null,
                FontFamily = ExcelRichTextRun.GetFontFamily(properties),
                FontCharacterSet = ExcelRichTextRun.GetFontCharacterSet(properties)
            });
        }

        private static bool IsLineBreak(OpenXmlElement element) =>
            string.Equals(element.LocalName, "br", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(element.LocalName, "brk", StringComparison.OrdinalIgnoreCase);

        private static string? NullIfWhiteSpace(string? value) {
            return string.IsNullOrWhiteSpace(value) ? null : value;
        }

        private static string NormalizeMultilineText(string value) {
            return value
                .Replace("\r\n", "\n")
                .Replace('\r', '\n');
        }
    }
}
