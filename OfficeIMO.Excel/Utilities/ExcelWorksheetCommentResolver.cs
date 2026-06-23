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
            IReadOnlyDictionary<string, string> people) {
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
                } else if (element is Break) {
                    builder.Append('\n');
                }
            }

            return NormalizeMultilineText(builder.ToString());
        }

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
