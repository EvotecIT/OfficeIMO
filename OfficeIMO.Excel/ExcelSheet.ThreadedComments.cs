using DocumentFormat.OpenXml.Packaging;
using Threaded = DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Options used when authoring a threaded comment.
    /// </summary>
    public sealed class ExcelThreadedCommentOptions {
        /// <summary>Cell address in A1 notation.</summary>
        public string Address { get; set; } = "A1";

        /// <summary>Comment text.</summary>
        public string Text { get; set; } = string.Empty;

        /// <summary>Display author stored in workbook person metadata.</summary>
        public string Author { get; set; } = "OfficeIMO";

        /// <summary>Optional parent threaded-comment id when adding a reply.</summary>
        public string? ParentId { get; set; }

        /// <summary>Optional stable comment id. A new GUID is generated when omitted.</summary>
        public string? Id { get; set; }

        /// <summary>Optional timestamp. UTC now is used when omitted.</summary>
        public DateTime? Date { get; set; }

        /// <summary>Marks the threaded comment as resolved/done.</summary>
        public bool Done { get; set; }
    }

    /// <summary>
    /// Result returned after authoring a threaded comment.
    /// </summary>
    public sealed class ExcelThreadedCommentResult {
        internal ExcelThreadedCommentResult(string sheetName, string cellReference, string id, string personId, string author, bool isReply, bool done) {
            SheetName = sheetName;
            CellReference = cellReference;
            Id = id;
            PersonId = personId;
            Author = author;
            IsReply = isReply;
            Done = done;
        }

        /// <summary>Worksheet name.</summary>
        public string SheetName { get; }

        /// <summary>Cell address in A1 notation.</summary>
        public string CellReference { get; }

        /// <summary>Threaded comment id.</summary>
        public string Id { get; }

        /// <summary>Workbook person id used by the threaded comment.</summary>
        public string PersonId { get; }

        /// <summary>Resolved author display name.</summary>
        public string Author { get; }

        /// <summary>True when the comment references a parent threaded comment.</summary>
        public bool IsReply { get; }

        /// <summary>True when the comment is marked resolved/done.</summary>
        public bool Done { get; }
    }

    public partial class ExcelSheet {
        /// <summary>
        /// Adds a threaded comment or threaded reply to the worksheet while maintaining workbook person metadata.
        /// </summary>
        /// <param name="options">Threaded comment options.</param>
        public ExcelThreadedCommentResult AddThreadedComment(ExcelThreadedCommentOptions options) {
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (string.IsNullOrWhiteSpace(options.Address)) throw new ArgumentException("Threaded comment address is required.", nameof(options));
            if (string.IsNullOrWhiteSpace(options.Text)) throw new ArgumentException("Threaded comment text is required.", nameof(options));

            var (row, column) = A1.ParseCellRef(options.Address);
            if (row <= 0 || column <= 0) {
                throw new ArgumentException($"Address '{options.Address}' is not a valid A1 reference.", nameof(options));
            }

            string cellReference = A1.CellReference(row, column);
            string author = string.IsNullOrWhiteSpace(options.Author) ? "OfficeIMO" : options.Author.Trim();
            string id = NormalizeThreadedId(options.Id);
            string? parentId = string.IsNullOrWhiteSpace(options.ParentId) ? null : options.ParentId!.Trim();
            DateTime timestamp = options.Date ?? DateTime.UtcNow;
            string personId = string.Empty;

            WriteLock(() => {
                personId = EnsureWorkbookPerson(author);
                WorksheetThreadedCommentsPart commentsPart = GetOrCreateThreadedCommentsPart();
                commentsPart.ThreadedComments ??= new Threaded.ThreadedComments();
                RemoveThreadedCommentById(commentsPart.ThreadedComments, id);

                var comment = new Threaded.ThreadedComment {
                    Ref = cellReference,
                    PersonId = personId,
                    Id = id,
                    DT = timestamp
                };
                if (parentId != null) {
                    comment.ParentId = parentId;
                }
                if (options.Done) {
                    comment.Done = true;
                }

                comment.Append(new Threaded.ThreadedCommentText(options.Text));
                commentsPart.ThreadedComments.Append(comment);
                commentsPart.ThreadedComments.Save();
            });

            return new ExcelThreadedCommentResult(Name, cellReference, id, personId, author, parentId != null, options.Done);
        }

        /// <summary>
        /// Adds a threaded comment or threaded reply to the worksheet.
        /// </summary>
        public ExcelThreadedCommentResult AddThreadedComment(string address, string text, string author = "OfficeIMO", string? parentId = null, bool done = false) {
            return AddThreadedComment(new ExcelThreadedCommentOptions {
                Address = address,
                Text = text,
                Author = author,
                ParentId = parentId,
                Done = done
            });
        }

        private WorksheetThreadedCommentsPart GetOrCreateThreadedCommentsPart() {
            return _worksheetPart.WorksheetThreadedCommentsParts.FirstOrDefault()
                ?? _worksheetPart.AddNewPart<WorksheetThreadedCommentsPart>();
        }

        private string EnsureWorkbookPerson(string author) {
            WorkbookPart workbookPart = _spreadSheetDocument.WorkbookPart ?? throw new InvalidOperationException("Workbook part is missing.");
            foreach (WorkbookPersonPart part in workbookPart.WorkbookPersonParts) {
                if (part.PersonList == null) {
                    continue;
                }

                foreach (Threaded.Person person in part.PersonList.Elements<Threaded.Person>()) {
                    if (string.Equals(person.DisplayName?.Value, author, StringComparison.OrdinalIgnoreCase)
                        && !string.IsNullOrWhiteSpace(person.Id?.Value)) {
                        return person.Id!.Value!;
                    }
                }
            }

            WorkbookPersonPart personPart = workbookPart.WorkbookPersonParts.FirstOrDefault()
                ?? workbookPart.AddNewPart<WorkbookPersonPart>();
            personPart.PersonList ??= new Threaded.PersonList();
            string personId = BracedGuid();
            personPart.PersonList.Append(new Threaded.Person {
                Id = personId,
                DisplayName = author
            });
            personPart.PersonList.Save();
            return personId;
        }

        private static string NormalizeThreadedId(string? id) {
            return string.IsNullOrWhiteSpace(id) ? BracedGuid() : id!.Trim();
        }

        private static string BracedGuid() {
            return "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}";
        }

        private static void RemoveThreadedCommentById(Threaded.ThreadedComments comments, string id) {
            Threaded.ThreadedComment? existing = comments.Elements<Threaded.ThreadedComment>()
                .FirstOrDefault(comment => string.Equals(comment.Id?.Value, id, StringComparison.OrdinalIgnoreCase));
            existing?.Remove();
        }
    }
}
