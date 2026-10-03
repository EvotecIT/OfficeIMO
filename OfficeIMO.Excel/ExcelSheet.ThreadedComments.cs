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
            return AddThreadedComments(new[] { options })[0];
        }

        /// <summary>
        /// Adds comments and replies as one prevalidated batch, indexing workbook identities and saving each part once.
        /// Replies must refer to an existing root or an earlier root in the batch, on this worksheet and cell.
        /// </summary>
        /// <param name="options">Comment options, captured before package mutation.</param>
        /// <returns>Results in input order.</returns>
        public IReadOnlyList<ExcelThreadedCommentResult> AddThreadedComments(IEnumerable<ExcelThreadedCommentOptions> options) {
            if (options == null) throw new ArgumentNullException(nameof(options));
            var prepared = new List<(string Address, string Text, string Author, string Id, string? ParentId, DateTime Date, bool Done)>();
            foreach (ExcelThreadedCommentOptions item in options) {
                if (item == null) throw new ArgumentException("Comment options cannot contain null.", nameof(options));
                if (string.IsNullOrWhiteSpace(item.Address)) throw new ArgumentException("Threaded comment address is required.", nameof(options));
                if (string.IsNullOrWhiteSpace(item.Text)) throw new ArgumentException("Threaded comment text is required.", nameof(options));
                prepared.Add((NormalizeThreadedCommentAddress(item.Address, nameof(options)), item.Text,
                    string.IsNullOrWhiteSpace(item.Author) ? "OfficeIMO" : item.Author.Trim(),
                    NormalizeThreadedId(item.Id, nameof(item.Id), generateIfMissing: true),
                    string.IsNullOrWhiteSpace(item.ParentId) ? null : NormalizeThreadedId(item.ParentId, nameof(item.ParentId), generateIfMissing: false),
                    NormalizeThreadedTimestamp(item.Date), item.Done));
            }
            if (prepared.Count == 0) return Array.Empty<ExcelThreadedCommentResult>();
            var results = new List<ExcelThreadedCommentResult>(prepared.Count);
            WriteLock(() => {
                WorkbookPart workbook = _spreadSheetDocument.WorkbookPart ?? throw new InvalidOperationException("Workbook part is missing.");
                var identities = new Dictionary<string, (WorksheetPart Worksheet, string? Address, bool IsReply)>(StringComparer.OrdinalIgnoreCase);
                foreach (WorksheetPart worksheet in workbook.WorksheetParts) {
                    foreach (WorksheetThreadedCommentsPart part in worksheet.WorksheetThreadedCommentsParts) {
                        if (part.ThreadedComments == null) continue;
                        foreach (Threaded.ThreadedComment existing in part.ThreadedComments.Elements<Threaded.ThreadedComment>()) {
                            string? id = existing.Id?.Value;
                            if (!string.IsNullOrWhiteSpace(id) && !identities.ContainsKey(id!))
                                identities.Add(id!, (worksheet, existing.Ref?.Value, !string.IsNullOrWhiteSpace(existing.ParentId?.Value)));
                        }
                    }
                }
                // Validate the complete ID/parent plan before creating parts, people or comments.
                foreach (var item in prepared) {
                    if (identities.ContainsKey(item.Id)) throw new InvalidOperationException($"A threaded comment with id '{item.Id}' already exists in the workbook.");
                    if (item.ParentId != null) {
                        if (!identities.TryGetValue(item.ParentId, out var parent))
                            throw new ArgumentException($"Parent threaded comment '{item.ParentId}' does not exist.", nameof(options));
                        if (!ReferenceEquals(parent.Worksheet, _worksheetPart)
                            || !string.Equals(parent.Address, item.Address, StringComparison.OrdinalIgnoreCase))
                            throw new ArgumentException("A threaded reply must use the same worksheet and cell as its parent comment.", nameof(options));
                        if (parent.IsReply) throw new ArgumentException("A threaded reply must reference the root comment rather than another reply.", nameof(options));
                    }
                    identities.Add(item.Id, (_worksheetPart, item.Address, item.ParentId != null));
                }
                var people = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (WorkbookPersonPart part in workbook.WorkbookPersonParts) {
                    if (part.PersonList == null) continue;
                    foreach (Threaded.Person person in part.PersonList.Elements<Threaded.Person>()) {
                        string? name = person.DisplayName?.Value, id = person.Id?.Value;
                        if (name != null && !string.IsNullOrWhiteSpace(id) && !people.ContainsKey(name)) people.Add(name, id!);
                    }
                }
                WorksheetThreadedCommentsPart commentsPart = GetOrCreateThreadedCommentsPart();
                commentsPart.ThreadedComments ??= new Threaded.ThreadedComments();
                foreach (var item in prepared) {
                    string personId = EnsureWorkbookPerson(item.Author, people);
                    var comment = new Threaded.ThreadedComment { Ref = item.Address, PersonId = personId, Id = item.Id, DT = item.Date };
                    if (item.ParentId != null) comment.ParentId = item.ParentId;
                    if (item.Done) comment.Done = true;
                    comment.Append(new Threaded.ThreadedCommentText(item.Text));
                    commentsPart.ThreadedComments.Append(comment);
                    results.Add(new ExcelThreadedCommentResult(Name, item.Address, item.Id, personId, item.Author, item.ParentId != null, item.Done));
                }
                commentsPart.ThreadedComments.Save();
                foreach (WorkbookPersonPart part in workbook.WorkbookPersonParts) part.PersonList?.Save();
                _excelDocument.MarkPackageDirty();
            });
            return results.AsReadOnly();
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

        private string EnsureWorkbookPerson(string author, IDictionary<string, string>? indexedPeople = null) {
            if (indexedPeople != null && indexedPeople.TryGetValue(author, out string? known)) return known;
            WorkbookPart workbookPart = _spreadSheetDocument.WorkbookPart ?? throw new InvalidOperationException("Workbook part is missing.");
            if (indexedPeople == null) foreach (WorkbookPersonPart part in workbookPart.WorkbookPersonParts) {
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
            if (indexedPeople == null) personPart.PersonList.Save();
            else indexedPeople.Add(author, personId);
            return personId;
        }

        private static string NormalizeThreadedCommentAddress(string address, string parameterName) {
            var (row, column) = A1.ParseCellRef(address);
            if (row <= 0 || column <= 0 || row > A1.MaxRows || column > A1.MaxColumns) {
                throw new ArgumentException($"Address '{address}' is not a valid A1 reference.", parameterName);
            }

            return A1.CellReference(row, column);
        }

        private static string NormalizeThreadedId(string? id, string parameterName, bool generateIfMissing) {
            if (string.IsNullOrWhiteSpace(id)) {
                if (generateIfMissing) {
                    return BracedGuid();
                }

                throw new ArgumentNullException(parameterName);
            }

            if (!Guid.TryParse(id!.Trim(), out Guid guid)) {
                throw new ArgumentException("Threaded comment ids must be GUID values.", parameterName);
            }

            return "{" + guid.ToString().ToUpperInvariant() + "}";
        }

        private static DateTime NormalizeThreadedTimestamp(DateTime? value) {
            DateTime timestamp = value ?? DateTime.UtcNow;
            if (timestamp.Kind == DateTimeKind.Local) {
                return timestamp.ToUniversalTime();
            }

            return timestamp.Kind == DateTimeKind.Unspecified
                ? DateTime.SpecifyKind(timestamp, DateTimeKind.Utc)
                : timestamp;
        }

        private static string BracedGuid() {
            return "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}";
        }

        private bool TryFindWorkbookThreadedComment(
            string id,
            out WorksheetPart? worksheetPart,
            out WorksheetThreadedCommentsPart? commentsPart,
            out Threaded.ThreadedComment? comment) {
            WorkbookPart workbookPart = _spreadSheetDocument.WorkbookPart ?? throw new InvalidOperationException("Workbook part is missing.");
            foreach (WorksheetPart candidateWorksheet in workbookPart.WorksheetParts) {
                foreach (WorksheetThreadedCommentsPart candidatePart in candidateWorksheet.WorksheetThreadedCommentsParts) {
                    Threaded.ThreadedComment? candidate = candidatePart.ThreadedComments?
                        .Elements<Threaded.ThreadedComment>()
                        .FirstOrDefault(item => string.Equals(item.Id?.Value, id, StringComparison.OrdinalIgnoreCase));
                    if (candidate != null) {
                        worksheetPart = candidateWorksheet;
                        commentsPart = candidatePart;
                        comment = candidate;
                        return true;
                    }
                }
            }

            worksheetPart = null;
            commentsPart = null;
            comment = null;
            return false;
        }
    }
}
