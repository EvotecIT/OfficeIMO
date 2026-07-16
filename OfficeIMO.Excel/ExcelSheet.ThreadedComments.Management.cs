using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel.Utilities;
using Threaded = DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Options used when updating an existing threaded comment.
    /// </summary>
    public sealed class ExcelThreadedCommentUpdateOptions {
        /// <summary>Threaded comment id.</summary>
        public string Id { get; set; } = string.Empty;

        /// <summary>Replacement text. Null preserves the current text.</summary>
        public string? Text { get; set; }

        /// <summary>Replacement author. Null preserves the current person.</summary>
        public string? Author { get; set; }

        /// <summary>Replacement timestamp. Null preserves the current timestamp.</summary>
        public DateTime? Date { get; set; }

        /// <summary>Replacement resolved/done state. Null preserves the current state.</summary>
        public bool? Done { get; set; }
    }

    public partial class ExcelSheet {
        /// <summary>
        /// Gets all threaded comments and replies on this worksheet.
        /// </summary>
        public IReadOnlyList<ExcelThreadedCommentSnapshot> GetThreadedComments() {
            return Locking.ExecuteRead(_excelDocument.EnsureLock(), () => {
                WorkbookPart workbookPart = _spreadSheetDocument.WorkbookPart ?? throw new InvalidOperationException("Workbook part is missing.");
                IReadOnlyDictionary<string, string> people = ExcelWorksheetCommentResolver.BuildThreadedCommentPersonMap(workbookPart);
                return ExcelWorksheetCommentResolver.BuildThreadedCommentMap(_worksheetPart, people, Name)
                    .Values
                    .SelectMany(comments => comments)
                    .ToArray();
            });
        }

        /// <summary>
        /// Gets a threaded comment or reply by id, or null when it is not present on this worksheet.
        /// </summary>
        public ExcelThreadedCommentSnapshot? GetThreadedComment(string id) {
            string normalizedId = NormalizeThreadedId(id, nameof(id), generateIfMissing: false);
            return GetThreadedComments()
                .FirstOrDefault(comment => string.Equals(comment.Id, normalizedId, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Adds a reply using the worksheet and cell of an existing root threaded comment.
        /// </summary>
        public ExcelThreadedCommentResult ReplyToThreadedComment(
            string parentId,
            string text,
            string author = "OfficeIMO",
            DateTime? date = null,
            string? id = null) {
            ExcelThreadedCommentSnapshot parent = GetThreadedComment(parentId)
                ?? throw new KeyNotFoundException($"Threaded comment '{parentId}' was not found on worksheet '{Name}'.");
            if (!string.IsNullOrWhiteSpace(parent.ParentId)) {
                throw new ArgumentException("A threaded reply must reference the root comment rather than another reply.", nameof(parentId));
            }

            return AddThreadedComment(new ExcelThreadedCommentOptions {
                Address = parent.CellReference,
                Text = text,
                Author = author,
                ParentId = parent.Id,
                Date = date,
                Id = id
            });
        }

        /// <summary>
        /// Updates the text, author, timestamp, or resolved state of an existing threaded comment.
        /// </summary>
        public ExcelThreadedCommentResult UpdateThreadedComment(ExcelThreadedCommentUpdateOptions options) {
            if (options == null) throw new ArgumentNullException(nameof(options));
            string id = NormalizeThreadedId(options.Id, nameof(options.Id), generateIfMissing: false);
            if (options.Text != null && string.IsNullOrWhiteSpace(options.Text)) {
                throw new ArgumentException("Threaded comment text cannot be empty.", nameof(options));
            }
            if (options.Author != null && string.IsNullOrWhiteSpace(options.Author)) {
                throw new ArgumentException("Threaded comment author cannot be empty.", nameof(options));
            }

            ExcelThreadedCommentResult? result = null;
            WriteLock(() => {
                if (!TryFindWorkbookThreadedComment(id, out WorksheetPart? worksheetPart, out WorksheetThreadedCommentsPart? commentsPart, out Threaded.ThreadedComment? comment)
                    || !ReferenceEquals(worksheetPart, _worksheetPart)
                    || commentsPart?.ThreadedComments == null
                    || comment == null) {
                    throw new KeyNotFoundException($"Threaded comment '{id}' was not found on worksheet '{Name}'.");
                }

                string personId = comment.PersonId?.Value ?? string.Empty;
                string author;
                if (options.Author != null) {
                    author = options.Author.Trim();
                    personId = EnsureWorkbookPerson(author);
                    comment.PersonId = personId;
                } else {
                    author = ResolveWorkbookPersonName(personId) ?? string.Empty;
                }

                if (options.Text != null) {
                    comment.RemoveAllChildren<Threaded.ThreadedCommentText>();
                    comment.Append(new Threaded.ThreadedCommentText(options.Text));
                }
                if (options.Date.HasValue) {
                    comment.DT = NormalizeThreadedTimestamp(options.Date);
                }
                if (options.Done.HasValue) {
                    comment.Done = options.Done.Value;
                }

                commentsPart.ThreadedComments.Save();
                _excelDocument.MarkPackageDirty();
                string cellReference = comment.Ref?.Value ?? string.Empty;
                result = new ExcelThreadedCommentResult(
                    Name,
                    cellReference,
                    id,
                    personId,
                    author,
                    !string.IsNullOrWhiteSpace(comment.ParentId?.Value),
                    comment.Done?.Value == true);
            });

            return result!;
        }

        /// <summary>
        /// Marks a threaded comment resolved or reopens it.
        /// </summary>
        public ExcelThreadedCommentResult SetThreadedCommentResolved(string id, bool resolved = true) {
            return UpdateThreadedComment(new ExcelThreadedCommentUpdateOptions {
                Id = id,
                Done = resolved
            });
        }

        /// <summary>
        /// Removes a threaded comment. Root-comment removal includes its reply tree by default.
        /// </summary>
        /// <param name="id">Threaded comment id.</param>
        /// <param name="removeReplies">Whether replies and imported descendant replies should be removed with the comment.</param>
        public bool RemoveThreadedComment(string id, bool removeReplies = true) {
            string normalizedId = NormalizeThreadedId(id, nameof(id), generateIfMissing: false);
            bool removed = false;
            WriteLock(() => {
                if (!TryFindWorkbookThreadedComment(normalizedId, out WorksheetPart? worksheetPart, out _, out _)
                    || !ReferenceEquals(worksheetPart, _worksheetPart)) {
                    return;
                }

                List<Threaded.ThreadedComment> comments = _worksheetPart.WorksheetThreadedCommentsParts
                    .SelectMany(part => part.ThreadedComments?.Elements<Threaded.ThreadedComment>() ?? Enumerable.Empty<Threaded.ThreadedComment>())
                    .ToList();
                var idsToRemove = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { normalizedId };
                bool changed;
                do {
                    changed = false;
                    foreach (Threaded.ThreadedComment comment in comments) {
                        string? commentId = comment.Id?.Value;
                        string? parentId = comment.ParentId?.Value;
                        if (!string.IsNullOrWhiteSpace(commentId)
                            && !string.IsNullOrWhiteSpace(parentId)
                            && idsToRemove.Contains(parentId!)
                            && idsToRemove.Add(commentId!)) {
                            changed = true;
                        }
                    }
                } while (changed);

                if (!removeReplies && idsToRemove.Count > 1) {
                    throw new InvalidOperationException("The threaded comment has replies. Remove the complete thread or remove the replies first.");
                }

                foreach (WorksheetThreadedCommentsPart part in _worksheetPart.WorksheetThreadedCommentsParts.ToList()) {
                    Threaded.ThreadedComments? root = part.ThreadedComments;
                    if (root == null) {
                        continue;
                    }

                    int partRemoved = 0;
                    foreach (Threaded.ThreadedComment comment in root.Elements<Threaded.ThreadedComment>().ToList()) {
                        if (!string.IsNullOrWhiteSpace(comment.Id?.Value) && idsToRemove.Contains(comment.Id!.Value!)) {
                            comment.Remove();
                            partRemoved++;
                            removed = true;
                        }
                    }

                    if (!root.Elements<Threaded.ThreadedComment>().Any()) {
                        _worksheetPart.DeletePart(part);
                    } else if (partRemoved > 0) {
                        root.Save();
                    }
                }

                if (removed) {
                    _excelDocument.MarkPackageDirty();
                }
            });

            return removed;
        }

        /// <summary>
        /// Removes every threaded comment and reply attached to a worksheet cell.
        /// </summary>
        public int RemoveThreadedCommentsAt(string address) {
            string cellReference = NormalizeThreadedCommentAddress(address, nameof(address));
            int removed = 0;
            WriteLock(() => {
                foreach (WorksheetThreadedCommentsPart part in _worksheetPart.WorksheetThreadedCommentsParts.ToList()) {
                    Threaded.ThreadedComments? root = part.ThreadedComments;
                    if (root == null) {
                        continue;
                    }

                    int partRemoved = 0;
                    foreach (Threaded.ThreadedComment comment in root.Elements<Threaded.ThreadedComment>().ToList()) {
                        if (string.Equals(comment.Ref?.Value, cellReference, StringComparison.OrdinalIgnoreCase)) {
                            comment.Remove();
                            partRemoved++;
                        }
                    }

                    removed += partRemoved;
                    if (!root.Elements<Threaded.ThreadedComment>().Any()) {
                        _worksheetPart.DeletePart(part);
                    } else if (partRemoved > 0) {
                        root.Save();
                    }
                }

                if (removed > 0) {
                    _excelDocument.MarkPackageDirty();
                }
            });

            return removed;
        }

        private string? ResolveWorkbookPersonName(string personId) {
            WorkbookPart workbookPart = _spreadSheetDocument.WorkbookPart ?? throw new InvalidOperationException("Workbook part is missing.");
            IReadOnlyDictionary<string, string> people = ExcelWorksheetCommentResolver.BuildThreadedCommentPersonMap(workbookPart);
            return people.TryGetValue(personId, out string? author) ? author : null;
        }
    }
}
