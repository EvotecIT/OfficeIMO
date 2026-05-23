using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Xml.Linq;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Immutable worksheet comment metadata.
    /// </summary>
    public sealed class ExcelCommentInfo {
        internal ExcelCommentInfo(string cellReference, int row, int column, string? author, string text, IReadOnlyList<ExcelRichTextRun>? richTextRuns = null) {
            CellReference = cellReference;
            Row = row;
            Column = column;
            Author = author;
            Text = text;
            RichTextRuns = richTextRuns ?? Array.Empty<ExcelRichTextRun>();
        }

        /// <summary>A1 cell reference where the comment is attached.</summary>
        public string CellReference { get; }

        /// <summary>1-based row index where the comment is attached.</summary>
        public int Row { get; }

        /// <summary>1-based column index where the comment is attached.</summary>
        public int Column { get; }

        /// <summary>Comment author display name, when available.</summary>
        public string? Author { get; }

        /// <summary>Comment text content.</summary>
        public string Text { get; }

        /// <summary>Rich text runs stored in the legacy comment text.</summary>
        public IReadOnlyList<ExcelRichTextRun> RichTextRuns { get; }
    }

    /// <summary>
    /// Filters worksheet comments by author, text, and A1 range.
    /// </summary>
    public sealed class ExcelCommentFilter {
        /// <summary>Only match comments whose author equals this value, ignoring case.</summary>
        public string? Author { get; set; }

        /// <summary>Only match comments whose text contains this value, ignoring case.</summary>
        public string? TextContains { get; set; }

        /// <summary>Only match comments attached to cells inside this A1 cell or range.</summary>
        public string? A1Range { get; set; }
    }

    /// <summary>
    /// Helpers for worksheet cell comments (notes).
    /// </summary>
    public partial class ExcelSheet {
        /// <summary>
        /// Adds or replaces a comment on the specified cell.
        /// </summary>
        /// <param name="row">1-based row index.</param>
        /// <param name="column">1-based column index.</param>
        /// <param name="text">Comment text.</param>
        /// <param name="author">Author name (optional).</param>
        /// <param name="initials">Author initials (optional).</param>
        public void SetComment(int row, int column, string text, string author = "OfficeIMO", string? initials = null) {
            if (row <= 0) throw new ArgumentOutOfRangeException(nameof(row), "Row and column are 1-based and must be positive.");
            if (column <= 0) throw new ArgumentOutOfRangeException(nameof(column), "Row and column are 1-based and must be positive.");
            if (string.IsNullOrEmpty(text)) throw new ArgumentException("Comment text is required.", nameof(text));
            SetCommentInternal(row, column, BuildCommentText(text), author, initials);
        }

        /// <summary>
        /// Adds or replaces a rich-text comment on the specified cell.
        /// </summary>
        /// <param name="row">1-based row index.</param>
        /// <param name="column">1-based column index.</param>
        /// <param name="runs">Rich text runs that make up the comment text.</param>
        /// <param name="author">Author name (optional).</param>
        /// <param name="initials">Author initials (optional).</param>
        public void SetCommentRichText(int row, int column, IEnumerable<ExcelRichTextRun> runs, string author = "OfficeIMO", string? initials = null) {
            if (row <= 0) throw new ArgumentOutOfRangeException(nameof(row), "Row and column are 1-based and must be positive.");
            if (column <= 0) throw new ArgumentOutOfRangeException(nameof(column), "Row and column are 1-based and must be positive.");
            SetCommentInternal(row, column, BuildCommentText(runs), author, initials);
        }

        private void SetCommentInternal(int row, int column, CommentText commentText, string author, string? initials) {
            WriteLock(() => {
                string reference = A1.CellReference(row, column);
                string authorDisplay = NormalizeAuthor(author, initials);

                var commentsPart = GetOrCreateCommentsPart();
                var comments = commentsPart.Comments ??= new Comments();
                comments.Authors ??= new Authors();
                comments.CommentList ??= new CommentList();

                uint authorId = EnsureAuthorId(comments.Authors, authorDisplay);
                RemoveCommentInternal(comments.CommentList, reference);

                var comment = new Comment { Reference = reference, AuthorId = authorId };
                comment.Append(commentText);
                comments.CommentList.Append(comment);
                comments.Save();

                EnsureCommentVmlShape(row, column);
                WorksheetRoot.Save();
            });
        }

        /// <summary>
        /// Adds or replaces a comment on the specified A1 cell reference.
        /// </summary>
        /// <param name="a1">A1 cell reference (e.g., "B5").</param>
        /// <param name="text">Comment text.</param>
        /// <param name="author">Author name (optional).</param>
        /// <param name="initials">Author initials (optional).</param>
        public void SetComment(string a1, string text, string author = "OfficeIMO", string? initials = null) {
            var (row, col) = A1.ParseCellRef(a1);
            if (row <= 0 || col <= 0) throw new ArgumentException($"Address '{a1}' is not a valid A1 reference.", nameof(a1));
            SetComment(row, col, text, author, initials);
        }

        /// <summary>
        /// Adds or replaces a rich-text comment on the specified A1 cell reference.
        /// </summary>
        /// <param name="a1">A1 cell reference (e.g., "B5").</param>
        /// <param name="runs">Rich text runs that make up the comment text.</param>
        /// <param name="author">Author name (optional).</param>
        /// <param name="initials">Author initials (optional).</param>
        public void SetCommentRichText(string a1, IEnumerable<ExcelRichTextRun> runs, string author = "OfficeIMO", string? initials = null) {
            var (row, col) = A1.ParseCellRef(a1);
            if (row <= 0 || col <= 0) throw new ArgumentException($"Address '{a1}' is not a valid A1 reference.", nameof(a1));
            SetCommentRichText(row, col, runs, author, initials);
        }

        /// <summary>
        /// Removes a comment from the specified cell (if present).
        /// </summary>
        /// <param name="row">1-based row index.</param>
        /// <param name="column">1-based column index.</param>
        public void ClearComment(int row, int column) {
            if (row <= 0) throw new ArgumentOutOfRangeException(nameof(row), "Row and column are 1-based and must be positive.");
            if (column <= 0) throw new ArgumentOutOfRangeException(nameof(column), "Row and column are 1-based and must be positive.");

            WriteLock(() => {
                string reference = A1.CellReference(row, column);

                var commentsPart = WorksheetCommentsPartRoot;
                if (commentsPart?.Comments?.CommentList == null) {
                    RemoveCommentVmlShape(row, column);
                    return;
                }

                bool removedComment = RemoveCommentInternal(commentsPart.Comments.CommentList, reference);
                if (removedComment) {
                    commentsPart.Comments.Save();
                }

                RemoveCommentVmlShape(row, column);
                bool removedArtifacts = CleanupCommentArtifacts();
                if (removedArtifacts) {
                    WorksheetRoot.Save();
                }
            });
        }

        /// <summary>
        /// Removes a comment from the specified A1 cell reference (if present).
        /// </summary>
        /// <param name="a1">A1 cell reference (e.g., "B5").</param>
        public void ClearComment(string a1) {
            var (row, col) = A1.ParseCellRef(a1);
            if (row <= 0 || col <= 0) throw new ArgumentException($"Address '{a1}' is not a valid A1 reference.", nameof(a1));
            ClearComment(row, col);
        }

        /// <summary>
        /// Returns true when a comment exists for the specified cell.
        /// </summary>
        /// <param name="row">1-based row index.</param>
        /// <param name="column">1-based column index.</param>
        public bool HasComment(int row, int column) {
            if (row <= 0) throw new ArgumentOutOfRangeException(nameof(row), "Row and column are 1-based and must be positive.");
            if (column <= 0) throw new ArgumentOutOfRangeException(nameof(column), "Row and column are 1-based and must be positive.");
            string reference = A1.CellReference(row, column);
            var commentsPart = WorksheetCommentsPartRoot;
            return commentsPart?.Comments?.CommentList?
                .Elements<Comment>()
                .Any(c => string.Equals(c.Reference?.Value, reference, StringComparison.OrdinalIgnoreCase)) is true;
        }

        /// <summary>
        /// Gets all legacy worksheet comments (notes) on this sheet.
        /// </summary>
        public IReadOnlyList<ExcelCommentInfo> GetComments() {
            return FindComments(null);
        }

        /// <summary>
        /// Finds legacy worksheet comments (notes) that match the supplied filter.
        /// </summary>
        /// <param name="filter">Optional author, text, and A1 range filter.</param>
        public IReadOnlyList<ExcelCommentInfo> FindComments(ExcelCommentFilter? filter) {
            var commentsPart = WorksheetCommentsPartRoot;
            var comments = commentsPart?.Comments;
            if (comments?.CommentList == null) {
                return Array.Empty<ExcelCommentInfo>();
            }

            var authors = comments.Authors?.Elements<Author>().Select(author => author.Text ?? string.Empty).ToList()
                ?? new List<string>();
            var results = new List<ExcelCommentInfo>();
            foreach (var comment in comments.CommentList.Elements<Comment>()) {
                var info = CreateCommentInfo(comment, authors);
                if (info != null && CommentMatchesFilter(info, filter)) {
                    results.Add(info);
                }
            }

            return results;
        }

        /// <summary>
        /// Replaces the text, and optionally author, for comments that match the supplied filter.
        /// </summary>
        /// <param name="filter">Author, text, and/or A1 range filter used to choose comments.</param>
        /// <param name="text">Replacement comment text.</param>
        /// <param name="author">Optional replacement author.</param>
        /// <param name="initials">Optional replacement author initials.</param>
        /// <returns>Number of comments updated.</returns>
        public int UpdateComments(ExcelCommentFilter filter, string text, string? author = null, string? initials = null) {
            if (filter == null) throw new ArgumentNullException(nameof(filter));
            if (string.IsNullOrEmpty(text)) throw new ArgumentException("Comment text is required.", nameof(text));

            return UpdateCommentsInternal(filter, BuildCommentText(text), author, initials);
        }

        /// <summary>
        /// Replaces rich text, and optionally author, for comments that match the supplied filter.
        /// </summary>
        /// <param name="filter">Author, text, and/or A1 range filter used to choose comments.</param>
        /// <param name="runs">Replacement rich text runs.</param>
        /// <param name="author">Optional replacement author.</param>
        /// <param name="initials">Optional replacement author initials.</param>
        /// <returns>Number of comments updated.</returns>
        public int UpdateCommentsRichText(ExcelCommentFilter filter, IEnumerable<ExcelRichTextRun> runs, string? author = null, string? initials = null) {
            if (filter == null) throw new ArgumentNullException(nameof(filter));

            return UpdateCommentsInternal(filter, BuildCommentText(runs), author, initials);
        }

        private int UpdateCommentsInternal(ExcelCommentFilter filter, CommentText commentText, string? author, string? initials) {
            int updated = 0;
            WriteLock(() => {
                var commentsPart = WorksheetCommentsPartRoot;
                var comments = commentsPart?.Comments;
                if (comments?.CommentList == null) {
                    return;
                }

                comments.Authors ??= new Authors();
                var authors = comments.Authors.Elements<Author>().Select(item => item.Text ?? string.Empty).ToList();
                uint? newAuthorId = string.IsNullOrWhiteSpace(author)
                    ? null
                    : EnsureAuthorId(comments.Authors, NormalizeAuthor(author!, initials));

                foreach (var comment in comments.CommentList.Elements<Comment>()) {
                    var info = CreateCommentInfo(comment, authors);
                    if (info == null || !CommentMatchesFilter(info, filter)) {
                        continue;
                    }

                    comment.RemoveAllChildren<CommentText>();
                    comment.Append((CommentText)commentText.CloneNode(true));
                    if (newAuthorId.HasValue) {
                        comment.AuthorId = newAuthorId.Value;
                    }

                    updated++;
                }

                if (updated > 0) {
                    comments.Save();
                }
            });

            return updated;
        }

        /// <summary>
        /// Removes comments that match the supplied filter.
        /// </summary>
        /// <param name="filter">Author, text, and/or A1 range filter used to choose comments.</param>
        /// <returns>Number of comments removed.</returns>
        public int ClearComments(ExcelCommentFilter filter) {
            if (filter == null) throw new ArgumentNullException(nameof(filter));

            int removed = 0;
            WriteLock(() => {
                var commentsPart = WorksheetCommentsPartRoot;
                var comments = commentsPart?.Comments;
                if (comments?.CommentList == null) {
                    return;
                }

                var authors = comments.Authors?.Elements<Author>().Select(author => author.Text ?? string.Empty).ToList()
                    ?? new List<string>();
                var removals = new List<(Comment Comment, int Row, int Column)>();
                foreach (var comment in comments.CommentList.Elements<Comment>()) {
                    var info = CreateCommentInfo(comment, authors);
                    if (info != null && CommentMatchesFilter(info, filter)) {
                        removals.Add((comment, info.Row, info.Column));
                    }
                }

                foreach (var removal in removals) {
                    removal.Comment.Remove();
                    RemoveCommentVmlShape(removal.Row, removal.Column);
                }

                removed = removals.Count;
                if (removed > 0) {
                    comments.Save();
                    bool removedArtifacts = CleanupCommentArtifacts();
                    if (removedArtifacts) {
                        WorksheetRoot.Save();
                    }
                }
            });

            return removed;
        }

        private static string NormalizeAuthor(string author, string? initials) {
            string name = string.IsNullOrWhiteSpace(author) ? "OfficeIMO" : author.Trim();
            if (string.IsNullOrWhiteSpace(initials)) return name;
            return $"{name} ({initials!.Trim()})";
        }

        private static uint EnsureAuthorId(Authors authors, string authorDisplay) {
            int idx = 0;
            foreach (var a in authors.Elements<Author>()) {
                if (string.Equals(a.Text, authorDisplay, StringComparison.OrdinalIgnoreCase)) {
                    return (uint)idx;
                }
                idx++;
            }

            authors.Append(new Author { Text = authorDisplay });
            return (uint)idx;
        }

        private static bool RemoveCommentInternal(CommentList list, string reference) {
            var existing = list.Elements<Comment>()
                .FirstOrDefault(c => string.Equals(c.Reference?.Value, reference, StringComparison.OrdinalIgnoreCase));
            if (existing == null) {
                return false;
            }

            existing.Remove();
            return true;
        }

        private static CommentText BuildCommentText(string text) {
            return BuildCommentText(new[] { new ExcelRichTextRun(text) });
        }

        private static CommentText BuildCommentText(IEnumerable<ExcelRichTextRun> runs) {
            var normalizedRuns = NormalizeCommentRuns(runs);
            var commentText = new CommentText();
            foreach (var richRun in normalizedRuns) {
                var run = new Run();
                var properties = new RunProperties();
                if (richRun.Bold) properties.Append(new Bold());
                if (richRun.Italic) properties.Append(new Italic());
                if (richRun.Underline) properties.Append(new Underline());
                if (!string.IsNullOrWhiteSpace(richRun.FontColor)) properties.Append(new Color { Rgb = NormalizeHexColor(richRun.FontColor!) });
                if (!string.IsNullOrWhiteSpace(richRun.FontName)) properties.Append(new RunFont { Val = richRun.FontName });
                if (richRun.FontSize.HasValue) properties.Append(new FontSize { Val = richRun.FontSize.Value });
                if (properties.HasChildren) {
                    run.Append(properties);
                }

                run.Append(new Text(NormalizeCommentText(richRun.Text)) { Space = SpaceProcessingModeValues.Preserve });
                commentText.Append(run);
            }

            return commentText;
        }

        private static IReadOnlyList<ExcelRichTextRun> NormalizeCommentRuns(IEnumerable<ExcelRichTextRun> runs) {
            if (runs == null) throw new ArgumentNullException(nameof(runs));
            var normalized = new List<ExcelRichTextRun>();
            foreach (var run in runs) {
                if (run == null) continue;
                normalized.Add(new ExcelRichTextRun(run.Text ?? string.Empty) {
                    Bold = run.Bold,
                    Italic = run.Italic,
                    Underline = run.Underline,
                    FontColor = run.FontColor,
                    FontName = run.FontName,
                    FontSize = run.FontSize
                });
            }

            if (normalized.Count == 0 || normalized.All(run => string.IsNullOrEmpty(run.Text))) {
                throw new ArgumentException("At least one comment text run is required.", nameof(runs));
            }

            return normalized;
        }

        private static string NormalizeCommentText(string? text) {
            return (text ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n');
        }

        private static ExcelCommentInfo? CreateCommentInfo(Comment comment, IReadOnlyList<string> authors) {
            string? reference = comment.Reference?.Value;
            if (string.IsNullOrWhiteSpace(reference)) {
                return null;
            }

            var parsed = A1.ParseCellRef(reference!);
            if (parsed.Row <= 0 || parsed.Col <= 0) {
                return null;
            }

            string? author = null;
            if (comment.AuthorId != null && comment.AuthorId.Value < authors.Count) {
                author = authors[(int)comment.AuthorId.Value];
            }

            return new ExcelCommentInfo(reference!, parsed.Row, parsed.Col, author, ExtractCommentText(comment.CommentText), ExtractCommentRuns(comment.CommentText));
        }

        private static string ExtractCommentText(CommentText? commentText) {
            if (commentText == null) {
                return string.Empty;
            }

            return string.Concat(commentText.Descendants<Text>().Select(text => text.Text ?? string.Empty));
        }

        private static IReadOnlyList<ExcelRichTextRun> ExtractCommentRuns(CommentText? commentText) {
            if (commentText == null) {
                return Array.Empty<ExcelRichTextRun>();
            }

            var runs = new List<ExcelRichTextRun>();
            foreach (var run in commentText.Elements<Run>()) {
                var properties = run.RunProperties;
                runs.Add(new ExcelRichTextRun(run.Text?.Text ?? string.Empty) {
                    Bold = properties?.GetFirstChild<Bold>() != null,
                    Italic = properties?.GetFirstChild<Italic>() != null,
                    Underline = properties?.GetFirstChild<Underline>() != null,
                    FontColor = properties?.GetFirstChild<Color>()?.Rgb?.Value,
                    FontName = properties?.GetFirstChild<RunFont>()?.Val?.Value,
                    FontSize = properties?.GetFirstChild<FontSize>()?.Val?.Value
                });
            }

            if (runs.Count == 0) {
                string plainText = ExtractCommentText(commentText);
                if (!string.IsNullOrEmpty(plainText)) {
                    runs.Add(new ExcelRichTextRun(plainText));
                }
            }

            return runs;
        }

        private static bool CommentMatchesFilter(ExcelCommentInfo info, ExcelCommentFilter? filter) {
            if (filter == null) {
                return true;
            }

            if (!string.IsNullOrWhiteSpace(filter.Author)
                && !string.Equals(info.Author, filter.Author!.Trim(), StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            if (!string.IsNullOrWhiteSpace(filter.TextContains)
                && info.Text.IndexOf(filter.TextContains!.Trim(), StringComparison.OrdinalIgnoreCase) < 0) {
                return false;
            }

            if (!string.IsNullOrWhiteSpace(filter.A1Range)) {
                var bounds = ParseCommentFilterRange(filter.A1Range!);
                if (info.Row < bounds.FirstRow
                    || info.Row > bounds.LastRow
                    || info.Column < bounds.FirstColumn
                    || info.Column > bounds.LastColumn) {
                    return false;
                }
            }

            return true;
        }

        private static (int FirstRow, int FirstColumn, int LastRow, int LastColumn) ParseCommentFilterRange(string a1Range) {
            if (A1.TryParseRange(a1Range, out int r1, out int c1, out int r2, out int c2)) {
                return (r1, c1, r2, c2);
            }

            var cell = A1.ParseCellRef(a1Range);
            if (cell.Row > 0 && cell.Col > 0) {
                return (cell.Row, cell.Col, cell.Row, cell.Col);
            }

            throw new ArgumentException($"Address '{a1Range}' is not a valid A1 cell or range.", nameof(a1Range));
        }

        private WorksheetCommentsPart GetOrCreateCommentsPart() {
            var part = WorksheetCommentsPartRoot;
            if (part == null) {
                part = _worksheetPart.AddNewPart<WorksheetCommentsPart>();
                part.Comments = new Comments(new Authors(), new CommentList());
            }
            return part;
        }

        private void EnsureCommentVmlShape(int row, int column) {
            var vmlPart = GetOrCreateCommentVmlPart();
            var doc = LoadOrCreateVmlDocument(vmlPart);
            var root = doc.Root;
            if (root == null) return;

            RemoveVmlShape(root, row, column);

            var v = XNamespace.Get("urn:schemas-microsoft-com:vml");
            var o = XNamespace.Get("urn:schemas-microsoft-com:office:office");
            var x = XNamespace.Get("urn:schemas-microsoft-com:office:excel");

            int shapeId = NextVmlShapeId(root);
            string anchor = BuildAnchor(row, column);

            var shape = new XElement(v + "shape",
                new XAttribute("id", $"_x0000_s{shapeId}"),
                new XAttribute("type", "#_x0000_t202"),
                new XAttribute("style", "position:absolute;margin-left:0pt;margin-top:0pt;width:108pt;height:59pt;z-index:1;visibility:hidden"),
                new XAttribute("fillcolor", "#ffffe1"),
                new XAttribute(o + "insetmode", "auto"),
                new XElement(v + "fill", new XAttribute("color2", "#ffffe1")),
                new XElement(v + "shadow", new XAttribute("on", "t"), new XAttribute("color", "black"), new XAttribute("obscured", "t")),
                new XElement(v + "path", new XAttribute(o + "connecttype", "none")),
                new XElement(v + "textbox", new XAttribute("style", "mso-direction-alt:auto"),
                    new XElement("div", new XAttribute("style", "text-align:left"))),
                new XElement(x + "ClientData",
                    new XAttribute("ObjectType", "Note"),
                    new XElement(x + "MoveWithCells"),
                    new XElement(x + "SizeWithCells"),
                    new XElement(x + "Anchor", anchor),
                    new XElement(x + "AutoFill", "False"),
                    new XElement(x + "Row", (row - 1).ToString(CultureInfo.InvariantCulture)),
                    new XElement(x + "Column", (column - 1).ToString(CultureInfo.InvariantCulture))
                )
            );

            root.Add(shape);
            SaveVmlDocument(vmlPart, doc);
        }

        private bool RemoveCommentVmlShape(int row, int column) {
            var vmlPart = TryGetCommentVmlPart();
            if (vmlPart == null) return false;

            var doc = LoadOrCreateVmlDocument(vmlPart);
            var root = doc.Root;
            if (root == null) return false;

            bool removed = RemoveVmlShape(root, row, column);
            if (removed) {
                SaveVmlDocument(vmlPart, doc);
            }

            return removed;
        }

        private static bool RemoveVmlShape(XElement root, int row, int column) {
            var v = XNamespace.Get("urn:schemas-microsoft-com:vml");
            var x = XNamespace.Get("urn:schemas-microsoft-com:office:excel");
            string rowText = (row - 1).ToString(CultureInfo.InvariantCulture);
            string colText = (column - 1).ToString(CultureInfo.InvariantCulture);

            if (!VmlShapesContainCell(root, v, x, rowText, colText)) {
                return false;
            }

            var shapes = root.Elements(v + "shape").ToList();
            bool removed = false;
            foreach (var shape in shapes) {
                if (VmlShapeMatchesCell(shape, x, rowText, colText)) {
                    shape.Remove();
                    removed = true;
                }
            }
            return removed;
        }

        private static bool VmlShapesContainCell(XElement root, XNamespace v, XNamespace x, string rowText, string colText) {
            foreach (var shape in root.Elements(v + "shape")) {
                if (VmlShapeMatchesCell(shape, x, rowText, colText)) {
                    return true;
                }
            }

            return false;
        }

        private static bool VmlShapeMatchesCell(XElement shape, XNamespace x, string rowText, string colText) {
            var clientData = shape.Element(x + "ClientData");
            if (clientData == null) return false;
            var rowEl = clientData.Element(x + "Row");
            var colEl = clientData.Element(x + "Column");
            return rowEl != null
                && colEl != null
                && string.Equals(rowEl.Value?.Trim(), rowText, StringComparison.OrdinalIgnoreCase)
                && string.Equals(colEl.Value?.Trim(), colText, StringComparison.OrdinalIgnoreCase);
        }

        private bool RemoveCommentVmlShapesInRange(int firstRow, int firstColumn, int lastRow, int lastColumn) {
            var vmlPart = TryGetCommentVmlPart();
            if (vmlPart == null) return false;

            var doc = LoadOrCreateVmlDocument(vmlPart);
            var root = doc.Root;
            if (root == null) return false;

            if (RemoveVmlShapesInRange(root, firstRow, firstColumn, lastRow, lastColumn)) {
                SaveVmlDocument(vmlPart, doc);
                return true;
            }

            return false;
        }

        private static bool RemoveVmlShapesInRange(XElement root, int firstRow, int firstColumn, int lastRow, int lastColumn) {
            var v = XNamespace.Get("urn:schemas-microsoft-com:vml");
            var x = XNamespace.Get("urn:schemas-microsoft-com:office:excel");
            int firstZeroBasedRow = firstRow - 1;
            int lastZeroBasedRow = lastRow - 1;
            int firstZeroBasedColumn = firstColumn - 1;
            int lastZeroBasedColumn = lastColumn - 1;
            bool removed = false;

            if (!VmlShapesOverlapRange(root, firstZeroBasedRow, firstZeroBasedColumn, lastZeroBasedRow, lastZeroBasedColumn)) {
                return false;
            }

            foreach (var shape in root.Elements(v + "shape").ToList()) {
                var clientData = shape.Element(x + "ClientData");
                if (clientData == null) continue;

                if (!TryParseVmlCoordinate(clientData.Element(x + "Row")?.Value, out int row)
                    || !TryParseVmlCoordinate(clientData.Element(x + "Column")?.Value, out int column)) {
                    continue;
                }

                if (row >= firstZeroBasedRow
                    && row <= lastZeroBasedRow
                    && column >= firstZeroBasedColumn
                    && column <= lastZeroBasedColumn) {
                    shape.Remove();
                    removed = true;
                }
            }

            return removed;
        }

        private static bool VmlShapesOverlapRange(XElement root, int firstZeroBasedRow, int firstZeroBasedColumn, int lastZeroBasedRow, int lastZeroBasedColumn) {
            var v = XNamespace.Get("urn:schemas-microsoft-com:vml");
            var x = XNamespace.Get("urn:schemas-microsoft-com:office:excel");

            foreach (var shape in root.Elements(v + "shape")) {
                var clientData = shape.Element(x + "ClientData");
                if (clientData == null) continue;

                if (!TryParseVmlCoordinate(clientData.Element(x + "Row")?.Value, out int row)
                    || !TryParseVmlCoordinate(clientData.Element(x + "Column")?.Value, out int column)) {
                    continue;
                }

                if (row >= firstZeroBasedRow
                    && row <= lastZeroBasedRow
                    && column >= firstZeroBasedColumn
                    && column <= lastZeroBasedColumn) {
                    return true;
                }
            }

            return false;
        }

        private static bool TryParseVmlCoordinate(string? text, out int value) {
            value = 0;
            return int.TryParse(text?.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out value);
        }

        internal bool CleanupCommentArtifacts() {
            var ws = WorksheetRoot;
            var commentsPart = WorksheetCommentsPartRoot;
            bool hasComments = commentsPart?.Comments?.CommentList?.Elements<Comment>().Any() is true;
            bool changed = false;

            if (!hasComments && commentsPart != null) {
                _worksheetPart.DeletePart(commentsPart);
                changed = true;
            }

            var legacy = ws.GetFirstChild<LegacyDrawing>();
            if (legacy?.Id?.Value is not string legacyRelId || string.IsNullOrWhiteSpace(legacyRelId)) {
                return changed;
            }

            OpenXmlPart? legacyPart = null;
            try {
                legacyPart = _worksheetPart.GetPartById(legacyRelId);
            } catch {
                ws.RemoveChild(legacy);
                return true;
            }

            if (!hasComments && legacyPart is VmlDrawingPart vmlPart) {
                _worksheetPart.DeletePart(vmlPart);
                ws.RemoveChild(legacy);
                changed = true;
            }

            return changed;
        }

        private VmlDrawingPart GetOrCreateCommentVmlPart() {
            var ws = WorksheetRoot;
            var legacy = ws.GetFirstChild<LegacyDrawing>();
            if (legacy?.Id?.Value is string legacyRelId && !string.IsNullOrWhiteSpace(legacyRelId)) {
                return (VmlDrawingPart)_worksheetPart.GetPartById(legacyRelId);
            }

            var vmlPart = _worksheetPart.AddNewPart<VmlDrawingPart>();
            string relId = _worksheetPart.GetIdOfPart(vmlPart);
            legacy = new LegacyDrawing { Id = relId };

            var legacyHeaderFooter = ws.GetFirstChild<LegacyDrawingHeaderFooter>();
            if (legacyHeaderFooter != null) {
                ws.InsertBefore(legacy, legacyHeaderFooter);
            } else {
                ws.Append(legacy);
            }

            return vmlPart;
        }

        private VmlDrawingPart? TryGetCommentVmlPart() {
            var legacy = WorksheetRoot.GetFirstChild<LegacyDrawing>();
            if (legacy?.Id?.Value is string legacyRelId && !string.IsNullOrWhiteSpace(legacyRelId)) {
                return (VmlDrawingPart)_worksheetPart.GetPartById(legacyRelId);
            }
            return null;
        }

        private static XDocument LoadOrCreateVmlDocument(VmlDrawingPart part) {
            try {
                using var stream = part.GetStream();
                if (stream.Length > 0) {
                    return XDocument.Load(stream);
                }
            } catch (Exception ex) {
                System.Diagnostics.Debug.WriteLine($"Failed to load VML drawing part stream: {ex}");
            }

            var v = XNamespace.Get("urn:schemas-microsoft-com:vml");
            var o = XNamespace.Get("urn:schemas-microsoft-com:office:office");
            var x = XNamespace.Get("urn:schemas-microsoft-com:office:excel");

            var root = new XElement("xml",
                new XAttribute(XNamespace.Xmlns + "v", v),
                new XAttribute(XNamespace.Xmlns + "o", o),
                new XAttribute(XNamespace.Xmlns + "x", x),
                new XElement(o + "shapelayout",
                    new XAttribute(v + "ext", "edit"),
                    new XElement(o + "idmap", new XAttribute(v + "ext", "edit"), new XAttribute("data", "1"))
                ),
                new XElement(v + "shapetype",
                    new XAttribute("id", "_x0000_t202"),
                    new XAttribute("coordsize", "21600,21600"),
                    new XAttribute(o + "spt", "202"),
                    new XAttribute("path", "m,l,21600r21600,l21600,xe"),
                    new XElement(v + "stroke", new XAttribute("joinstyle", "miter")),
                    new XElement(v + "path", new XAttribute("gradientshapeok", "t"), new XAttribute(o + "connecttype", "rect"))
                )
            );

            return new XDocument(root);
        }

        private static void SaveVmlDocument(VmlDrawingPart part, XDocument doc) {
            using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
            doc.Save(stream);
        }

        private static int NextVmlShapeId(XElement root) {
            var v = XNamespace.Get("urn:schemas-microsoft-com:vml");
            int max = 1024;
            foreach (var shape in root.Elements(v + "shape")) {
                var idAttr = shape.Attribute("id")?.Value;
                if (idAttr == null || idAttr.Length == 0) {
                    continue;
                }
                if (!idAttr.StartsWith("_x0000_s", StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                var numPart = idAttr.Substring("_x0000_s".Length);
                if (int.TryParse(numPart, NumberStyles.Integer, CultureInfo.InvariantCulture, out int id)) {
                    if (id > max) max = id;
                }
            }
            return max + 1;
        }

        private static string BuildAnchor(int row, int column) {
            int col1 = Math.Max(0, column - 1);
            int row1 = Math.Max(0, row - 1);
            int col2 = col1 + 2;
            int row2 = row1 + 3;
            return string.Format(CultureInfo.InvariantCulture, "{0}, 15, {1}, 2, {2}, 15, {3}, 4", col1, row1, col2, row2);
        }
    }
}
