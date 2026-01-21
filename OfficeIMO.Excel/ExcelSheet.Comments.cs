using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Xml.Linq;

namespace OfficeIMO.Excel {
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

            WriteLock(() => {
                string reference = A1.ColumnIndexToLetters(column) + row.ToString(CultureInfo.InvariantCulture);
                string authorDisplay = NormalizeAuthor(author, initials);

                var commentsPart = GetOrCreateCommentsPart();
                var comments = commentsPart.Comments ??= new Comments();
                comments.Authors ??= new Authors();
                comments.CommentList ??= new CommentList();

                uint authorId = EnsureAuthorId(comments.Authors, authorDisplay);
                RemoveCommentInternal(comments.CommentList, reference);

                var comment = new Comment { Reference = reference, AuthorId = authorId };
                comment.Append(BuildCommentText(text));
                comments.CommentList.Append(comment);
                comments.Save();

                EnsureCommentVmlShape(row, column);
                _worksheetPart.Worksheet.Save();
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
        /// Removes a comment from the specified cell (if present).
        /// </summary>
        /// <param name="row">1-based row index.</param>
        /// <param name="column">1-based column index.</param>
        public void ClearComment(int row, int column) {
            if (row <= 0) throw new ArgumentOutOfRangeException(nameof(row), "Row and column are 1-based and must be positive.");
            if (column <= 0) throw new ArgumentOutOfRangeException(nameof(column), "Row and column are 1-based and must be positive.");

            WriteLock(() => {
                string reference = A1.ColumnIndexToLetters(column) + row.ToString(CultureInfo.InvariantCulture);

                var commentsPart = _worksheetPart.WorksheetCommentsPart;
                if (commentsPart?.Comments?.CommentList == null) {
                    RemoveCommentVmlShape(row, column);
                    return;
                }

                RemoveCommentInternal(commentsPart.Comments.CommentList, reference);
                commentsPart.Comments.Save();
                RemoveCommentVmlShape(row, column);
                _worksheetPart.Worksheet.Save();
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
            string reference = A1.ColumnIndexToLetters(column) + row.ToString(CultureInfo.InvariantCulture);
            var commentsPart = _worksheetPart.WorksheetCommentsPart;
            return commentsPart?.Comments?.CommentList?
                .Elements<Comment>()
                .Any(c => string.Equals(c.Reference?.Value, reference, StringComparison.OrdinalIgnoreCase)) is true;
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

        private static void RemoveCommentInternal(CommentList list, string reference) {
            var existing = list.Elements<Comment>()
                .FirstOrDefault(c => string.Equals(c.Reference?.Value, reference, StringComparison.OrdinalIgnoreCase));
            existing?.Remove();
        }

        private static CommentText BuildCommentText(string text) {
            var commentText = new CommentText();
            var run = new Run();
            var lines = text.Replace("\r\n", "\n").Split('\n');
            for (int i = 0; i < lines.Length; i++) {
                if (i > 0) {
                    run.Append(new Break());
                }
                run.Append(new Text(lines[i]) { Space = SpaceProcessingModeValues.Preserve });
            }
            commentText.Append(run);
            return commentText;
        }

        private WorksheetCommentsPart GetOrCreateCommentsPart() {
            var part = _worksheetPart.WorksheetCommentsPart;
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

        private void RemoveCommentVmlShape(int row, int column) {
            var vmlPart = TryGetCommentVmlPart();
            if (vmlPart == null) return;

            var doc = LoadOrCreateVmlDocument(vmlPart);
            var root = doc.Root;
            if (root == null) return;

            bool removed = RemoveVmlShape(root, row, column);
            if (removed) {
                SaveVmlDocument(vmlPart, doc);
            }
        }

        private static bool RemoveVmlShape(XElement root, int row, int column) {
            var v = XNamespace.Get("urn:schemas-microsoft-com:vml");
            var x = XNamespace.Get("urn:schemas-microsoft-com:office:excel");
            string rowText = (row - 1).ToString(CultureInfo.InvariantCulture);
            string colText = (column - 1).ToString(CultureInfo.InvariantCulture);

            var shapes = root.Elements(v + "shape").ToList();
            bool removed = false;
            foreach (var shape in shapes) {
                var clientData = shape.Element(x + "ClientData");
                if (clientData == null) continue;
                var rowEl = clientData.Element(x + "Row");
                var colEl = clientData.Element(x + "Column");
                if (rowEl != null && colEl != null
                    && string.Equals(rowEl.Value?.Trim(), rowText, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(colEl.Value?.Trim(), colText, StringComparison.OrdinalIgnoreCase)) {
                    shape.Remove();
                    removed = true;
                }
            }
            return removed;
        }

        private VmlDrawingPart GetOrCreateCommentVmlPart() {
            var ws = _worksheetPart.Worksheet;
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
            var legacy = _worksheetPart.Worksheet.GetFirstChild<LegacyDrawing>();
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
