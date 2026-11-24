using DocumentFormat.OpenXml.Packaging;
using System.Globalization;
using System.Xml.Linq;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Loads the note drawing for an edit that may change any of its shapes. The cell index kept for
        /// <see cref="SetCommentInternal"/> is dropped, because it no longer describes the tree once such an edit runs.
        /// </summary>
        private XDocument LoadOrCreateVmlDocument(VmlDrawingPart part) {
            Persistence.CommentVmlShapes = null;
            return LoadCommentVmlDocument(part);
        }

        private XDocument LoadCommentVmlDocument(VmlDrawingPart part) {
            if (PendingCommentVmlDocuments.TryGetValue(part, out XDocument? pending)) return pending;
            try {
                using var stream = part.GetStream();
                if (stream.Length > 0) {
                    return LoadVmlXDocument(stream);
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

        /// <summary>
        /// Keeps an edited note drawing loaded until the save, so each note costs the same however many the sheet
        /// has. Anything that reads the part stream instead calls <see cref="CommitCommentVmlChanges"/> or
        /// <see cref="CommitPendingCommentDrawings"/> first. Once the package is handed out, edits write through.
        /// </summary>
        private void SaveVmlDocument(VmlDrawingPart part, XDocument doc) {
            if (_excelDocument.IsOpenXmlDocumentExposed) {
                PendingCommentVmlDocuments.Remove(part);
                using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
                doc.Save(stream);
                return;
            }
            PendingCommentVmlDocuments[part] = doc;
            MarkRequiresSavePreparation();
        }

        /// <summary>Flushes kept note drawings before a native writer or stream-based planner inspects them.</summary>
        internal void CommitCommentVmlChanges() {
            if (PendingCommentVmlDocuments.Count == 0) return;
            CommitCommentVmlDocuments(_worksheetPart, Persistence);
        }

        /// <summary>Flushes the kept note drawings of every worksheet before something reads part streams across the workbook.</summary>
        internal static void CommitPendingCommentDrawings(WorkbookPart workbookPart) {
            foreach (WorksheetPart worksheetPart in workbookPart.WorksheetParts) {
                if (WorksheetPersistenceStates.TryGetValue(worksheetPart, out WorksheetPersistenceState? state)
                    && state.CommentVmlDocuments.Count > 0) {
                    CommitCommentVmlDocuments(worksheetPart, state);
                }
            }
        }

        private static void CommitCommentVmlDocuments(WorksheetPart worksheetPart, WorksheetPersistenceState state) {
            foreach (var pair in state.CommentVmlDocuments) {
                if (!worksheetPart.Parts.Any(item => ReferenceEquals(item.OpenXmlPart, pair.Key))) continue;
                using var stream = pair.Key.GetStream(FileMode.Create, FileAccess.Write);
                pair.Value.Save(stream);
            }
            state.CommentVmlDocuments.Clear();
            state.CommentVmlShapes = null;
        }

        /// <summary>Gets the note-shape index of a loaded drawing, building it with one pass when the drawing changed.</summary>
        private CommentVmlShapeIndex GetCommentVmlShapeIndex(XDocument document, XElement root) {
            CommentVmlShapeIndex? index = Persistence.CommentVmlShapes;
            if (index == null || !ReferenceEquals(index.Document, document)) {
                index = new CommentVmlShapeIndex(document, root);
                Persistence.CommentVmlShapes = index;
            }
            return index;
        }

        /// <summary>
        /// The direct <c>v:shape</c> children of a note drawing by the cell their client data names, and the shape ids in use,
        /// so replacing a note and choosing the next id don't walk every shape. It matches cells and ids as
        /// <see cref="RemoveVmlShape"/> and <see cref="NextVmlShapeId"/> do.
        /// </summary>
        private sealed class CommentVmlShapeIndex {
            private static readonly XNamespace V = "urn:schemas-microsoft-com:vml";
            private static readonly XNamespace X = "urn:schemas-microsoft-com:office:excel";
            private readonly Dictionary<(string Row, string Column), List<XElement>> _shapesByCell = new(CellKeyComparer.Instance);
            private readonly Dictionary<int, int> _idCounts = new();
            private int _maxId = MinimumShapeId;
            private const int MinimumShapeId = 1024;

            internal CommentVmlShapeIndex(XDocument document, XElement root) {
                Document = document;
                foreach (XElement shape in root.Elements(V + "shape")) Add(shape);
            }

            internal XDocument Document { get; }

            internal int NextShapeId => _maxId + 1;

            internal void Add(XElement shape) {
                if (TryGetCellKey(shape, out (string Row, string Column) key)) {
                    if (!_shapesByCell.TryGetValue(key, out List<XElement>? shapes)) {
                        shapes = new List<XElement>(1);
                        _shapesByCell.Add(key, shapes);
                    }
                    shapes.Add(shape);
                }
                if (TryGetShapeId(shape, out int id)) {
                    _idCounts[id] = _idCounts.TryGetValue(id, out int count) ? count + 1 : 1;
                    if (id > _maxId) _maxId = id;
                }
            }

            internal void RemoveCell(int row, int column) {
                var key = ((row - 1).ToString(CultureInfo.InvariantCulture), (column - 1).ToString(CultureInfo.InvariantCulture));
                if (!_shapesByCell.TryGetValue(key, out List<XElement>? shapes)) return;
                _shapesByCell.Remove(key);
                bool recomputeMax = false;
                foreach (XElement shape in shapes) {
                    shape.Remove();
                    if (!TryGetShapeId(shape, out int id) || !_idCounts.TryGetValue(id, out int count)) continue;
                    if (count > 1) {
                        _idCounts[id] = count - 1;
                    } else {
                        _idCounts.Remove(id);
                        recomputeMax |= id == _maxId;
                    }
                }
                if (recomputeMax) {
                    _maxId = MinimumShapeId;
                    foreach (int id in _idCounts.Keys) {
                        if (id > _maxId) _maxId = id;
                    }
                }
            }

            private static bool TryGetCellKey(XElement shape, out (string Row, string Column) key) {
                key = default;
                XElement? clientData = shape.Element(X + "ClientData");
                string? row = clientData?.Element(X + "Row")?.Value;
                string? column = clientData?.Element(X + "Column")?.Value;
                if (row == null || column == null) return false;
                key = (row.Trim(), column.Trim());
                return true;
            }

            private sealed class CellKeyComparer : IEqualityComparer<(string Row, string Column)> {
                internal static readonly CellKeyComparer Instance = new();

                public bool Equals((string Row, string Column) left, (string Row, string Column) right) =>
                    StringComparer.OrdinalIgnoreCase.Equals(left.Row, right.Row)
                    && StringComparer.OrdinalIgnoreCase.Equals(left.Column, right.Column);

                public int GetHashCode((string Row, string Column) key) =>
                    StringComparer.OrdinalIgnoreCase.GetHashCode(key.Row) * 31 + StringComparer.OrdinalIgnoreCase.GetHashCode(key.Column);
            }

            private static bool TryGetShapeId(XElement shape, out int id) {
                id = 0;
                string? value = shape.Attribute("id")?.Value;
                return value != null
                    && value.StartsWith("_x0000_s", StringComparison.OrdinalIgnoreCase)
                    && int.TryParse(value.Substring("_x0000_s".Length), NumberStyles.Integer, CultureInfo.InvariantCulture, out id);
            }
        }
    }
}
