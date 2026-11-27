using System.Globalization;
using System.Text;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Text written to a table's header row in a workbook that has not opened its SDK package. The column takes the
    /// text as its name, as the model renames it: every header cell of the table is rewritten as its column's name,
    /// and a renamed column drops the calculation chain and asks for a full calculation on open. A rename that formulas
    /// could refer to, a table feeding a query table, or a header cell holding a formula opens the SDK instead, which
    /// rewrites those references and fields.
    /// </summary>
    public partial class ExcelDocument {
        private const string DrawingRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "drawing";
        private const string ChartRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "chart";
        private const string ExtendedChartRelationshipType = "http://schemas.microsoft.com/office/2014/relationships/chartEx";
        private const string PivotTablePartRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "pivotTable";
        private const string ChartsheetRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "chartsheet";
        // Structured references escape these, and the package escapes some of them, so a name holding one is not
        // searched for as written.
        private const string EscapedNameCharacters = "[]#'&<>\"";

        /// <summary>
        /// Records text written to a table header cell as the model would apply it; false, having changed nothing, when
        /// the model has to apply it.
        /// </summary>
        /// <exception cref="ArgumentException">The table's column names would be empty, too long or duplicated.</exception>
        private bool TryRecordDeferredTableHeader(ExcelSheet sheet, ExcelDeferredSheet entry, in ExcelDeferredCellEdit edit) {
            ReaderWriterLockSlim lck = EnsureLock();
            // A reader cannot take the write lock the package rewrite needs.
            if (!Locking.IsNoLock && (lck.IsReadLockHeld || lck.IsUpgradeableReadLockHeld) && !lck.IsWriteLockHeld) return false;
            ExcelDeferredCellEdit written = edit;
            return Locking.ExecuteWrite(lck, () => RecordDeferredTableHeaderLocked(sheet, entry, written));
        }

        private bool RecordDeferredTableHeaderLocked(ExcelSheet sheet, ExcelDeferredSheet entry, ExcelDeferredCellEdit edit) {
            ExcelDeferredWorkbook? workbook = _deferredWorkbook;
            byte[]? packageBytes = _deferredPackageBytes;
            if (workbook == null || packageBytes == null || !ReferenceEquals(sheet.DeferredEntry, entry)) return false;
            if (WorkbookHasCellMetadata(workbook, packageBytes)) return false;
            if (!TryFindDeferredHeaderTable(entry, edit.Row, edit.Column, out ExcelTablePart table, out int left, out int index)) return false;
            if (table.HasQueryTable || index >= table.Columns.Count) return false;

            string[] current = new string[table.Columns.Count];
            for (int column = 0; column < current.Length; column++) current[column] = table.Columns[column].Name ?? string.Empty;
            string[] requested = (string[])current.Clone();
            requested[index] = edit.Text;
            string[] names = ExcelSheet.ValidateTableColumnNames(requested);

            var renamed = new List<string>();
            for (int column = 0; column < names.Length; column++) {
                if (!string.Equals(current[column], names[column], StringComparison.Ordinal)) renamed.Add(current[column]);
            }
            foreach (string oldName in renamed) {
                if (oldName.Length > 0 && MayBeReferencedByFormulas(workbook, packageBytes, oldName)) return false;
            }

            // Each header cell becomes its column's name unless it already holds that text. Writing a shared string's
            // own text again leaves the cell as it was, so only text the cell holds inline needs reading.
            HeaderCellKind[]? kinds = ReadHeaderCellKinds(packageBytes, entry.PartName, edit.Row, left, names.Length);
            if (kinds == null || Array.IndexOf(kinds, HeaderCellKind.Formula) >= 0) return false;
            var headerEdits = new List<ExcelDeferredCellEdit>(names.Length);
            ExcelFormulaSheetModel? model = null;
            for (int column = 0; column < names.Length; column++) {
                if (kinds[column] == HeaderCellKind.InlineText) {
                    model ??= GetFormulaSheetModel(sheet);
                    if (!ReferenceEquals(sheet.DeferredEntry, entry) || model is not ExcelXmlFormulaSheetModel) return false;
                    ExcelFormulaCellRecord? cell = model.Find(edit.Row, left + column);
                    if (cell != null && cell.HasFormula) return false;
                    if (cell != null && cell.ValueType is ExcelCellValueType.SharedString or ExcelCellValueType.InlineString or ExcelCellValueType.String
                        && model.ReadValue(edit.Row, left + column).Value is string text && string.Equals(text, names[column], StringComparison.Ordinal)) continue;
                }
                headerEdits.Add(new ExcelDeferredCellEdit {
                    Row = edit.Row, Column = checked((ushort)(left + column)), Kind = ExcelDeferredCellKind.Text, Payload = names[column]
                });
            }

            if (renamed.Count > 0 && !TryRenameDeferredTableColumns(workbook, table, names)) return false;
            if (headerEdits.Count > 0) {
                lock (_deferredEditsGate) {
                    ExcelDeferredCellStore? store = entry.Cells;
                    if (store == null) {
                        entry.Cells = store = new ExcelDeferredCellStore();
                        _deferredEditedSheets.Add(sheet);
                    }
                    foreach (ExcelDeferredCellEdit header in headerEdits) store.Add(header);
                }
            }
            if (renamed.Count > 0 || headerEdits.Count > 0) {
                if (!_packageDirty) MarkPackageDirty();
                MarkFormulaInputMutation();
            }
            return true;
        }

        private enum HeaderCellKind : byte { Other, InlineText, Formula }

        /// <summary>
        /// What the package holds in the header cells: a formula, text held in the cell rather than as a shared string,
        /// or anything else. Null when a row or cell of the worksheet has no reference, which the model resolves.
        /// </summary>
        private static HeaderCellKind[]? ReadHeaderCellKinds(byte[] packageBytes, string partName, int row, int left, int count) {
            var kinds = new HeaderCellKind[count];
            using var source = new MemoryStream(packageBytes, writable: false);
            using OfficeOpcPackage package = OfficeOpcPackage.Open(source, leaveOpen: true);
            OfficeXmlPullReader reader = package.OpenXml(partName, out OfficePooledBytes content, processMarkupCompatibility: false);
            using (content) {
                bool inRow = false;
                int cellIndex = -1;
                while (reader.Read()) {
                    if (reader.NodeKind != OfficeXmlNodeKind.StartElement || reader.NamespaceId != OfficeXmlNamespaces.SpreadsheetML) continue;
                    // worksheet / sheetData / row / c / f
                    if (reader.Depth == 3 && reader.LocalNameEquals("row")) {
                        if (!reader.TryGetAttribute(OfficeXmlNamespaces.None, "r", out string number)
                            || !int.TryParse(number, NumberStyles.None, CultureInfo.InvariantCulture, out int index)) return null;
                        if (index > row) break;
                        inRow = index == row;
                        cellIndex = -1;
                    } else if (inRow && reader.Depth == 4 && reader.LocalNameEquals("c")) {
                        if (!reader.TryGetAttribute(OfficeXmlNamespaces.None, "r", out string reference)
                            || !ExcelA1.TryParseCellReferenceFast(reference.AsSpan(), out int cellRow, out int column) || cellRow != row) return null;
                        cellIndex = column >= left && column < left + count ? column - left : -1;
                        if (cellIndex >= 0 && reader.TryGetAttribute(OfficeXmlNamespaces.None, "t", out string type) && type is "inlineStr" or "str") {
                            kinds[cellIndex] = HeaderCellKind.InlineText;
                        }
                    } else if (inRow && cellIndex >= 0 && reader.Depth == 5 && reader.LocalNameEquals("f")) {
                        kinds[cellIndex] = HeaderCellKind.Formula;
                    }
                }
            }
            return kinds;
        }

        /// <summary>The first table, in relationship order, whose header row holds the cell, as the model finds it.</summary>
        private bool TryFindDeferredHeaderTable(ExcelDeferredSheet entry, int row, int column, out ExcelTablePart found, out int left, out int index) {
            found = null!;
            left = index = 0;
            List<ExcelTablePart>? tables = TryGetDeferredTableParts(entry);
            if (tables == null) return false;
            foreach (ExcelTablePart table in tables) {
                uint headerRows;
                try {
                    headerRows = table.HeaderRowCount ?? 1U;
                } catch (FormatException) {
                    // The model fails on the invalid count; let it report that.
                    return false;
                }
                if (headerRows == 0U || !ExcelA1.TryParseRange(table.Reference ?? string.Empty, out int top, out int first, out _, out int last)
                    || row != top || column < first || column > last) continue;
                found = table;
                left = first;
                index = column - first;
                return true;
            }
            return false;
        }

        /// <summary>
        /// Whether a formula the model rewrites on a column rename could name <paramref name="columnName"/>: a formula
        /// in the workbook, a worksheet, a table, a chart or a PivotTable that contains it. Names a structured reference
        /// or the package escapes always count.
        /// </summary>
        private static bool MayBeReferencedByFormulas(ExcelDeferredWorkbook workbook, byte[] packageBytes, string columnName) {
            foreach (char character in columnName) {
                if (EscapedNameCharacters.IndexOf(character) >= 0) return true;
            }
            using var source = new MemoryStream(packageBytes, writable: false);
            using OfficeOpcPackage package = OfficeOpcPackage.Open(source, leaveOpen: true);
            var visited = new HashSet<string>(StringComparer.Ordinal);
            bool Contains(string partName) => visited.Add(partName) && package.PartExists(partName) && FormulaLeavesContain(package, partName, columnName);
            bool DrawingContains(string sheetPartName) {
                foreach (OfficeOpcRelationship drawing in Related(package, sheetPartName, DrawingRelationshipType)) {
                    foreach (OfficeOpcRelationship chart in package.GetRelationships(drawing.TargetPartName)) {
                        if (chart.IsExternal || chart.TargetPartName is not { } target) continue;
                        if ((string.Equals(chart.Type, ChartRelationshipType, StringComparison.Ordinal)
                            || string.Equals(chart.Type, ExtendedChartRelationshipType, StringComparison.Ordinal)) && Contains(target)) return true;
                    }
                }
                return false;
            }

            if (Contains(workbook.WorkbookPartName)) return true;
            foreach (OfficeOpcRelationship worksheet in Related(package, workbook.WorkbookPartName, WorksheetRelationshipType)) {
                string sheet = worksheet.TargetPartName!;
                if (Contains(sheet)) return true;
                foreach (OfficeOpcRelationship table in Related(package, sheet, ExcelTablePart.TableRelationshipType)) {
                    if (Contains(table.TargetPartName!)) return true;
                }
                if (DrawingContains(sheet)) return true;
                foreach (OfficeOpcRelationship pivot in Related(package, sheet, PivotTablePartRelationshipType)) {
                    if (Contains(pivot.TargetPartName!)) return true;
                }
            }
            foreach (OfficeOpcRelationship chartsheet in Related(package, workbook.WorkbookPartName, ChartsheetRelationshipType)) {
                if (DrawingContains(chartsheet.TargetPartName!)) return true;
            }
            return false;
        }

        private static IEnumerable<OfficeOpcRelationship> Related(OfficeOpcPackage package, string? sourcePartName, string type) {
            if (sourcePartName == null) yield break;
            foreach (OfficeOpcRelationship relationship in package.GetRelationships(sourcePartName)) {
                if (!relationship.IsExternal && relationship.TargetPartName != null
                    && string.Equals(relationship.Type, type, StringComparison.Ordinal)) yield return relationship;
            }
        }

        /// <summary>
        /// Whether a formula element of the part (any element named <c>f</c> or <c>definedName</c>, or whose name
        /// contains "formula", at any depth, as the model rewrites them) contains the text, ignoring case.
        /// </summary>
        private static bool FormulaLeavesContain(OfficeOpcPackage package, string partName, string text) {
            using (OfficePooledBytes raw = package.ReadPart(partName)) {
                // Most parts never mention the name, which a search of the part settles without parsing it.
                if (!MayContain(raw.Array, raw.Length, text)) return false;
            }
            OfficeXmlPullReader reader = package.OpenXml(partName, out OfficePooledBytes content, processMarkupCompatibility: false);
            using (content) {
                while (reader.Read()) {
                    if (reader.NodeKind != OfficeXmlNodeKind.StartElement || reader.IsEmptyElement) continue;
                    string name = reader.LocalName;
                    if (!string.Equals(name, "f", StringComparison.OrdinalIgnoreCase)
                        && !string.Equals(name, "definedName", StringComparison.OrdinalIgnoreCase)
                        && name.IndexOf("formula", StringComparison.OrdinalIgnoreCase) < 0) continue;
                    int depth = reader.Depth;
                    var formula = new StringBuilder();
                    while (reader.Read()) {
                        if (reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == depth) break;
                        if (reader.NodeKind == OfficeXmlNodeKind.Text) formula.Append(reader.GetText());
                    }
                    if (formula.ToString().IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0) return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Whether the UTF-8 bytes may contain the text, ignoring case: an ASCII name is searched for in the bytes, any
        /// other in the decoded text, and markup holding a character reference always may.
        /// </summary>
        private static bool MayContain(byte[] bytes, int length, string text) {
            // A character reference can spell any part of the name, which only the parsed text shows.
            for (int index = 0; index + 1 < length; index++) {
                if (bytes[index] == (byte)'&' && bytes[index + 1] == (byte)'#') return true;
            }
            foreach (char character in text) {
                if (character > 0x7F) return Encoding.UTF8.GetString(bytes, 0, length).IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0;
            }
            int last = length - text.Length;
            byte first = AsciiLower((byte)text[0]);
            for (int start = 0; start <= last; start++) {
                if (AsciiLower(bytes[start]) != first) continue;
                int matched = 1;
                while (matched < text.Length && AsciiLower(bytes[start + matched]) == AsciiLower((byte)text[matched])) matched++;
                if (matched == text.Length) return true;
            }
            return false;
        }

        private static byte AsciiLower(byte value) => value is >= (byte)'A' and <= (byte)'Z' ? (byte)(value | 0x20) : value;

        /// <summary>
        /// Writes the column names into the table part and, as the model does after a rename, drops the calculation
        /// chain and asks for a full calculation on open when a worksheet has formulas.
        /// </summary>
        private bool TryRenameDeferredTableColumns(ExcelDeferredWorkbook workbook, ExcelTablePart table, string[] names) {
            byte[] bytes;
            using (var input = new MemoryStream(_deferredPackageBytes!, writable: false))
            using (OfficeOpcEditablePackage package = OfficeOpcEditablePackage.Open(input, leaveOpen: true)) {
                OfficeXmlTree tree = package.GetXml(table.PartName);
                int columns = tree.FindChild(tree.Root, OfficeXmlNamespaces.SpreadsheetML, "tableColumns");
                if (columns == OfficeXmlTree.NoNode) return false;
                int index = 0;
                for (int column = tree.FindChild(columns, OfficeXmlNamespaces.SpreadsheetML, "tableColumn");
                    column != OfficeXmlTree.NoNode && index < names.Length;
                    column = tree.FindNextSibling(column, OfficeXmlNamespaces.SpreadsheetML, "tableColumn"), index++) {
                    if (!tree.AttributeEquals(column, OfficeXmlNamespaces.None, "name", names[index])) {
                        tree.SetAttribute(column, OfficeXmlNamespaces.None, "name", names[index]);
                    }
                }
                if (index != names.Length) return false;
                tree.SetAttribute(columns, OfficeXmlNamespaces.None, "count", names.Length.ToString(CultureInfo.InvariantCulture));

                RemoveCalculationChain(package, workbook.WorkbookPartName);
                if (AnyWorksheetHasCellFormulas(package, workbook.WorkbookPartName)) RequestFullCalculationOnOpen(package.GetXml(workbook.WorkbookPartName));

                using var output = new MemoryStream(_deferredPackageBytes!.Length + 1024);
                package.Save(output);
                bytes = output.ToArray();
            }
            _deferredPackageBytes = bytes;
            _deferredPackageWrittenHere = true;
            ForgetDeferredTables();
            _xmlFormulaWorkbook = null;
            return true;
        }

        /// <summary>Whether a cell of any worksheet the workbook relates has a formula.</summary>
        private static bool AnyWorksheetHasCellFormulas(OfficeOpcEditablePackage package, string workbookPartName) {
            OfficeXmlTree relationships = package.GetRelationships(workbookPartName);
            for (int relationship = relationships.FirstChildOf(relationships.Root); relationship != OfficeXmlTree.NoNode;
                relationship = relationships.NextSiblingOf(relationship)) {
                if (!relationships.IsElement(relationship, OfficeXmlNamespaces.PackageRelationships, "Relationship")
                    || !relationships.AttributeEquals(relationship, OfficeXmlNamespaces.None, "Type", WorksheetRelationshipType)
                    || relationships.AttributeEquals(relationship, OfficeXmlNamespaces.None, "TargetMode", "External")
                    || relationships.GetAttribute(relationship, OfficeXmlNamespaces.None, "Id") is not string id
                    || package.GetRelationshipTarget(workbookPartName, id) is not string target || !package.PartExists(target)) continue;
                if (!package.TryReadUnparsed(target, out OfficeXmlPullReader reader, out OfficePooledBytes content)) {
                    if (TreeHasCellFormulas(package.GetXml(target))) return true;
                    continue;
                }
                using (content) {
                    while (reader.Read()) {
                        // worksheet / sheetData / row / c / f
                        if (reader.NodeKind == OfficeXmlNodeKind.StartElement && reader.Depth == 5
                            && reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "f")) return true;
                    }
                }
            }
            return false;
        }

        private static bool TreeHasCellFormulas(OfficeXmlTree tree) {
            for (int data = tree.FindChild(tree.Root, OfficeXmlNamespaces.SpreadsheetML, "sheetData"); data != OfficeXmlTree.NoNode;
                data = tree.FindNextSibling(data, OfficeXmlNamespaces.SpreadsheetML, "sheetData")) {
                for (int row = tree.FindChild(data, OfficeXmlNamespaces.SpreadsheetML, "row"); row != OfficeXmlTree.NoNode;
                    row = tree.FindNextSibling(row, OfficeXmlNamespaces.SpreadsheetML, "row")) {
                    for (int cell = tree.FindChild(row, OfficeXmlNamespaces.SpreadsheetML, "c"); cell != OfficeXmlTree.NoNode;
                        cell = tree.FindNextSibling(cell, OfficeXmlNamespaces.SpreadsheetML, "c")) {
                        if (tree.FindChild(cell, OfficeXmlNamespaces.SpreadsheetML, "f") != OfficeXmlTree.NoNode) return true;
                    }
                }
            }
            return false;
        }

        /// <summary>Sets the calculation properties the model sets when a table schema change asks for a full calculation.</summary>
        private static void RequestFullCalculationOnOpen(OfficeXmlTree workbook) {
            int properties = workbook.FindChild(workbook.Root, OfficeXmlNamespaces.SpreadsheetML, "calcPr");
            if (properties == OfficeXmlTree.NoNode) {
                properties = workbook.CreateElement(OfficeXmlNamespaces.SpreadsheetML, "calcPr");
                int later = OfficeXmlTree.NoNode;
                for (int child = workbook.FirstChildOf(workbook.Root); child != OfficeXmlTree.NoNode; child = workbook.NextSiblingOf(child)) {
                    if (workbook.KindOf(child) == OfficeXmlTreeNodeKind.Element && workbook.LocalNameOf(child) is "oleSize" or "customWorkbookViews"
                        or "pivotCaches" or "smartTagPr" or "smartTagTypes" or "webPublishing" or "fileRecoveryPr" or "webPublishObjects" or "extLst") {
                        later = child;
                        break;
                    }
                }
                workbook.InsertBefore(workbook.Root, properties, later);
            }
            workbook.SetAttribute(properties, OfficeXmlNamespaces.None, "calcId", "191029");
            workbook.SetAttribute(properties, OfficeXmlNamespaces.None, "fullCalcOnLoad", "1");
            workbook.SetAttribute(properties, OfficeXmlNamespaces.None, "forceFullCalc", "1");
        }
    }
}
