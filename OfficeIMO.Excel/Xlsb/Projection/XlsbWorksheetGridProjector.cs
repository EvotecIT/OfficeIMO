using System.Xml;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Excel.Utilities;
using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Projection {
    /// <summary>Streams the XLSB cell grid to SpreadsheetML in the validated BIFF12 row and column order.</summary>
    internal static partial class XlsbWorksheetGridProjector {
        private static readonly string Namespace = OfficeXmlNamespaces.KnownUri(OfficeXmlNamespaces.SpreadsheetML)!;

        internal static void WriteWorksheet(Stream output, XlsbSheet source, Func<string, int?> sharedStringIndex,
            Action<string?, bool> recordMutation, CancellationToken cancellationToken, OfficeXmlNode layout = default,
            XlsbStylesheet? styles = null) {
            using XmlWriter writer = XmlWriter.Create(output, new XmlWriterSettings {
                Encoding = new UTF8Encoding(false),
                NewLineHandling = NewLineHandling.Entitize,
                NamespaceHandling = NamespaceHandling.OmitDuplicates
            });
            writer.WriteStartElement("x", "worksheet", Namespace);
            if (layout.IsNone) {
                WriteGrid(writer, source, sharedStringIndex, recordMutation, includeLayout: false, styles, cancellationToken);
            } else {
                foreach (OfficeXmlNode child in layout.Children()) {
                    if (child.Is(OfficeXmlNamespaces.SpreadsheetML, "sheetData"))
                        WriteGrid(writer, source, sharedStringIndex, recordMutation, includeLayout: true, styles, cancellationToken);
                    else
                        WriteFragment(writer, child);
                }
            }
            writer.WriteEndElement();
        }

        private static void WriteGrid(XmlWriter writer, XlsbSheet source, Func<string, int?> sharedStringIndex,
            Action<string?, bool> recordMutation, bool includeLayout, XlsbStylesheet? styles, CancellationToken cancellationToken) {
            writer.WriteStartElement("sheetData", Namespace);
            ArrayProjection? arrays = ArrayProjection.Create(source, cancellationToken);
            IReadOnlyList<ArrayCell> arrayCells = arrays == null ? Array.Empty<ArrayCell>() : arrays.Cells;
            int cellIndex = 0, rowIndex = 0, arrayIndex = 0;
            IReadOnlyList<XlsbRowInfo> rows = includeLayout ? source.Rows : Array.Empty<XlsbRowInfo>();
            while (cellIndex < source.Cells.Count || rowIndex < rows.Count || arrayIndex < arrayCells.Count) {
                cancellationToken.ThrowIfCancellationRequested();
                XlsbRowInfo? metadata = rowIndex < rows.Count ? rows[rowIndex] : null;
                int cellRow = cellIndex < source.Cells.Count ? source.Cells[cellIndex].Row : int.MaxValue;
                int arrayRow = arrayIndex < arrayCells.Count ? arrayCells[arrayIndex].Row : int.MaxValue;
                int row = Math.Min(Math.Min(cellRow, arrayRow), metadata?.Row ?? int.MaxValue);
                writer.WriteStartElement("row", Namespace);
                writer.WriteAttributeString("r", InvariantNumberText.Get(row));
                if (metadata?.Row == row) {
                    XlsbWorksheetGeometryMarkup.WriteRowAttributes(writer, metadata);
                    rowIndex++;
                }
                if (row == arrayRow) {
                    WriteArrayRow(row);
                } else {
                    while (cellIndex < source.Cells.Count && source.Cells[cellIndex].Row == row) {
                        cancellationToken.ThrowIfCancellationRequested();
                        XlsbCell cell = source.Cells[cellIndex++];
                        WriteCell(writer, cell, row, cell.Column, sharedStringIndex, recordMutation, includeLayout, styles);
                    }
                }
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
            arrays?.RecordOwners(recordMutation);

            void WriteArrayRow(int row) {
                while (true) {
                    int sourceColumn = cellIndex < source.Cells.Count && source.Cells[cellIndex].Row == row
                        ? source.Cells[cellIndex].Column : int.MaxValue;
                    int arrayColumn = arrayIndex < arrayCells.Count && arrayCells[arrayIndex].Row == row
                        ? arrayCells[arrayIndex].Column : int.MaxValue;
                    int column = Math.Min(sourceColumn, arrayColumn);
                    if (column == int.MaxValue) break;
                    cancellationToken.ThrowIfCancellationRequested();

                    XlsbCell? cell = column == sourceColumn ? source.Cells[cellIndex++] : null;
                    bool arrayMember = column == arrayColumn;
                    ArrayFormula? owner = arrayMember ? arrayCells[arrayIndex++].Owner : null;
                    WriteCell(writer, cell, row, column, sharedStringIndex, recordMutation, includeLayout, styles, arrayMember, owner);
                }
            }
        }

        private static void WriteCell(XmlWriter writer, XlsbCell? cell, int row, int column, Func<string, int?> sharedStringIndex,
            Action<string?, bool> recordMutation, bool richText, XlsbStylesheet? styles,
            bool arrayMember = false, ArrayFormula? owner = null) {
            string reference = ExcelA1.CellReference(row, column);
            (string? type, string? value) = cell == null ? (null, null) : CellValue(cell, sharedStringIndex);
            string? sourceFormula = string.IsNullOrWhiteSpace(cell?.FormulaText) ? null : ExcelSanitizer.SanitizeFormula(cell!.FormulaText!);
            string? formula = arrayMember ? owner?.Text : sourceFormula;
            writer.WriteStartElement("c", Namespace);
            writer.WriteAttributeString("r", reference);
            if (cell != null) writer.WriteAttributeString("s", InvariantNumberText.Get(cell.StyleIndex));
            // BIFF12 text runs project to an inline string. Keep the initial shared-string
            // lookup because it determines indices of later plain strings in the same workbook.
            bool inline = richText && cell is { TextRuns: not null, FormulaBytes: null, Value: string };
            if (inline) {
                type = "inlineStr";
                value = null;
            }
            if (type != null) writer.WriteAttributeString("t", type);
            if (formula != null) {
                writer.WriteStartElement("f", Namespace);
                if (owner != null) {
                    writer.WriteAttributeString("t", "array");
                    writer.WriteAttributeString("ref", owner.Reference);
                    owner.RetainsCache = value != null;
                }
                if (value != null) writer.WriteAttributeString("ca", "1");
                writer.WriteString(formula);
                writer.WriteEndElement();
            }
            if (inline) WriteRichText(writer, (string)cell!.Value!, cell.TextRuns!, styles);
            else if (value != null) writer.WriteElementString("v", Namespace, value);
            writer.WriteEndElement();

            if (cell != null) recordMutation(sourceFormula == null ? null : reference, value != null || inline);
        }

        private static void WriteRichText(XmlWriter writer, string text, XlsbTextRun[] runs, XlsbStylesheet? styles) {
            writer.WriteStartElement("is", Namespace);
            if (runs[0].Start > 0) WriteRun(writer, text.Substring(0, runs[0].Start), null);
            for (int index = 0; index < runs.Length; index++) {
                int begin = runs[index].Start;
                int end = index + 1 < runs.Length ? runs[index + 1].Start : text.Length;
                int fontIndex = runs[index].Font;
                XlsbFont? font = styles != null && fontIndex < styles.Fonts.Count ? styles.Fonts[fontIndex] : null;
                WriteRun(writer, text.Substring(begin, end - begin), font);
            }
            writer.WriteEndElement();
        }

        private static void WriteRun(XmlWriter writer, string text, XlsbFont? font) {
            writer.WriteStartElement("r", Namespace);
            if (font != null) {
                OfficeXmlTree tree = OfficeXmlTree.Create(new OfficeXmlNamespaces.Table(), OfficeXmlNamespaces.SpreadsheetML, "rPr", 12);
                var props = new OfficeXmlNode(tree, tree.Root);
                XlsbStylesheetMarkup.WriteFont(props, font, run: true);
                WriteFragment(writer, props);
            }
            writer.WriteStartElement("t", Namespace);
            writer.WriteAttributeString("xml", "space", "http://www.w3.org/XML/1998/namespace", "preserve");
            writer.WriteString(ExcelSanitizer.EncodeXString(text));
            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        private static void WriteFragment(XmlWriter writer, OfficeXmlNode node) {
            using var stream = new MemoryStream(node.OuterXml(), writable: false);
            using XmlReader reader = XmlReader.Create(stream);
            reader.MoveToContent();
            writer.WriteNode(reader, defattr: true);
        }

        private static (string? Type, string? Value) CellValue(XlsbCell cell, Func<string, int?> sharedStringIndex) {
            switch (cell.Kind) {
                case XlsbCellValueKind.Number:
                    double number = (double)cell.Value!;
                    return double.IsNaN(number) || double.IsInfinity(number)
                        ? ("e", ExcelSheet.NonFiniteNumberErrorText)
                        : ("n", InvariantNumberText.Get(number));
                case XlsbCellValueKind.Boolean:
                    return ("b", cell.Value is true ? "1" : "0");
                case XlsbCellValueKind.Error:
                    return ("e", cell.Value as string ?? "#VALUE!");
                case XlsbCellValueKind.Text:
                    string text = (string?)cell.Value ?? string.Empty;
                    if (text.Length == 0) return ("str", string.Empty);
                    CoerceValueHelper.ValidateSharedStringLength(text, nameof(text));
                    return sharedStringIndex(text) is int index ? ("s", InvariantNumberText.Get(index)) : ("str", ExcelSanitizer.EncodeXString(text));
                default:
                    return (null, null);
            }
        }
    }
}
