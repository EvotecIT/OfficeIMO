using System.Xml;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Excel.Utilities;
using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Projection {
    /// <summary>Streams the XLSB cell grid to SpreadsheetML in the validated BIFF12 row and column order.</summary>
    internal static class XlsbWorksheetGridProjector {
        private static readonly string Namespace = OfficeXmlNamespaces.KnownUri(OfficeXmlNamespaces.SpreadsheetML)!;

        /// <summary>
        /// Array groups require a second projection of their members. Line breaks can add a wrap-text style during
        /// import even when the cell is later assigned its native style index.
        /// </summary>
        internal static bool CanStream(XlsbSheet sheet) {
            if (sheet.FormulaGroups.Any(group => group.Type == 426)) return false;
            foreach (XlsbCell cell in sheet.Cells) {
                if (cell.Value is string text && (text.IndexOf('\n') >= 0 || text.IndexOf('\r') >= 0)) return false;
            }
            return true;
        }

        internal static void WriteWorksheet(Stream output, XlsbSheet source, Func<string, int?> sharedStringIndex,
            Action<string?, bool> recordMutation, CancellationToken cancellationToken, OfficeXmlNode layout = default,
            XlsbStylesheet? styles = null) {
            using XmlWriter writer = XmlWriter.Create(output, new XmlWriterSettings {
                Encoding = new UTF8Encoding(false),
                NewLineHandling = NewLineHandling.Entitize
            });
            writer.WriteStartElement("worksheet", Namespace);
            if (layout.IsNone) {
                WriteGrid(writer, source, sharedStringIndex, recordMutation, includeLayout: false, styles, cancellationToken);
            } else {
                foreach (OfficeXmlNode child in layout.Children()) {
                    if (child.Is(OfficeXmlNamespaces.SpreadsheetML, "sheetData"))
                        WriteGrid(writer, source, sharedStringIndex, recordMutation, includeLayout: true, styles, cancellationToken);
                    else
                        writer.WriteRaw(Encoding.UTF8.GetString(child.OuterXml()));
                }
            }
            writer.WriteEndElement();
        }

        private static void WriteGrid(XmlWriter writer, XlsbSheet source, Func<string, int?> sharedStringIndex,
            Action<string?, bool> recordMutation, bool includeLayout, XlsbStylesheet? styles, CancellationToken cancellationToken) {
            writer.WriteStartElement("sheetData", Namespace);
            int cellIndex = 0, rowIndex = 0;
            IReadOnlyList<XlsbRowInfo> rows = includeLayout ? source.Rows : Array.Empty<XlsbRowInfo>();
            while (cellIndex < source.Cells.Count || rowIndex < rows.Count) {
                cancellationToken.ThrowIfCancellationRequested();
                XlsbRowInfo? metadata = rowIndex < rows.Count ? rows[rowIndex] : null;
                int cellRow = cellIndex < source.Cells.Count ? source.Cells[cellIndex].Row : int.MaxValue;
                int row = Math.Min(cellRow, metadata?.Row ?? int.MaxValue);
                writer.WriteStartElement("row", Namespace);
                writer.WriteAttributeString("r", InvariantNumberText.Get(row));
                if (metadata?.Row == row) {
                    XlsbWorksheetGeometryMarkup.WriteRowAttributes(writer, metadata);
                    rowIndex++;
                }
                while (cellIndex < source.Cells.Count && source.Cells[cellIndex].Row == row) {
                    cancellationToken.ThrowIfCancellationRequested();
                    WriteCell(writer, source.Cells[cellIndex++], sharedStringIndex, recordMutation, includeLayout, styles);
                }
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
        }

        private static void WriteCell(XmlWriter writer, XlsbCell cell, Func<string, int?> sharedStringIndex,
            Action<string?, bool> recordMutation, bool richText, XlsbStylesheet? styles) {
            string reference = ExcelA1.CellReference(cell.Row, cell.Column);
            (string? type, string? value) = CellValue(cell, sharedStringIndex);
            string? formula = string.IsNullOrWhiteSpace(cell.FormulaText) ? null : ExcelSanitizer.SanitizeFormula(cell.FormulaText!);
            writer.WriteStartElement("c", Namespace);
            writer.WriteAttributeString("r", reference);
            writer.WriteAttributeString("s", InvariantNumberText.Get(cell.StyleIndex));
            // BIFF12 text runs project to an inline string. Keep the initial shared-string
            // lookup because it determines indices of later plain strings in the same workbook.
            if (richText && cell.TextRuns != null && cell.FormulaBytes == null && cell.Value is string text) {
                writer.WriteAttributeString("t", "inlineStr");
                WriteRichText(writer, text, cell.TextRuns, styles);
                writer.WriteEndElement();
                recordMutation(null, true);
                return;
            }
            if (type != null) writer.WriteAttributeString("t", type);
            if (formula != null) {
                writer.WriteStartElement("f", Namespace);
                if (value != null) writer.WriteAttributeString("ca", "1");
                writer.WriteString(formula);
                writer.WriteEndElement();
            }
            if (value != null) writer.WriteElementString("v", Namespace, value);
            writer.WriteEndElement();

            recordMutation(formula == null ? null : reference, value != null);
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
                writer.WriteRaw(Encoding.UTF8.GetString(props.OuterXml()));
            }
            writer.WriteStartElement("t", Namespace);
            writer.WriteAttributeString("xml", "space", "http://www.w3.org/XML/1998/namespace", "preserve");
            writer.WriteString(ExcelSanitizer.EncodeXString(text));
            writer.WriteEndElement();
            writer.WriteEndElement();
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
