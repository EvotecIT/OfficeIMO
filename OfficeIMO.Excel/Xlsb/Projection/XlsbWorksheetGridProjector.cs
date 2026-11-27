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
            Action<string?, bool> recordMutation, CancellationToken cancellationToken) {
            using XmlWriter writer = XmlWriter.Create(output, new XmlWriterSettings {
                Encoding = new UTF8Encoding(false),
                NewLineHandling = NewLineHandling.Entitize
            });
            writer.WriteStartElement("worksheet", Namespace);
            writer.WriteStartElement("sheetData", Namespace);

            int row = 0;
            foreach (XlsbCell cell in source.Cells) {
                cancellationToken.ThrowIfCancellationRequested();
                if (cell.Row != row) {
                    if (row != 0) writer.WriteEndElement();
                    row = cell.Row;
                    writer.WriteStartElement("row", Namespace);
                    writer.WriteAttributeString("r", InvariantNumberText.Get(row));
                }
                WriteCell(writer, cell, sharedStringIndex, recordMutation);
            }
            if (row != 0) writer.WriteEndElement();

            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        private static void WriteCell(XmlWriter writer, XlsbCell cell, Func<string, int?> sharedStringIndex,
            Action<string?, bool> recordMutation) {
            string reference = ExcelA1.CellReference(cell.Row, cell.Column);
            (string? type, string? value) = CellValue(cell, sharedStringIndex);
            string? formula = string.IsNullOrWhiteSpace(cell.FormulaText) ? null : ExcelSanitizer.SanitizeFormula(cell.FormulaText!);
            writer.WriteStartElement("c", Namespace);
            writer.WriteAttributeString("r", reference);
            writer.WriteAttributeString("s", InvariantNumberText.Get(cell.StyleIndex));
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
