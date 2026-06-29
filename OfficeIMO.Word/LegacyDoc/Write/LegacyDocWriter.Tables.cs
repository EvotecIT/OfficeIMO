using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static void AppendTable(StringBuilder text, List<LegacyDocWritableRun> runs, List<LegacyDocWritableParagraph> paragraphFormats, Table table) {
            ThrowIfUnsupportedTableShape(table);

            TableRow[] rows = table.Elements<TableRow>().ToArray();
            if (rows.Length == 0) {
                throw new NotSupportedException("Native DOC saving supports simple tables only when at least one row is present.");
            }

            foreach (TableRow row in rows) {
                TableCell[] cells = row.Elements<TableCell>().ToArray();
                if (cells.Length == 0) {
                    throw new NotSupportedException("Native DOC saving supports simple tables only when every row contains at least one cell.");
                }

                foreach (TableCell cell in cells) {
                    int cellStart = text.Length;
                    LegacyDocWritableParagraphFormatting paragraphFormatting = AppendTableCell(text, runs, cell);
                    text.Append('\a');
                    if (paragraphFormatting.HasFormatting) {
                        paragraphFormats.Add(new LegacyDocWritableParagraph(cellStart, text.Length - cellStart, paragraphFormatting));
                    }
                }

                text.Append('\a');
            }

            text.Append('\r');
        }

        private static void ThrowIfUnsupportedTableShape(Table table) {
            foreach (OpenXmlElement child in table.ChildElements) {
                switch (child) {
                    case TableProperties tableProperties:
                        ThrowIfUnsupportedTableProperties(tableProperties);
                        break;
                    case TableGrid tableGrid:
                        ThrowIfUnsupportedTableGrid(tableGrid);
                        break;
                    case TableRow:
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple tables only. Unsupported table element: {child.LocalName}.");
                }
            }
        }

        private static void ThrowIfUnsupportedTableProperties(TableProperties tableProperties) {
            foreach (OpenXmlElement property in tableProperties.ChildElements) {
                switch (property) {
                    case TableStyle tableStyle:
                        if (!string.Equals(tableStyle.Val?.Value, "TableGrid", StringComparison.OrdinalIgnoreCase)) {
                            throw new NotSupportedException("Native DOC saving supports simple tables only with the TableGrid table style.");
                        }
                        break;
                    case TableWidth tableWidth:
                        if (tableWidth.Type?.Value != TableWidthUnitValues.Auto || tableWidth.Width?.Value != "0") {
                            throw new NotSupportedException("Native DOC saving supports simple tables only with the default automatic table width.");
                        }
                        break;
                    case TableLook:
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple tables only. Unsupported table property: {property.LocalName}.");
                }
            }
        }

        private static void ThrowIfUnsupportedTableGrid(TableGrid tableGrid) {
            foreach (OpenXmlElement child in tableGrid.ChildElements) {
                if (child is GridColumn) {
                    continue;
                }

                throw new NotSupportedException($"Native DOC saving supports simple tables only. Unsupported table grid element: {child.LocalName}.");
            }
        }

        private static LegacyDocWritableParagraphFormatting AppendTableCell(StringBuilder text, List<LegacyDocWritableRun> runs, TableCell cell) {
            Paragraph? paragraph = null;
            foreach (OpenXmlElement child in cell.ChildElements) {
                switch (child) {
                    case TableCellProperties cellProperties:
                        ThrowIfUnsupportedTableCellProperties(cellProperties);
                        break;
                    case Paragraph cellParagraph:
                        if (paragraph != null) {
                            throw new NotSupportedException("Native DOC saving supports simple tables only with one paragraph per cell.");
                        }

                        paragraph = cellParagraph;
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple tables only. Unsupported table cell element: {child.LocalName}.");
                }
            }

            if (paragraph != null) {
                return AppendTableCellParagraph(text, runs, paragraph);
            }

            return LegacyDocWritableParagraphFormatting.Plain;
        }

        private static void ThrowIfUnsupportedTableCellProperties(TableCellProperties cellProperties) {
            foreach (OpenXmlElement property in cellProperties.ChildElements) {
                switch (property) {
                    case TableCellWidth cellWidth:
                        if (cellWidth.Type?.Value != TableWidthUnitValues.Dxa || cellWidth.Width?.Value != "2400") {
                            throw new NotSupportedException("Native DOC saving supports simple tables only with the default table cell width.");
                        }
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple tables only. Unsupported table cell property: {property.LocalName}.");
                }
            }
        }

        private static LegacyDocWritableParagraphFormatting AppendTableCellParagraph(StringBuilder text, List<LegacyDocWritableRun> runs, Paragraph paragraph) {
            LegacyDocWritableParagraphFormatting paragraphFormatting = ReadSupportedParagraphFormatting(paragraph.ParagraphProperties);

            foreach (OpenXmlElement child in paragraph.ChildElements) {
                switch (child) {
                    case ParagraphProperties:
                        break;
                    case Run run:
                        AppendSupportedRunText(text, runs, run);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple table cell paragraphs only with text runs. Unsupported paragraph element: {child.LocalName}.");
                }
            }

            return paragraphFormatting;
        }
    }
}
