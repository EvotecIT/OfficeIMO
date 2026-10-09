namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private static partial class DirectDataSetWorkbookWriter {
            // Blank records retain formatting and the first/last columns of a
            // sparse range without turning missing values into empty strings.
            private static void WriteSparseMissingTabularCell(
                TextWriter writer,
                string rowReference,
                string cellReferencePrefix,
                string? styleAttribute,
                bool retainRangeBoundary,
                ref bool rowStarted) {
                if (styleAttribute == null && !retainRangeBoundary) return;
                if (!rowStarted) {
                    writer.Write("<row r=\"");
                    writer.Write(rowReference);
                    writer.Write("\">");
                    rowStarted = true;
                }
                writer.Write(cellReferencePrefix);
                writer.Write(rowReference);
                writer.Write('"');
                writer.Write(styleAttribute);
                writer.Write("/>");
            }

            private static void CompleteSparseTabularRow(
                TextWriter writer,
                string rowReference,
                string[] cellReferencePrefixes,
                string?[]? styleAttributes,
                bool rowStarted) {
                if (rowStarted) {
                    writer.Write("</row>");
                    return;
                }

                // Blank records at both ends retain the tabular range even for
                // readers that discover the first column from cell references.
                writer.Write("<row r=\"");
                writer.Write(rowReference);
                writer.Write("\">");
                if (cellReferencePrefixes.Length != 0) {
                    int lastColumn = cellReferencePrefixes.Length - 1;
                    if (lastColumn != 0) {
                        writer.Write(cellReferencePrefixes[0]);
                        writer.Write(rowReference);
                        writer.Write('"');
                        writer.Write(styleAttributes?[0]);
                        writer.Write("/>");
                    }
                    writer.Write(cellReferencePrefixes[lastColumn]);
                    writer.Write(rowReference);
                    writer.Write('"');
                    writer.Write(styleAttributes?[lastColumn]);
                    writer.Write("/>");
                }
                writer.Write("</row>");
            }
        }
    }
}
