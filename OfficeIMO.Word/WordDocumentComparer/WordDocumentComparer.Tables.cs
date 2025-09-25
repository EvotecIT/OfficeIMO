namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private static void CompareTables(WordDocument source, WordDocument target, WordDocument result) {
            int count = Math.Min(source.Tables.Count, target.Tables.Count);

            for (int i = 0; i < count; i++) {
                CompareTable(source.Tables[i], target.Tables[i], result.Tables[i]);
            }

            for (int i = count; i < source.Tables.Count; i++) {
                WordTable resTable = result.Tables[i];
                WordTable srcTable = source.Tables[i];
                for (int r = 0; r < srcTable.RowsCount; r++) {
                    for (int c = 0; c < srcTable.Rows[r].CellsCount; c++) {
                        string text = srcTable.Rows[r].Cells[c].Paragraphs.FirstOrDefault()?.Text ?? string.Empty;
                        resTable.Rows[r].Cells[c].AddParagraph(removeExistingParagraphs: true).AddDeletedText(text, "Comparer");
                    }
                }
            }

            for (int i = count; i < target.Tables.Count; i++) {
                WordTable tgtTable = target.Tables[i];
                WordTable resTable = result.AddTable(tgtTable.RowsCount, tgtTable.Rows.First().CellsCount);
                for (int r = 0; r < tgtTable.RowsCount; r++) {
                    for (int c = 0; c < tgtTable.Rows[r].CellsCount; c++) {
                        string text = tgtTable.Rows[r].Cells[c].Paragraphs.FirstOrDefault()?.Text ?? string.Empty;
                        resTable.Rows[r].Cells[c].AddParagraph(removeExistingParagraphs: true).AddInsertedText(text, "Comparer");
                    }
                }
            }
        }

        private static void CompareTable(WordTable source, WordTable target, WordTable result) {
            int rowCount = Math.Min(source.RowsCount, target.RowsCount);

            for (int i = 0; i < rowCount; i++) {
                CompareTableRow(source.Rows[i], target.Rows[i], result.Rows[i]);
            }

            for (int i = rowCount; i < source.RowsCount; i++) {
                WordTableRow srcRow = source.Rows[i];
                WordTableRow resRow = result.Rows[i];
                for (int c = 0; c < srcRow.CellsCount; c++) {
                    string text = srcRow.Cells[c].Paragraphs.FirstOrDefault()?.Text ?? string.Empty;
                    resRow.Cells[c].AddParagraph(removeExistingParagraphs: true).AddDeletedText(text, "Comparer");
                }
            }

            for (int i = rowCount; i < target.RowsCount; i++) {
                WordTableRow tgtRow = target.Rows[i];
                WordTableRow resRow = result.AddRow(tgtRow.CellsCount);
                for (int c = 0; c < tgtRow.CellsCount; c++) {
                    string text = tgtRow.Cells[c].Paragraphs.FirstOrDefault()?.Text ?? string.Empty;
                    resRow.Cells[c].AddParagraph(removeExistingParagraphs: true).AddInsertedText(text, "Comparer");
                }
            }
        }

        private static void CompareTableRow(WordTableRow source, WordTableRow target, WordTableRow result) {
            int cellCount = Math.Min(source.CellsCount, target.CellsCount);

            for (int i = 0; i < cellCount; i++) {
                WordParagraph srcPara = source.Cells[i].Paragraphs.FirstOrDefault() ?? source.Cells[i].AddParagraph();
                WordParagraph tgtPara = target.Cells[i].Paragraphs.FirstOrDefault() ?? target.Cells[i].AddParagraph();
                WordParagraph resPara = result.Cells[i].Paragraphs.FirstOrDefault() ?? result.Cells[i].AddParagraph();
                CompareParagraph(srcPara, tgtPara, resPara);
            }

            for (int i = cellCount; i < source.CellsCount; i++) {
                string text = source.Cells[i].Paragraphs.FirstOrDefault()?.Text ?? string.Empty;
                result.Cells[i].AddParagraph(removeExistingParagraphs: true).AddDeletedText(text, "Comparer");
            }
        }
    }
}
