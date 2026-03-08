using DocumentFormat.OpenXml.Wordprocessing;

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
                var clonedTable = (Table)tgtTable._table.CloneNode(true);
                MarkTableAsInserted(clonedTable);
                result._wordprocessingDocument.MainDocumentPart!.Document!.Body!.Append(clonedTable);
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
                var clonedRow = (TableRow)tgtRow._tableRow.CloneNode(true);
                MarkRowAsInserted(clonedRow);
                result._table.Append(clonedRow);
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

            for (int i = cellCount; i < target.CellsCount; i++) {
                var clonedCell = (TableCell)target.Cells[i]._tableCell.CloneNode(true);
                MarkCellAsInserted(clonedCell);
                result._tableRow.Append(clonedCell);
            }
        }

        private static void MarkTableAsInserted(Table table) {
            foreach (var paragraph in table.Descendants<Paragraph>()) {
                MarkParagraphAsInserted(paragraph);
            }
        }

        private static void MarkRowAsInserted(TableRow row) {
            foreach (var paragraph in row.Descendants<Paragraph>()) {
                MarkParagraphAsInserted(paragraph);
            }
        }

        private static void MarkCellAsInserted(TableCell cell) {
            foreach (var paragraph in cell.Descendants<Paragraph>()) {
                MarkParagraphAsInserted(paragraph);
            }
        }

        private static void MarkParagraphAsInserted(Paragraph paragraph) {
            var paragraphProperties = (ParagraphProperties?)paragraph.ParagraphProperties?.CloneNode(true);
            var text = paragraph.InnerText;

            paragraph.RemoveAllChildren();

            if (paragraphProperties != null) {
                paragraph.Append(paragraphProperties);
            }

            if (string.IsNullOrEmpty(text)) {
                return;
            }

            var run = new Run();
            run.RsidRunAddition = WordHeadersAndFooters.GenerateRsid();
            run.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });

            var inserted = new InsertedRun {
                Author = "Comparer",
                Date = DateTime.Now,
                Id = WordHeadersAndFooters.GenerateRevisionId()
            };
            inserted.Append(run);
            paragraph.Append(inserted);
        }
    }
}
