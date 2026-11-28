using OfficeIMO.Word.OpenXml;

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
                for (int r = 0; r < srcTable.RowCount; r++) {
                    for (int c = 0; c < srcTable.Rows[r].CellsCount; c++) {
                        string text = srcTable.Rows[r].Cells[c].Paragraphs.FirstOrDefault()?.Text ?? string.Empty;
                        resTable.Rows[r].Cells[c].AddParagraph(removeExistingParagraphs: true).AddDeletedText(text, "Comparer");
                    }
                }
            }

            for (int i = count; i < target.Tables.Count; i++) {
                WordTable tgtTable = target.Tables[i];
                OfficeXmlNode clonedTable = WordPackageView.ElementOf(tgtTable).Clone();
                MarkParagraphsAsInserted(clonedTable);
                WordPackageView.SdkBody(result).Append(clonedTable);
            }
        }

        private static void CompareTable(WordTable source, WordTable target, WordTable result) {
            int rowCount = Math.Min(source.RowCount, target.RowCount);

            for (int i = 0; i < rowCount; i++) {
                CompareTableRow(source.Rows[i], target.Rows[i], result.Rows[i]);
            }

            for (int i = rowCount; i < source.RowCount; i++) {
                WordTableRow srcRow = source.Rows[i];
                WordTableRow resRow = result.Rows[i];
                for (int c = 0; c < srcRow.CellsCount; c++) {
                    string text = srcRow.Cells[c].Paragraphs.FirstOrDefault()?.Text ?? string.Empty;
                    resRow.Cells[c].AddParagraph(removeExistingParagraphs: true).AddDeletedText(text, "Comparer");
                }
            }

            for (int i = rowCount; i < target.RowCount; i++) {
                WordTableRow tgtRow = target.Rows[i];
                OfficeXmlNode clonedRow = WordPackageView.ElementOf(tgtRow).Clone();
                MarkParagraphsAsInserted(clonedRow);
                WordPackageView.ElementOf(result).Append(clonedRow);
                result.InvalidateRowElements();
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
                OfficeXmlNode clonedCell = WordPackageView.ElementOf(target.Cells[i]).Clone();
                MarkParagraphsAsInserted(clonedCell);
                WordPackageView.ElementOf(result).Append(clonedCell);
            }
        }

        private static void MarkParagraphsAsInserted(OfficeXmlNode element) {
            foreach (OfficeXmlNode paragraph in TypedDescendants(element).Where(node => node.Is(W, "p"))) {
                MarkParagraphAsInserted(paragraph);
            }
        }

        private static void MarkParagraphAsInserted(OfficeXmlNode paragraph) {
            OfficeXmlNode paragraphProperties = paragraph.Child(W, "pPr");
            paragraphProperties = paragraphProperties.IsNone ? default : paragraphProperties.Clone();
            string text = InnerText(paragraph);

            RemoveAllChildren(paragraph);

            if (!paragraphProperties.IsNone) {
                paragraph.Append(paragraphProperties);
            }

            if (string.IsNullOrEmpty(text)) {
                return;
            }

            OfficeXmlNode run = paragraph.CreateElement(W, "r");
            run.SetAttribute(W, "rsidR", WordHeadersAndFooters.GenerateRsid());
            run.Append(CreateTextElement(run, "t", text));

            OfficeXmlNode inserted = paragraph.CreateElement(W, "ins");
            inserted.SetAttribute(W, "author", "Comparer");
            inserted.SetAttribute(W, "date", WordPackageView.FormatDateTime(DateTime.Now));
            inserted.SetAttribute(W, "id", WordHeadersAndFooters.GenerateRevisionId());
            inserted.Append(run);
            paragraph.Append(inserted);
        }
    }
}
