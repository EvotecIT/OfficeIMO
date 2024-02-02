using System;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Lists {
        internal static void Example_BasicLists7(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with lists - Document 7");
            string filePath = System.IO.Path.Combine(folderPath, "Document with Lists10.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {

                // add list and nest a list
                WordList wordList1 = document.AddList(WordListStyle.Headings111, false);
                Console.WriteLine("List (0) ElementsCount (0): " + wordList1.ListItems.Count);
                Console.WriteLine("List (0) ElementsCount (1): " + document.Lists[0].ListItems.Count);
                wordList1.AddItem("Text 1");
                Console.WriteLine("List (0) ElementsCount (1): " + document.Lists[0].ListItems.Count);
                Console.WriteLine("Lists count (1): " + document.Lists.Count);
                Console.WriteLine("List (0) ElementsCount (1): " + wordList1.ListItems.Count);

                document.AddBreak();
                wordList1.RestartNumberingAfterBreak = true;
                wordList1.AddItem("Text 1");

                wordList1.AddItem("Text 2");
                Console.WriteLine("List (0) ElementsCount (2): " + wordList1.ListItems.Count);
                wordList1.AddItem("Text 2.1", 1);
                Console.WriteLine("List (0) ElementsCount (3): " + wordList1.ListItems.Count);

                WordList wordListNested = document.AddList(WordListStyle.Bulleted, false);
                wordListNested.AddItem("Nested 1", 1);
                wordListNested.AddItem("Nested 2", 1);

                WordList wordList2 = document.AddList(WordListStyle.Headings111, true);
                Console.WriteLine("List 2 - Restart numbering: " + wordList2.RestartNumbering);
                wordList2.AddItem("Section 2");
                wordList2.AddItem("Section 2.1", 1);

                WordList wordList3 = document.AddList(WordListStyle.Headings111, true);
                Console.WriteLine("List 3 - Restart numbering: " + wordList3.RestartNumbering);
                wordList3.RestartNumbering = true;
                Console.WriteLine("List 3 - Restart numbering after change: " + wordList3.RestartNumbering);
                wordList3.AddItem("Section 1");
                wordList3.AddItem("Section 1.1", 1);

                WordList wordList4 = document.AddList(WordListStyle.Headings111, true);
                //wordList4.RestartNumbering = true;
                wordList4.AddItem("Section 2");
                wordList4.AddItem("Section 2.1", 1);

                WordList wordList5 = document.AddList(WordListStyle.Headings111, true);
                //wordList5.RestartNumbering = true;
                wordList5.AddItem("Section 3");
                wordList5.AddItem("Section 3.1", 1);

                WordList wordList6 = document.AddList(WordListStyle.Headings111);
                wordList1.AddItem("Text 4");
                wordList1.AddItem("Text 4.1", 1);

                document.AddBreak();

                //// add a table
                var table = document.AddTable(3, 3);

                table.AddRow(2);

                //// add a list to a table and attach it to a first paragraph
                var listInsideTable = table.Rows[0].Cells[0].Paragraphs[0].AddList(WordListStyle.Bulleted);

                // this will force the current Paragraph to be converted into a list item and overwrite it's text
                Console.WriteLine("Table List (0) ElementsCount (0): " + listInsideTable.ListItems.Count);
                listInsideTable.AddItem("text", 0, table.Rows[0].Cells[0].Paragraphs[0]);
                Console.WriteLine("Table List (0) ElementsCount (1): " + listInsideTable.ListItems.Count);

                // add new items to the list (as last paragraph)
                listInsideTable.AddItem("Test 1");
                Console.WriteLine("Table List (0) ElementsCount: " + listInsideTable.ListItems.Count);

                // add new items to the list (as last paragraph)
                listInsideTable.AddItem("Test 2");
                Console.WriteLine("Table List (0) ElementsCount: " + listInsideTable.ListItems.Count);

                table.Rows[0].Cells[0].AddParagraph("Test Text 1");
                listInsideTable.AddItem("Test 3");
                table.Rows[0].Cells[0].AddParagraph("Test Text 2");

                table.Rows[1].Cells[0].Paragraphs[0].Text = "Text Row 1 - 0";
                table.Rows[1].Cells[0].AddParagraph("Text Row 1 - 1").AddText(" More text").AddParagraph("Text Row 1 - 2");

                // add a list to a table by adding it to a cell, notice that that the first paragraph is empty
                var listInsideTableColumn2 = table.Rows[0].Cells[1].AddList(WordListStyle.Bulleted);
                Console.WriteLine("Table List (1) ElementsCount (0): " + listInsideTableColumn2.ListItems.Count);
                listInsideTableColumn2.AddItem("Test 1 - Column 2");
                Console.WriteLine("Table List (1) ElementsCount (1): " + listInsideTableColumn2.ListItems.Count);
                listInsideTableColumn2.AddItem("Test 2  - Column 2");
                Console.WriteLine("Table List (1) ElementsCount (2): " + listInsideTableColumn2.ListItems.Count);

                // add a list to a table by adding it to a cell, notice that I'm adding text before list first
                table.Rows[0].Cells[2].Paragraphs[0].Text = "This is list: ";
                // add list, and add all items
                var listInsideTableColumn3 = table.Rows[0].Cells[2].AddList(WordListStyle.Bulleted);
                Console.WriteLine("Table List (2) ElementsCount: " + listInsideTableColumn3.ListItems.Count);
                listInsideTableColumn3.AddItem("Test 1 - Column 2");
                Console.WriteLine("Table List (2) ElementsCount: " + listInsideTableColumn3.ListItems.Count);
                listInsideTableColumn3.AddItem("Test 2  - Column 2");
                Console.WriteLine("Table List (2) ElementsCount: " + listInsideTableColumn3.ListItems.Count);


                // add a list to a table by adding it to a cell, notice that I'm adding text before list first
                // but then convert that line into a list item
                table.Rows[1].Cells[2].Paragraphs[0].Text = "This is list as list item: ";
                // add list, and add all items
                var listInsideTableColumn4 = table.Rows[0].Cells[2].AddList(WordListStyle.Bulleted);

                listInsideTableColumn4.AddItem(table.Rows[1].Cells[2].Paragraphs[0]); // convert to list item

                Console.WriteLine("Table List (2) ElementsCount: " + listInsideTableColumn4.ListItems.Count);
                listInsideTableColumn4.AddItem("Test 1 - Column 2");
                Console.WriteLine("Table List (2) ElementsCount: " + listInsideTableColumn4.ListItems.Count);
                listInsideTableColumn4.AddItem("Test 2  - Column 2");
                Console.WriteLine("Table List (2) ElementsCount: " + listInsideTableColumn4.ListItems.Count);

                Console.WriteLine("Lists count in a document (11): " + document.Lists.Count);

                document.Save(false);
            }


            using (WordDocument document = WordDocument.Load(filePath)) {
                Console.WriteLine("Lists count in a document (11): " + document.Lists.Count);

                document.Lists[0].AddItem("More then enough");

                document.AddHeadersAndFooters();

                var listInHeader = document.Header.Default.AddList(WordListStyle.Bulleted);

                listInHeader.AddItem("Test Header 1");

                document.Footer.Default.AddParagraph("Test Me Header");

                listInHeader.AddItem("Test Header 2");


                var listInFooter = document.Footer.Default.AddList(WordListStyle.Headings111);

                listInFooter.AddItem("Test Footer 1");

                document.Footer.Default.AddParagraph("Test Me Footer");

                listInFooter.AddItem("Test Footer 2");

                document.Save(openWord);
            }
        }
    }
}
