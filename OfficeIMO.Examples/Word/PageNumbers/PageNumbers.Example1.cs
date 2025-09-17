using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Examples.Word {
    internal static partial class PageNumbers {
        internal static void Example_PageNumbers1(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with Page Numbers 1");
            string filePath = System.IO.Path.Combine(folderPath, "Document with PageNumbers.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.Settings.UpdateFieldsOnOpen = true;
                document.AddTableOfContent(tableOfContentStyle: TableOfContentStyle.Template2);
                document.AddHeadersAndFooters();
                var pageNumber = GetDocumentHeaderOrThrow(document).AddPageNumber(WordPageNumberStyle.Dots);
                //var pageNumber = document.Footer!.Default.AddPageNumber(WordPageNumberStyle.VerticalOutline2);
                //var pageNumber = document.Footer!.Default.AddPageNumber(WordPageNumberStyle.Dots);

                pageNumber.ParagraphAlignment = JustificationValues.Center;

                document.AddPageBreak();

                document.AddHorizontalLine(BorderValues.Double);

                if (document.Sections.Count > 0) {
                    document.Sections[0].AddHorizontalLine();
                }

                var wordListToc = document.AddTableOfContentList(WordListStyle.Numbered);

                wordListToc.AddItem("This is first item");

                wordListToc.AddItem("This is second item");

                document.AddPageBreak();

                wordListToc.AddItem("Text 2.1", 1);

                wordListToc.AddItem("Text 2.1", 1);

                wordListToc.AddItem("Text 2.1", 1);

                wordListToc.AddItem("Text 2.2", 2);

                var para = document.AddParagraph("Let's show everyone how to create a list within already defined list");
                para.CapsStyle = CapsStyle.Caps;
                para.Highlight = HighlightColorValues.DarkMagenta;

                var wordList = document.AddList(WordListStyle.Bulleted);

                wordList.AddItem("List Item 1");
                wordList.AddItem("List Item 2");
                wordList.AddItem("List Item 3");
                wordList.AddItem("List Item 3.1", 1);
                wordList.AddItem("List Item 3.2", 1);
                wordList.AddItem("List Item 3.3", 2);

                wordListToc.AddItem("Text 2.3", 2);

                wordListToc.AddItem("Text 3.3", 3);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                // we loaded document, lets add some text to continue
                document.AddParagraph().SetColor(Color.CornflowerBlue).SetText("This is some text");

                // we loaded document, lets add page break to continue
                document.AddPageBreak();

                // lets find a list which has items which suggest it's a TOC attached list
                WordList? wordListToc = null;
                foreach (var list in document.Lists) {
                    if (list.IsToc) {
                        wordListToc = list;
                    }
                }

                // finally lets add another list item
                if (wordListToc != null) {
                    wordListToc.AddItem("Text 4.4", 2);
                }

                document.Settings.UpdateFieldsOnOpen = true;
                document.Save(openWord);
            }
        }
    }
}
