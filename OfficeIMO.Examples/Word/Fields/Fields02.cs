using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Examples.Utils;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Fields {
        internal static void Example_DocumentWithFields02(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with tables");
            string filePath = System.IO.Path.Combine(folderPath, "Document with Fields02.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Basic paragraph");
                paragraph.ParagraphAlignment = JustificationValues.Center;

                document.AddParagraph();

                document.AddHeadersAndFooters();

                // added page number using fields which triggers fields refresh
                document.AddField(WordFieldType.Page).AddText(" of ").AddField(WordFieldType.NumPages);

                document.AddField(WordFieldType.Author);

                document.AddField(WordFieldType.GreetingLine);

                // added page number using dedicated way
                var headers = Guard.NotNull(document.Header, "Document headers must exist after enabling headers.");
                var defaultHeader = Guard.NotNull(headers.Default, "Default header must exist after enabling headers.");
                var pageNumber = defaultHeader.AddPageNumber(WordPageNumberStyle.Roman);

                document.Save(openWord);
            }
        }
    }
}
