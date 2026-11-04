using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class BasicDocument {
        public static void Example_BasicWordWithLineSpacing(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with margins");
            string filePath = System.IO.Path.Combine(folderPath, "DocumentWithLineSpacing.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var par00 = document.AddParagraph("My text");
                par00.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
                par00.LineSpacingBefore = global::OfficeIMO.OfficeLength.Zero;

                var par01 = document.AddParagraph("My declaration");
                par01.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
                par01.LineSpacingBefore = global::OfficeIMO.OfficeLength.Zero;

                var par02 = document.AddParagraph("My declaration");
                par02.LineSpacing = global::OfficeIMO.Word.WordLineSpacing.Exact(global::OfficeIMO.OfficeLength.Emus(checked((long)(360) * 635)));

                document.Save();
                if (openWord) document.OpenInApplication();
            }
        }
    }
}
