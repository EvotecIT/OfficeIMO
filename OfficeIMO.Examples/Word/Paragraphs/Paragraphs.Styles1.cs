using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

internal static partial class Paragraphs {

    internal static void Example_BasicParagraphStyles(string folderPath, bool openWord) {
        Console.WriteLine("[*] Creating standard document with Paragraph Styles");
        string filePath = System.IO.Path.Combine(folderPath, "Document with Paragraph Styles.docx");
        using (WordDocument document = WordDocument.Create(filePath)) {
            var listOfStyles = (WordParagraphStyles[])Enum.GetValues(typeof(WordParagraphStyles));
            foreach (var style in listOfStyles) {
                var paragraph = document.AddParagraph(style.ToString());
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Style = style;
            }

            document.Save(openWord);
        }
    }

}
