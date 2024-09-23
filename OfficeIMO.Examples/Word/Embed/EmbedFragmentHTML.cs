using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Embed {

        public static void Example_EmbedHTMLFragment(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with embedded HTML fragment");
            string filePath = System.IO.Path.Combine(folderPath, "EmbeddedFragmentHTML.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                Console.WriteLine("Embedded documents in word: " + document.EmbeddedDocuments.Count);
                Console.WriteLine("Embedded documents in Section 0: " + document.Sections[0].EmbeddedDocuments.Count);

                document.AddParagraph("Add HTML document in DOCX");

                document.AddSection();

                Console.WriteLine("Embedded documents in Section 1: " + document.Sections[1].EmbeddedDocuments.Count);

                var htmlContent = """
                                  <html lang="en">
                                  <P>This is a paragraph.</P>
                                  <P>This is another paragraph.</P>
                                  <P>This is a paragraph with <STRONG>bold</STRONG> text.</P>
                                  <P>This is a paragraph with <EM>italic</EM> text.</P>
                                  <ul>
                                    <li>Item 1</li>
                                    <li>Item 2</li>
                                    <li>Item 3</li>
                                  </ul>
                                  <ol>
                                    <li>Item 1</li>
                                    <li>Item 2</li>
                                    <li>Item 3</li>
                                  </ol>
                                  <P>This is a paragraph with a <A href=""https://www.google.com"">link</A>.</P>
                                  </html>
                                  """;

                document.AddEmbeddedFragment(htmlContent, WordAlternativeFormatImportPartType.Html);

                document.EmbeddedDocuments[0].Save("C:\\TEMP\\EmbeddedFragment.html");

                Console.WriteLine("Embedded documents in word: " + document.EmbeddedDocuments.Count);
                Console.WriteLine("Embedded documents in Section 0: " + document.Sections[0].EmbeddedDocuments.Count);
                Console.WriteLine("Embedded documents in Section 1: " + document.Sections[1].EmbeddedDocuments.Count);
                Console.WriteLine("Content type: " + document.EmbeddedDocuments[0].ContentType);

                document.Save(openWord);
            }
        }
    }
}
