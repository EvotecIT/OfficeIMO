using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Lists {
        internal static void Example_BasicLists11(string folderPath, bool openWord) {
            string filePath = System.IO.Path.Combine(folderPath, "Document with Lists - Insertion Order.docx");
            using (var document = WordDocument.Create(filePath)) {
                Console.WriteLine("Creating document with paragraphs and a list in between...\n");

                // Add first paragraph
                Console.WriteLine("1. Adding first paragraph: 'First'");
                var first = document.AddParagraph("First");

                // Add last paragraph
                Console.WriteLine("2. Adding last paragraph: 'Last'");
                document.AddParagraph("Last");

                // Add list between paragraphs
                Console.WriteLine("\n3. Adding bulleted list after 'First' paragraph:");
                var list = first.AddList(WordListStyle.Bulleted);
                Console.WriteLine("   - Adding first list item: 'Important'");
                list.AddItem("Important", 0, first);
                Console.WriteLine("   - Adding second list item: 'List', which will be added to the list, but after last paragraph.");
                list.AddItem("List");

                Console.WriteLine("\nFinal document structure:");
                Console.WriteLine("-------------------------");
                for (int i = 0; i < document.Paragraphs.Count; i++) {
                    var prefix = document.Paragraphs[i].IsListItem ? "   • " : "";
                    Console.WriteLine($"{i + 1}. {prefix}{document.Paragraphs[i].Text}");
                }

                document.Save(openWord);
                Console.WriteLine($"\nDocument saved to: {filePath}");
                if (openWord) {
                    Console.WriteLine("Opening document in Word...");
                }
            }
        }
    }
}