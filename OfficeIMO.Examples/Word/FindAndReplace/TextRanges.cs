using System.Text.RegularExpressions;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class FindAndReplace {
        internal static void Example_TextRanges(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document - Text ranges: regex replace, spanning bookmarks, range comments and links");
            string filePath = Path.Combine(folderPath, "Text Ranges.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordParagraph terms = document.AddParagraph("The contract value is ");
                terms.AddText("1,000").SetBold();
                terms.AddText(" EUR, signed on 01/12/2026. Delivery: TBD.");
                document.AddParagraph("Read [the guide](https://officeimo.com) before signing.");
                document.AddParagraph("Total: ").AddCrossReference("Amount", WordCrossReferenceType.Bookmark);

                // Regex replace: group substitution, a computed value, and formatted content.
                document.ReplaceAll(new Regex(@"(\d{2})/(\d{2})/(\d{4})"), "$3-$2-$1");
                document.ReplaceAll(new Regex(@"\bTBD\b"), match => "to be confirmed");
                document.ReplaceAll(new Regex(@"\[(.+?)\]\((.+?)\)"),
                    (match, paragraph) => paragraph.AddHyperLink(match.Groups[1].Value, new Uri(match.Groups[2].Value), addStyle: true));

                // A search hit is a range: bookmark it, comment on it and format it.
                WordRange amount = document.Find("1,000 EUR").Ranges[0];
                amount.AddBookmark("Amount");
                amount.AddComment("Przemysław Kłys", "PK", comment => comment.AddText("Confirm with ").AddText("finance").SetBold());
                document.Find("to be confirmed").Ranges[0].Format(run => run.SetItalic());

                // Bookmark text drives REF fields.
                document.Bookmarks.First(bookmark => bookmark.Name == "Amount").Text = "2,500 EUR";
                document.UpdateFieldsAndGetReport();
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Console.WriteLine($"Bookmark 'Amount': {document.Bookmarks.First(bookmark => bookmark.Name == "Amount").Text}");
                Console.WriteLine($"First paragraph: {document.Paragraphs[0].GetRange().Text}");
                if (openWord) document.OpenInApplication();
            }
        }
    }
}
