using System;
using System.IO;
using OfficeIMO.Examples.Utils;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class HyperLinks {

        internal static void Example_FormattedHyperLinksAdvanced(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with advanced formatted hyperlinks");
            string filePath = Path.Combine(folderPath, "FormattedHyperLinksAdvanced.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var defaultHeader = document.HeaderDefaultOrCreate;
                var defaultFooter = document.FooterDefaultOrCreate;

                var paragraph = document.AddParagraph("Visit ");
                var google = paragraph.AddHyperLink("Google", new Uri("https://google.com"), addStyle: true);
                google.Bold = true;
                var baseLink = Guard.NotNull(google.Hyperlink, "Expected the hyperlink to be created for Google link.");

                baseLink.InsertFormattedHyperlinkBefore("Bing", new Uri("https://bing.com"));
                var duplicate = WordHyperLink.DuplicateHyperlink(baseLink);
                duplicate.Text = "Google Copy";

                var yahoo = baseLink.InsertFormattedHyperlinkAfter("Yahoo", new Uri("https://yahoo.com"));
                yahoo.CopyFormattingFrom(baseLink);

                var headerPara = defaultHeader.AddParagraph("Search with ");
                var duck = headerPara.AddHyperLink("DuckDuckGo", new Uri("https://duckduckgo.com"), addStyle: true);
                var duckLink = Guard.NotNull(duck.Hyperlink, "Expected DuckDuckGo hyperlink to be created.");
                duckLink.InsertFormattedHyperlinkAfter("Startpage", new Uri("https://startpage.com"));

                var footerPara = defaultFooter.AddParagraph("Code on ");
                var gitHub = footerPara.AddHyperLink("GitHub", new Uri("https://github.com"), addStyle: true);
                var gitHubLink = Guard.NotNull(gitHub.Hyperlink, "Expected GitHub hyperlink to be created.");
                gitHubLink.InsertFormattedHyperlinkBefore("GitLab", new Uri("https://gitlab.com"));

                document.Save(openWord);
            }
        }
    }
}
