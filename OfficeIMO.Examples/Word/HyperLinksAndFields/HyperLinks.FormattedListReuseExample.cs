using System;
using System.IO;
using OfficeIMO.Word;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Examples.Word {
    internal static partial class HyperLinks {
        internal static void Example_FormattedHyperLinksListReuse(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with hyperlink lists and formatting reuse");
            string filePath = Path.Combine(folderPath, "FormattedHyperLinksLists.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var firstList = document.AddList(WordListStyle.Bulleted);
                var googlePara = firstList.AddItem("").AddHyperLink("Google", new Uri("https://google.com"), addStyle: true);
                googlePara.Bold = true;
                var googleLink = googlePara.Hyperlink;

                var bingPara = firstList.AddItem("").AddHyperLink("Bing", new Uri("https://bing.com"), addStyle: true);
                bingPara.Italic = true;
                var bingLink = bingPara.Hyperlink;

                var yahooPara = firstList.AddItem("").AddHyperLink("Yahoo", new Uri("https://yahoo.com"), addStyle: true);
                yahooPara.Color = Color.Purple;
                var yahooLink = yahooPara.Hyperlink;

                document.AddParagraph("Some paragraph separating the lists.");
                document.AddParagraph("Another paragraph.");

                var secondList = document.AddList(WordListStyle.Bulleted);
                secondList.AddItem("").AddHyperLink("DuckDuckGo", new Uri("https://duckduckgo.com")).Hyperlink
                    .CopyFormattingFrom(googleLink);
                secondList.AddItem("").AddHyperLink("Startpage", new Uri("https://startpage.com")).Hyperlink
                    .CopyFormattingFrom(bingLink);
                secondList.AddItem("").AddHyperLink("GitHub", new Uri("https://github.com")).Hyperlink
                    .CopyFormattingFrom(yahooLink);

                document.Save(openWord);
            }
        }
    }
}
