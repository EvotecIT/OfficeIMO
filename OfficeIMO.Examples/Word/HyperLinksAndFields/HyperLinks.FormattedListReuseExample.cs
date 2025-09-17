using System;
using System.IO;
using OfficeIMO.Examples.Utils;
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
                var googleLink = Guard.NotNull(googlePara.Hyperlink, "Expected Google hyperlink to be created.");

                var bingPara = firstList.AddItem("").AddHyperLink("Bing", new Uri("https://bing.com"), addStyle: true);
                bingPara.Italic = true;
                var bingLink = Guard.NotNull(bingPara.Hyperlink, "Expected Bing hyperlink to be created.");

                var yahooPara = firstList.AddItem("").AddHyperLink("Yahoo", new Uri("https://yahoo.com"), addStyle: true);
                yahooPara.Color = Color.Purple;
                var yahooLink = Guard.NotNull(yahooPara.Hyperlink, "Expected Yahoo hyperlink to be created.");

                document.AddParagraph("Some paragraph separating the lists.");
                document.AddParagraph("Another paragraph.");

                var secondList = document.AddList(WordListStyle.Bulleted);
                var duckLink = Guard.NotNull(secondList.AddItem("").AddHyperLink("DuckDuckGo", new Uri("https://duckduckgo.com"))
                    .Hyperlink, "Expected DuckDuckGo hyperlink to be created.");
                duckLink.CopyFormattingFrom(googleLink);

                var startPageLink = Guard.NotNull(secondList.AddItem("").AddHyperLink("Startpage", new Uri("https://startpage.com"))
                    .Hyperlink, "Expected Startpage hyperlink to be created.");
                startPageLink.CopyFormattingFrom(bingLink);

                var gitHubLink = Guard.NotNull(secondList.AddItem("").AddHyperLink("GitHub", new Uri("https://github.com"))
                    .Hyperlink, "Expected GitHub hyperlink to be created.");
                gitHubLink.CopyFormattingFrom(yahooLink);

                document.Save(openWord);
            }
        }
    }
}
