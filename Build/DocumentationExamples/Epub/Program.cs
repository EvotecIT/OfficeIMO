#region QuickStart
using OfficeIMO.Epub;

EpubPublication publication = EpubPublication.Create("Field guide", "en");
publication.Creator = "Operations";
publication.AddStylesheet("style", "EPUB/style.css",
    "body { font-family: sans-serif; } p { line-height: 1.4; }");
publication.AddChapter("opening", "EPUB/opening.xhtml", "Opening",
    "<h1 id='opening'>Opening</h1><p>A chapter with <em>formatted</em> text.</p>",
    new[] { "style" });
publication.SetNavigation(new[] {
    new EpubNavigationEntry("Opening", "EPUB/opening.xhtml#opening")
});
publication.Save("Field-guide.epub").ThrowIfLoss();

EpubDocument book = EpubDocument.Load("Field-guide.epub");
foreach (EpubChapter chapter in book.Chapters) Console.WriteLine($"{chapter.Order}. {chapter.Title}");
#endregion
