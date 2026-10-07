using OfficeIMO.Epub;
using System.Xml.Linq;

internal static class MergeLanguageFixture {
    internal static EpubPublication Create() {
        var book = EpubPublication.Create("Multilingual chapter merge qualification", "en", "urn:officeimo:fixture:merge-language");
        book.AddChapter("one", "EPUB/one.xhtml", "English chapter", "<h1 id='english'>English chapter</h1><p>This English paragraph precedes an Arabic chapter.</p><p><a href='two.xhtml#arabic'>Continue to Arabic</a></p>");
        book.AddChapter("two", "EPUB/two.xhtml", "الفصل العربي", "<h1 id='arabic'>الفصل العربي</h1><p>هذه فقرة باللغة العربية.</p><p lang='fr' xml:lang='fr' dir='ltr'>Bonjour, cette phrase reste en français.</p><p><a href='one.xhtml#english'>العودة إلى الفصل الإنجليزي</a></p>");
        var second = book.GetContentXml("two");
        second.Root!.SetAttributeValue("lang", "ar"); second.Root.SetAttributeValue(XNamespace.Xml + "lang", "ar"); second.Root.SetAttributeValue("dir", "rtl");
        book.SetContentXml("two", second);
        book.MergeChapters("one", "two", "arabic-start", new EpubChapterMergeOptions { PreserveSecondChapterLanguageAndDirection = true });
        book.SetAccessibilityMetadata(new EpubAccessibilityMetadata {
            AccessModes = ["textual"], SufficientAccessModes = [new[] { "textual" }],
            Features = ["structuralNavigation", "tableOfContents"], Hazards = ["none"],
            Summary = "English and Arabic chapters with an explicitly marked French passage and reciprocal navigation links. No flashing, motion or audio."
        });
        return book;
    }
}
