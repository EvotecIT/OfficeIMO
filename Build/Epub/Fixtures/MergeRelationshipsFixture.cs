using OfficeIMO.Epub;
using System.Xml.Linq;

internal static class MergeRelationshipsFixture {
    internal static EpubPublication Create() {
        var book = EpubPublication.Create("Merge relationship selector qualification", "en", "urn:officeimo:fixture:merge-relationships");
        book.AddStylesheet("shared", "EPUB/shared.css", """
            @import "styles/relationships.css";
            html { font:18px/1.5 sans-serif; color:#202020; background:white }
            body { max-width:42rem; margin:2rem auto; padding:0 1rem }
            section { margin-bottom:2rem; padding:1rem }
            table { border-collapse:collapse; width:100% }
            th, td { border:1px solid #777; padding:.5rem; text-align:left }
            a { color:#0645ad }
            """);
        book.AddStylesheet("relationships", "EPUB/styles/relationships.css", """
            section[aria-labelledby~=heading] { color:#123e64; border-left:6px solid currentColor }
            [itemref=heading] { border-top:3px solid #805000 }
            td[headers~=column] { font-weight:bold; color:#123e64 }
            """);
        Add("one", "First", "two.xhtml#heading", "Continue to second chapter");
        Add("two", "Second", "three.xhtml#heading", "Continue to third chapter");
        Add("three", "Third", "two.xhtml#heading", "Return to second chapter");
        XNamespace html = "http://www.w3.org/1999/xhtml";
        var second = book.GetContentXml("two");
        second.Root!.Element(html + "head")!.Add(new XElement(html + "style", """
            section[aria-labelledby='heading'] { color:#185b3a }
            td[headers='column'] { color:#185b3a }
            """));
        book.SetContentXml("two", second);
        book.MergeChapters("one", "two", "second-start", new EpubChapterMergeOptions {
            StylePolicy = EpubChapterMergeStylePolicy.AppendSecondStyles, RewriteSecondChapterIdSelectors = true,
            SecondChapterIdMap = new Dictionary<string, string> { ["heading"] = "second-heading", ["column"] = "second-column" }
        });
        book.SetAccessibilityMetadata(new EpubAccessibilityMetadata {
            AccessModes = ["textual"], SufficientAccessModes = [new[] { "textual" }],
            Features = ["structuralNavigation", "tableOfContents"], Hazards = ["none"],
            Summary = "Labelled sections and tables with repaired document-local references and navigation. No flashing, motion or audio."
        });
        return book;

        void Add(string id, string title, string href, string linkText) {
            book.AddChapter(id, "EPUB/" + id + ".xhtml", title + " chapter",
                "<section aria-labelledby='heading' itemscope='' itemref='heading'><h1 id='heading'>" + title + " chapter</h1>" +
                "<p>The first and third sections are blue; the second is green. All retain their left and top borders.</p>" +
                "<table><caption>Chapter data</caption><thead><tr><th id='column' scope='col'>Value</th></tr></thead>" +
                "<tbody><tr><td headers='column'>" + title + " value</td></tr></tbody></table>" +
                "<p><a href='" + href + "'>" + linkText + "</a></p></section>", ["shared"]);
        }
    }
}
