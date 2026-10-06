using OfficeIMO.Epub;
using System.Xml.Linq;

internal static class MergeSelectorsFixture {
    internal static EpubPublication Create(bool nested = false) {
        var book = EpubPublication.Create(nested ? "Nested merge selector qualification" : "Merge selector qualification", "en",
            nested ? "urn:officeimo:fixture:merge-nested-selectors" : "urn:officeimo:fixture:merge-selectors");
        book.AddStylesheet("shared", "EPUB/shared.css", """
            @import "styles/headings.css";
            html { font:18px/1.5 sans-serif; color:#202020; background:white }
            body { max-width:42rem; margin:2rem auto; padding:0 1rem }
            a { color:#0645ad }
            """);
        book.AddStylesheet("headings", "EPUB/styles/headings.css", nested ? """
            section {
              --heading-data: {#heading};
              & > #heading.lead { color:#123e64; border-left:6px solid currentColor; padding:1rem }
            }
            """ : """
            /* All original chapters share this heading rule. */
            #heading.lead { color:#123e64; border-left:6px solid currentColor; padding:1rem }
            """);
        book.AddChapter("one", "EPUB/one.xhtml", "First chapter",
            "<section aria-labelledby='heading'><h1 id='heading' class='lead'>First chapter</h1><p>This heading remains blue.</p><a href='two.xhtml#heading'>Continue to second chapter</a></section>", ["shared"]);
        book.AddChapter("two", "EPUB/two.xhtml", "Second chapter",
            "<section aria-labelledby='heading'><h1 id='heading' class='lead'>Second chapter</h1><p>This renamed heading is green and retains its border.</p><a href='three.xhtml#heading'>Continue to third chapter</a></section>", ["shared"]);
        book.AddChapter("three", "EPUB/three.xhtml", "Third chapter",
            "<section aria-labelledby='heading'><h1 id='heading' class='lead'>Third chapter</h1><p>This unmerged chapter retains its original blue heading.</p><a href='two.xhtml#heading'>Return to second chapter</a></section>", ["shared"]);
        XNamespace html = "http://www.w3.org/1999/xhtml";
        var second = book.GetContentXml("two");
        second.Root!.Element(html + "head")!.Add(new XElement(html + "style", nested ? """
            section {
              color:#202020;
              @media screen { & > #heading[id='heading'].lead { color:#185b3a } }
              padding-bottom:.5rem;
            }
            """ : "@media screen { #heading[id='heading'].lead { color:#185b3a } }"));
        book.SetContentXml("two", second);
        book.MergeChapters("one", "two", "second-start", new EpubChapterMergeOptions {
            StylePolicy = EpubChapterMergeStylePolicy.AppendSecondStyles, RewriteChapterSelectors = true,
            SecondChapterIdMap = new Dictionary<string, string> { ["heading"] = "second-heading" }
        });
        book.SetAccessibilityMetadata(new EpubAccessibilityMetadata {
            AccessModes = ["textual"], SufficientAccessModes = [new[] { "textual" }],
            Features = ["structuralNavigation", "tableOfContents"], Hazards = ["none"],
            Summary = "Text chapters with headings, labelled sections and repaired navigation. No flashing, motion or audio."
        });
        return book;
    }
}
