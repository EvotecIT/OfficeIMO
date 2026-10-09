using OfficeIMO.Epub;

internal static class MergeIdSelectorsFixture {
    internal static EpubPublication Create(bool merge = true) {
        var book = EpubPublication.Create("Merge ID selector qualification", "en", "urn:officeimo:fixture:merge-id-selectors");
        book.AddStylesheet("base", "EPUB/base.css", "html{font:18px/1.5 sans-serif;color:#202020;background:white} body{margin:2rem} p{padding:.5rem}");
        book.AddStylesheet("rules", "EPUB/rules.css", """
            [id^='chapter-'] {border-left:6px solid #185b3a}
            [id$='old'] {font-weight:bold}
            #revised, [id=revised], [id^=rev] {color:#b00020}
            .sample:not(#revised) {color:#185b3a}
            """);
        book.AddChapter("one", "EPUB/one.xhtml", "First", "<h1>First chapter</h1><p>The next section retains its original selector matches after merging.</p>", ["base"]);
        book.AddChapter("two", "EPUB/two.xhtml", "Second", """
            <h1>Second chapter</h1>
            <p id='chapter-old' class='sample'>Renamed: green, bold, with a left border.</p>
            <p id='chapter-stable' class='sample'>Retained: green, normal weight, with a left border.</p>
            <p id='other' class='sample'>Unselected: green, normal weight, without a border.</p>
            """, ["base", "rules"]);
        book.AddChapter("three", "EPUB/three.xhtml", "Unchanged", "<h1>Unchanged chapter</h1><p id='chapter-third' class='sample'>The original shared stylesheet remains available here.</p>", ["base", "rules"]);
        if (merge) book.MergeChapters("one", "two", "second-start", new EpubChapterMergeOptions {
            StylePolicy = EpubChapterMergeStylePolicy.AppendSecondStyles, RewriteChapterSelectors = true,
            SecondChapterIdMap = new Dictionary<string, string> { ["chapter-old"] = "revised" }
        });
        book.SetAccessibilityMetadata(new EpubAccessibilityMetadata {
            AccessModes = ["textual"], SufficientAccessModes = new IReadOnlyList<string>[] { ["textual"] },
            Features = ["structuralNavigation", "tableOfContents"], Hazards = ["none"],
            Summary = "Text describing selector preservation, with heading navigation. No images, flashing, motion or audio."
        });
        return book;
    }
}
