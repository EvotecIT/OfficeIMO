using OfficeIMO.Epub;

internal static class MergeContextSelectorsFixture {
    internal static EpubPublication Create(bool merge = true) {
        var book = EpubPublication.Create("Merge selector context qualification", "en", "urn:officeimo:fixture:merge-context-selectors");
        book.AddStylesheet("base", "EPUB/base.css", "html{font:18px/1.5 sans-serif;color:#202020;background:white} body{margin:2rem} p{padding:.5rem}");
        book.AddStylesheet("rules", "EPUB/rules.css", """
            map[name=map] p {color:#185b3a;font-weight:bold;border-left:6px solid #185b3a}
            map[name=second-map] p {color:#b00020}
            output[name=second-map] {display:block;color:#123e64;border-left:6px solid #123e64;padding:.5rem;margin:1rem 0}
            """);
        book.AddChapter("one", "EPUB/one.xhtml", "First", "<h1>First chapter</h1><p>The next section keeps map and output styles separate after their name values converge.</p>", ["base"]);
        foreach (string id in new[] { "two", "three" }) {
            book.AddChapter(id, "EPUB/" + id + ".xhtml", id == "two" ? "Second" : "Unchanged", """
                <h1>Selector context</h1>
                <map id='map' name='map'><p>Map content: green, bold, with a green border.</p></map>
                <output name='second-map' aria-label='Unrelated output'>Output content: blue, normal weight, with a blue border.</output>
                """, ["base", "rules"]);
        }
        if (merge) book.MergeChapters("one", "two", "second-start", new EpubChapterMergeOptions {
            StylePolicy = EpubChapterMergeStylePolicy.AppendSecondStyles, RewriteChapterSelectors = true,
            SecondChapterIdMap = new Dictionary<string, string> { ["map"] = "second-map" }
        });
        book.SetAccessibilityMetadata(new EpubAccessibilityMetadata {
            AccessModes = ["textual"], SufficientAccessModes = new IReadOnlyList<string>[] { ["textual"] },
            Features = ["structuralNavigation", "tableOfContents"], Hazards = ["none"],
            Summary = "Text describing preservation of styles across a chapter merge. The labelled output is static and has no interactive controls. No images, flashing, motion or audio."
        });
        return book;
    }
}
