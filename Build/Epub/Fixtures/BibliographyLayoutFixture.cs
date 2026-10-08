using OfficeIMO.Bibliography;
using OfficeIMO.Epub;
using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;

internal static class BibliographyLayoutFixture {
    internal static EpubPublication Create(string mode) {
        string attributes = mode switch {
            "hanging" => "hanging-indent='true'",
            "flush" => "second-field-align='flush'",
            "margin" => "second-field-align='margin'",
            _ => throw new ArgumentOutOfRangeException(nameof(mode))
        };
        var sources = BibliographyDocument.Parse("""
            [{"id":"one","type":"book","title":"Accessible publishing and reference navigation across narrow and wide reading environments","author":[{"family":"Example","given":"Ada"}],"issued":{"date-parts":[[2026]]}},
             {"id":"two","type":"book","title":"Typography and editorial workflows for long scholarly publications with multiple sources","author":[{"family":"Writer","given":"Alex"}],"issued":{"date-parts":[[2025]]}}]
            """, BibliographyFormat.CslJson).Document;
        var style = CslStyle.Parse("<style xmlns='http://purl.org/net/xbiblio/csl' version='1.0' class='in-text'>" +
            "<citation><layout prefix='[' suffix=']'><text variable='citation-number'/></layout></citation>" +
            "<bibliography " + attributes + " line-spacing='2' entry-spacing='1'><layout>" +
            (mode == "hanging" ? "" : "<text variable='citation-number' suffix='.'/>") +
            "<group delimiter='. '><names variable='author'><name/></names><text variable='title' font-style='italic'/>" +
            "<date variable='issued'><date-part name='year'/></date></group></layout></bibliography></style>");
        var citations = new[] { new CslCitation("cite-one"), new CslCitation("cite-two") };
        citations[0].Items.Add(new CslCitationItem("one")); citations[1].Items.Add(new CslCitationItem("two"));
        var result = new CslProcessor(sources, style, new CslRenderOptions { OutputFormat = CslOutputFormat.Html }).Render(citations);
        var layout = result.BibliographyLayout ?? throw new InvalidDataException("Missing CSL layout.");
        string fieldWidth = Math.Max(3, layout.MaximumLeftMarginCharacters + 1).ToString(CultureInfo.InvariantCulture) + "ch";
        string css = "body{margin:1em;font-family:serif;overflow-wrap:anywhere}h1{font-size:1.5em}ol{list-style:none;padding:0}" +
            "li{overflow-wrap:anywhere;line-height:" + layout.LineSpacing + ";margin-bottom:" + (layout.EntrySpacing * layout.LineSpacing) + "em}" +
            "li::after{content:'';display:block;clear:both}.csl-left-margin{float:left;width:" + fieldWidth + "}";
        if (layout.SecondFieldAlignment == CslSecondFieldAlignment.Margin) css += "main{margin-left:" + fieldWidth + "}";
        css += layout.HangingIndent ? "li{padding-left:1.5em;text-indent:-1.5em}" :
            layout.SecondFieldAlignment == CslSecondFieldAlignment.Flush ? ".csl-right-inline{margin-left:" + fieldWidth + "}" :
            ".csl-left-margin{margin-left:-" + fieldWidth + "}.csl-right-inline{margin-left:0}";
        var book = EpubPublication.Create("Bibliography " + mode + " qualification", "en", "urn:officeimo:fixture:bibliography-layout:" + mode);
        book.AddStylesheet("style", "EPUB/style.css", css);
        book.AddChapter("source", "EPUB/text/chapter.xhtml", "Sources", "<main><h1>Two sources for one passage</h1><p id='passage'>" +
            "This passage draws on both works. Follow each source separately: " +
            string.Join(", ", result.Citations.Select(c => "<a id='" + c.Key + "'>" + c.Content + "</a>")) + ".</p></main>", ["style"]);
        book.AddChapter("references", "EPUB/back/references.xhtml", "References", "<main><section aria-labelledby='heading'><h1 id='heading'>References</h1><ol id='entries' role='list'/></section></main>", ["style"]);
        book.SetDocumentMatter("references", EpubDocumentMatter.BackMatter);
        foreach (var entry in result.Bibliography) book.AddBibliographyEntry("references", "entries", entry.Key, entry.Content);
        book.LinkBibliographyEntry("source", "cite-one", "references", "one", "Return to source one citation");
        book.LinkBibliographyEntry("source", "cite-two", "references", "two", "Return to source two citation");
        book.SetAccessibilityMetadata(new() {
            AccessModes = ["textual"], SufficientAccessModes = [new[] { "textual" }],
            Features = ["structuralNavigation", "tableOfContents"], Hazards = ["none"],
            Summary = "Two formatted references with separate source links and labelled return links. No images or audio."
        });
        return book;
    }

    internal static void WritePreviews(string output, string name, byte[] bytes) {
        foreach (string mode in new[] { "light", "large", "dark" }) {
            string directory = Path.Combine(output, name + "-" + mode);
            using var archive = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
            archive.ExtractToDirectory(directory);
            foreach (string path in Directory.GetFiles(directory, "*.xhtml", SearchOption.AllDirectories)) {
                var document = XDocument.Load(path, LoadOptions.PreserveWhitespace);
                document.DocumentType?.Remove();
                document.AddFirst(new XDocumentType("html", null, null, null));
                document.Root!.SetAttributeValue("lang", (string?)document.Root.Attribute(XNamespace.Xml + "lang") ?? "en");
                XNamespace html = "http://www.w3.org/1999/xhtml";
                document.Root!.Element(html + "head")!.Add(new XElement(html + "meta", new XAttribute("name", "viewport"),
                    new XAttribute("content", "width=device-width, initial-scale=1")), new XElement(html + "style",
                    "html{font-size:" + (mode == "large" ? "32" : "18") + "px}" +
                    (mode == "dark" ? "html{color:#eee;background:#171717}a{color:#9cc8ff}" : "html{color:#111;background:#fff}")));
                foreach (var link in document.Descendants(html + "a")) {
                    var href = link.Attribute("href");
                    if (href != null) href.Value = href.Value.Replace(".xhtml", ".html", StringComparison.Ordinal);
                }
                document.Save(Path.ChangeExtension(path, ".html"));
            }
        }
    }
}
