using OfficeIMO.Epub;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using System.Xml.Linq;

internal static class EmbeddedFontFixture {
    private static readonly string[] Fonts = ["Carlito-Regular.ttf", "NotoSansArabic-Regular.ttf", "NotoSansJP-OfficeIMO-Common.ttf"];
    private static readonly string[] Licenses = ["OFL-Carlito.txt", "OFL-Noto.txt", "OFL-NotoCJK.txt"];

    internal static EpubPublication Create() {
        var book = EpubPublication.Create("Embedded multilingual fonts", "en", "urn:officeimo:fixture:embedded-fonts");
        foreach (string font in Fonts) book.AddResource(font, "EPUB/fonts/" + font, "font/ttf", Read(font));
        foreach (string license in Licenses) book.AddResource(license, "EPUB/licenses/" + license, "text/plain", Read(license));
        book.AddStylesheet("style", "EPUB/style.css", EpubTypography.CreateStylesheet(EpubTypographyProfile.Technical) + """

            @font-face { font-family: 'Fixture Latin'; src: url('fonts/Carlito-Regular.ttf') format('truetype'); font-weight:400; font-style:normal; }
            @font-face { font-family: 'Fixture Arabic'; src: url('fonts/NotoSansArabic-Regular.ttf') format('truetype'); font-weight:400; font-style:normal; }
            @font-face { font-family: 'Fixture Japanese'; src: url('fonts/NotoSansJP-OfficeIMO-Common.ttf') format('truetype'); font-weight:400; font-style:normal; }
            body { font-family:'Fixture Latin',sans-serif; }
            [lang='ar'] { font-family:'Fixture Arabic',sans-serif; }
            [lang='ja'] { font-family:'Fixture Japanese',sans-serif; }
            """);
        book.AddChapter("chapter", "EPUB/text/chapter.xhtml", "Embedded fonts", """
            <h1>Embedded fonts</h1>
            <p id="latin">Readable text should survive a change of font, size or theme. This Latin paragraph uses the embedded Carlito font.</p>
            <section lang="ar" xml:lang="ar" dir="rtl"><h2>قراءة النص</h2>
            <p id="arabic">هذا نص عربي لاختبار اتجاه القراءة والمسافات بين الفقرات.</p>
            <p>رقم الإصدار <bdi dir="ltr">EPUB 3.3</bdi> ورقم الصفحة <bdi dir="ltr">123</bdi>.</p></section>
            <section lang="ja" xml:lang="ja"><h2>日本語の文章</h2>
            <p id="japanese">これは日本語の表示を確認するための文章です。文字の大きさを変えても、文章が読みやすく表示されることを確認します。</p></section>
            <h2>Technical content</h2>
            <table><caption>Language and direction</caption><thead><tr><th scope="col">Language</th><th scope="col">Direction</th></tr></thead>
            <tbody><tr><th scope="row">Arabic</th><td>Right to left</td></tr><tr><th scope="row">Japanese</th><td>Horizontal left to right</td></tr></tbody></table>
            <pre><code>publication.AddChapter("a-long-chapter-identifier", "EPUB/text/chapter.xhtml", "Chapter", content);</code></pre>
            <p><a href="#arabic">Return to the Arabic paragraph</a></p>
            """, ["style"]);
        book.SetAccessibilityMetadata(new EpubAccessibilityMetadata {
            AccessModes = ["textual"], SufficientAccessModes = [new[] { "textual" }],
            Features = ["structuralNavigation", "tableOfContents"], Hazards = ["none"],
            Summary = "Latin, Arabic and Japanese passages with explicit languages, an Arabic direction scope, isolated Latin identifiers, a table and code. No audio, flashing or motion."
        });
        return book;
    }

    internal static void WriteEvidence(string output, byte[] bytes) {
        string expanded = Path.Combine(output, "embedded-fonts");
        using var archive = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        using var provenance = JsonDocument.Parse(Read("font-pack.json"));
        var fonts = new List<object>();
        foreach (string font in Fonts) {
            using var entry = archive.GetEntry("EPUB/fonts/" + font)!.Open(); using var copy = new MemoryStream(); entry.CopyTo(copy);
            byte[] original = Read(font);
            if (!original.SequenceEqual(copy.ToArray())) throw new InvalidDataException("Embedded font bytes changed: " + font);
            string hash = Convert.ToHexString(SHA256.HashData(original));
            string expected = provenance.RootElement.GetProperty("fonts").EnumerateArray().SelectMany(item => item.GetProperty("files").EnumerateArray())
                .Single(item => item.GetProperty("name").GetString() == font).GetProperty("sha256").GetString()!;
            if (!hash.Equals(expected, StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("Fixture font differs from its recorded provenance: " + font);
            fonts.Add(new { file = font, sha256 = hash, bytes = original.Length });
        }
        archive.ExtractToDirectory(expanded);
        var previews = new List<string>(); XNamespace html = "http://www.w3.org/1999/xhtml";
        foreach (string mode in new[] { "light", "dark", "large", "override" }) {
            var document = XDocument.Load(Path.Combine(expanded, "EPUB/text/chapter.xhtml"), LoadOptions.PreserveWhitespace);
            document.DocumentType?.Remove(); document.AddFirst(new XDocumentType("html", null, null, null));
            string css = "html{font-size:" + (mode == "large" ? "32" : "18") + "px}body{margin:1rem}" +
                (mode == "dark" ? "html{color:#eee;background:#171717}a{color:#9cc8ff}" : "html{color:#111;background:#fff}") +
                (mode == "override" ? "body,body *{font-family:serif!important}" : "");
            document.Root!.Element(html + "head")!.Add(new XElement(html + "meta", new XAttribute("name", "viewport"), new XAttribute("content", "width=device-width, initial-scale=1")),
                new XElement(html + "style", css));
            string path = Path.Combine(expanded, "EPUB/text/preview-" + mode + ".html"); document.Save(path);
            previews.Add(Path.GetRelativePath(output, path).Replace('\\', '/'));
        }
        File.WriteAllText(Path.Combine(output, "embedded-fonts-evidence.json"), JsonSerializer.Serialize(new {
            epubSha256 = Convert.ToHexString(SHA256.HashData(bytes)), fonts, previews,
            provenance = provenance.RootElement,
            boundary = "Bundled font bytes and package structure only; previews simulate reader settings. Native font selection, pagination and assistive technology require separate evidence."
        }, new JsonSerializerOptions { WriteIndented = true }));
    }

    private static byte[] Read(string name) {
        using var stream = typeof(EmbeddedFontFixture).Assembly.GetManifestResourceStream("OfficeIMO.Epub.Fixtures." + name)
            ?? throw new InvalidDataException("Missing licensed fixture asset: " + name);
        using var copy = new MemoryStream(); stream.CopyTo(copy); return copy.ToArray();
    }
}
