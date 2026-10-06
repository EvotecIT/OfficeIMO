using OfficeIMO.Epub;
using OfficeIMO.Html;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using System.Xml.Linq;

if (args.Length != 1) throw new ArgumentException("Supply one new, task-owned output directory.");
string outputDirectory = Path.GetFullPath(args[0]);
if (Directory.Exists(outputDirectory) || File.Exists(outputDirectory)) throw new IOException("Output already exists; use a new directory.");
Directory.CreateDirectory(outputDirectory);
var evidence = new List<object>();
foreach (EpubTypographyProfile profile in Enum.GetValues<EpubTypographyProfile>()) {
    string name = profile.ToString().ToLowerInvariant();
    var imported = EpubManuscript.ImportHtml(HtmlConversionDocument.Parse(FixtureContent.Manuscript), new EpubManuscriptOptions {
        Title = profile + " typography", TypographyProfile = profile, ChapterHeadingLevel = 0,
        Identifier = "urn:officeimo:fixture:typography:" + name
    });
    imported.Report.RequireNoLoss();
    imported.Publication.SetAccessibilityMetadata(new EpubAccessibilityMetadata {
        AccessModes = new[] { "textual", "visual" }, SufficientAccessModes = new IReadOnlyList<string>[] { new[] { "textual" } },
        Features = new[] { "structuralNavigation", "tableOfContents", "alternativeText" }, Hazards = new[] { "none" },
        Summary = "Prose, Arabic and Japanese text, a table, wrapping code, and a labelled vector. No flashing, motion, or audio."
    });
    byte[] bytes = imported.Publication.Write(new EpubWriteOptions { ModifiedAt = new DateTimeOffset(2026, 1, 1, 0, 0, 0, TimeSpan.Zero) }).Bytes;
    string epub = Path.Combine(outputDirectory, name + ".epub");
    File.WriteAllBytes(epub, bytes);
    string expanded = Path.Combine(outputDirectory, name);
    ZipFile.ExtractToDirectory(epub, expanded);
    var previewPaths = new List<string>();
    foreach (string mode in new[] { "light", "dark", "large" }) {
        string chapterPath = imported.Publication.Manifest.Single(item => item.Id == "chapter-1").Reference.ContainerPath!;
        string source = Path.Combine(expanded, chapterPath.Replace('/', Path.DirectorySeparatorChar));
        XDocument preview = XDocument.Load(source, LoadOptions.PreserveWhitespace);
        preview.DocumentType?.Remove();
        preview.AddFirst(new XDocumentType("html", null, null, null));
        XNamespace html = "http://www.w3.org/1999/xhtml";
        string readerCss = "reader-" + mode + ".css";
        preview.Root!.Element(html + "head")!.Add(new XElement(html + "link", new XAttribute("rel", "stylesheet"), new XAttribute("href", readerCss)));
        preview.Root.Element(html + "head")!.Add(new XElement(html + "meta", new XAttribute("name", "viewport"), new XAttribute("content", "width=device-width, initial-scale=1")));
        string previewPath = Path.Combine(Path.GetDirectoryName(source)!, "preview-" + mode + ".html");
        preview.Save(previewPath);
        File.WriteAllText(Path.Combine(Path.GetDirectoryName(source)!, readerCss),
            mode == "dark" ? "html{font:16px serif;color:#eee;background:#171717}body{margin:1em}a{color:#9cc8ff}" :
            mode == "large" ? "html{font:32px sans-serif;color:#111;background:#fff}body{margin:.5em}" :
            "html{font:16px serif;color:#111;background:#fff}body{margin:1em}");
        previewPaths.Add(Path.GetRelativePath(outputDirectory, previewPath).Replace('\\', '/'));
    }
    evidence.Add(new { profile = name, epub = name + ".epub", sha256 = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant(), previews = previewPaths,
        nativePreflight = InspectFixture(imported.Publication) });
}
WriteFixture("read-aloud", MediaOverlayFixture.Create());
WriteFixture("read-aloud-revised", MediaOverlayFixture.Revised());
WriteFixture("read-aloud-nested", MediaOverlayFixture.Nested());
WriteFixture("read-aloud-nested-revised", MediaOverlayFixture.Nested(true));
WriteFixture("read-aloud-svg", SvgNarrationFixture.Create(false));
WriteFixture("read-aloud-inline-svg", SvgNarrationFixture.Create(true));
WriteFixture("fixed-layout-svg", SvgPageFixture.Create());
WriteFixture("fixed-layout-escaped-id", FixedLayoutFixture.EscapedIdentifier());
WriteFixture("fixed-layout-item-overrides", FixedLayoutFixture.Create(false, false));
WriteFixture("fixed-layout-ltr", FixedLayoutFixture.Create(false));
WriteFixture("fixed-layout-rtl", FixedLayoutFixture.Create(true));
WriteFixture("glossary", GlossaryFixture.Create());
WriteFixture("bibliography", BibliographyFixture.Create());
WriteFixture("index", IndexFixture.Create());
var renamed = IndexFixture.Create();
renamed.RenameResource("source", "EPUB/revised/part one.xhtml");
renamed.RenameResource("index", "EPUB/revised/back/index.xhtml");
renamed.RenameResource("navigation", "EPUB/revised/navigation.xhtml");
WriteFixture("renamed-index", renamed);
WriteFixture("split-chapter", SplitFixture.Create());
var merged = SplitFixture.Create();
merged.MergeChapters("source", "second", "second-start");
WriteFixture("merged-chapters", merged);
WriteFixture("merged-styles", MergeStyleFixture.Create());
File.WriteAllText(Path.Combine(outputDirectory, "manifest.json"), JsonSerializer.Serialize(new {
    publications = evidence, previewBoundary = "Browser previews add simulated reader theme/font CSS. They are not EPUB reading-system acceptance."
}, new JsonSerializerOptions { WriteIndented = true }));
Console.WriteLine(Path.Combine(outputDirectory, "manifest.json"));

void WriteFixture(string name, EpubPublication publication) {
    byte[] bytes = publication.Write(new EpubWriteOptions {
        ModifiedAt = new DateTimeOffset(2026, 1, 1, 0, 0, 0, TimeSpan.Zero)
    }).Bytes;
    File.WriteAllBytes(Path.Combine(outputDirectory, name + ".epub"), bytes);
    evidence.Add(new { fixture = name, epub = name + ".epub", sha256 = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant(),
        nativePreflight = InspectFixture(publication) });
}

object InspectFixture(EpubPublication publication) {
    EpubPreflightReport report = publication.Preflight();
    if (report.HasErrors) throw new InvalidDataException("Fixture failed native preflight: " + publication.Title + ": " +
        string.Join("; ", report.Checks.SelectMany(check => check.Diagnostics).Where(d => d.Severity == EpubDiagnosticSeverity.Error).Select(d => d.Message)));
    return new { report.HasErrors, report.HasUncheckedItems,
        checks = report.Checks.Select(check => new { check.Code, status = check.Status.ToString(), check.Diagnostics }).ToArray() };
}

internal static class FixtureContent {
    internal const string Manuscript = """
    <!DOCTYPE html><html lang="en"><head><title>Typography qualification</title></head><body>
    <h1>Reading across settings</h1>
    <p id="first-paragraph">A book should remain readable when its reader changes the font, theme, or available width. This paragraph begins a short passage.</p>
    <p id="second-paragraph">A second paragraph makes the prose indentation visible. Reader preferences should retain their effect without clipping the text.</p>
    <h2>Technical content</h2>
    <table><caption>Small-screen resource limits</caption><thead><tr><th scope="col">Resource</th><th scope="col">Policy</th><th scope="col">Example</th></tr></thead>
    <tbody><tr><th scope="row">Chapter</th><td>Bounded content</td><td>chapter-with-a-deliberately-long-name.xhtml</td></tr>
    <tr><th scope="row">Image</th><td>Fit available width</td><td>1200-pixel source</td></tr></tbody></table>
    <pre><code>publication.AddChapter("chapter-with-a-long-identifier", "EPUB/text/chapter.xhtml", "Chapter", content);</code></pre>
    <p id="long-link"><a href="#long-link">https://example.org/a-deliberately-long-unbroken-reference-for-small-screen-wrapping</a></p>
    <figure><svg xmlns="http://www.w3.org/2000/svg" role="img" aria-labelledby="diagram-title" viewBox="0 0 1200 120" width="1200" height="120">
    <title id="diagram-title">A plain blue rectangle illustrating flexible image width</title><rect width="1200" height="120" fill="#4472c4"/></svg>
    <figcaption>The vector scales to the available width.</figcaption></figure>
    <h2>Text direction and scripts</h2>
    <section lang="ar" dir="rtl"><h3>قراءة النص</h3><p>هذا نص عربي لاختبار اتجاه القراءة والمسافات بين الفقرات.</p><p>يمكن للقارئ تغيير حجم الخط ولون الخلفية.</p></section>
    <section lang="ja"><h3>日本語の文章</h3><p>これは日本語の表示を確認するための文章です。文字の大きさを変えても、文章が読みやすく表示されることを確認します。</p></section>
    </body></html>
    """;
}
