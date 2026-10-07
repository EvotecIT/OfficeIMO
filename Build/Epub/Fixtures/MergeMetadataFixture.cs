using OfficeIMO.Epub;
using System.IO.Compression;
using System.Xml.Linq;

internal static class MergeMetadataFixture {
    internal static EpubPublication Create(bool merged) {
        XNamespace opf = "http://www.idpf.org/2007/opf";
        var source = EpubPublication.Create("Package refinement merge qualification", "en", "urn:officeimo:fixture:merge-metadata");
        source.AddChapter("one", "EPUB/one.xhtml", "First chapter", "<h1>First chapter</h1><p>The package records a description for each chapter.</p>");
        source.AddChapter("two", "EPUB/two.xhtml", "Second chapter", "<h1>Second chapter</h1><p>After the explicit merge, both descriptions apply to the combined chapter.</p>");
        source.SetAccessibilityMetadata(new EpubAccessibilityMetadata {
            AccessModes = ["textual"], SufficientAccessModes = [new[] { "textual" }],
            Features = ["structuralNavigation", "tableOfContents"], Hazards = ["none"],
            Summary = "Two text chapters with headings and table-of-contents navigation. No images, flashing, motion or audio."
        });
        XDocument package = source.GetPackageXml();
        package.Root!.Element(opf + "spine")!.Elements().Last().SetAttributeValue("id", "second-position");
        package.Root.Element(opf + "metadata")!.Add(
            new XElement(opf + "meta", new XAttribute("id", "first-description"), new XAttribute("property", "dcterms:description"), new XAttribute("refines", "#one"), "First chapter description"),
            new XElement(opf + "meta", new XAttribute("id", "second-description"), new XAttribute("property", "dcterms:description"), new XAttribute("refines", "#two"), "Second chapter description"),
            new XElement(opf + "meta", new XAttribute("property", "dcterms:description"), new XAttribute("refines", "#second-description"), "Annotation on the second description"),
            new XElement(opf + "meta", new XAttribute("property", "dcterms:description"), new XAttribute("refines", "#second-position"), "Reading-position description"),
            new XElement(opf + "link", new XAttribute("rel", "dcterms:description"), new XAttribute("href", "https://example.org/description.html"), new XAttribute("media-type", "text/html"), new XAttribute("refines", "#two")));
        using var stream = new MemoryStream();
        stream.Write(source.Write().Bytes);
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Update, true)) {
            zip.GetEntry(source.PackagePath)!.Delete();
            using var output = zip.CreateEntry(source.PackagePath).Open();
            package.Save(output);
        }
        var book = EpubPublication.Load(new MemoryStream(stream.ToArray()));
        if (merged) book.MergeChapters("one", "two", "second-start", new EpubChapterMergeOptions { RetargetPackageRefinements = true });
        return book;
    }
}
