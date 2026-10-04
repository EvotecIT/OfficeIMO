using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Epub;
using OfficeIMO.Html;

namespace OfficeIMO.Workflows;

public sealed partial class BookProject {
    /// <summary>Appends a chapter, retaining the first XHTML chapter's head and resolving its resource references from the new location.</summary>
    public string AddChapter(string title, CancellationToken cancellationToken = default) {
        ArgumentException.ThrowIfNullOrWhiteSpace(title);
        string addedId = string.Empty;
        Mutate(proposed => {
            int next = 1;
            string path;
            do { addedId = "book-chapter-" + next; path = "EPUB/text/book-chapter-" + next++ + ".xhtml"; }
            while (proposed.Manifest.Any(item => item.Id == addedId) || proposed.EntryPaths.Contains(path));
            var previous = proposed.Spine.Select(position => proposed.Manifest.Single(item => item.Id == position.ManifestId))
                .FirstOrDefault(item => string.Equals(item.MediaType, "application/xhtml+xml", StringComparison.OrdinalIgnoreCase));
            XNamespace html = "http://www.w3.org/1999/xhtml";
            proposed.AddChapter(addedId, path, title, new XElement(html + "h1", title).ToString() + new XElement(html + "p", string.Empty).ToString());
            if (previous == null) return;
            XDocument source = proposed.GetContentXml(previous.Id), created = proposed.GetContentXml(addedId);
            XElement head = new XElement(source.Root!.Element(html + "head")!);
            string owner = previous.Reference.ContainerPath!;
            string? baseHref = head.Elements(html + "base").Select(element => (string?)element.Attribute("href")).FirstOrDefault();
            string Rewrite(string value) {
                EpubReference reference = EpubReference.Resolve(owner, baseHref, value);
                if (reference.Kind != EpubReferenceKind.Container) return reference.ResolvedValue ?? value;
                return RelativePackageHref(path, reference.ContainerPath!) +
                    (reference.Fragment == null ? string.Empty : "#" + Uri.EscapeDataString(reference.Fragment));
            }
            head.Elements(html + "base").Remove();
            foreach (XAttribute attribute in head.Descendants().Attributes().Where(attribute => attribute.Name.LocalName is "href" or "src")) attribute.Value = Rewrite(attribute.Value);
            foreach (XElement style in head.Descendants(html + "style")) style.Value = HtmlResourcePipeline.RewriteCssResourceUrls(style.Value, (value, _) => Rewrite(value));
            foreach (XAttribute style in head.Descendants().Attributes("style")) style.Value = HtmlResourcePipeline.RewriteCssResourceUrls(style.Value, (value, _) => Rewrite(value));
            head.Element(html + "title")!.Value = title;
            created.Root!.Element(html + "head")!.ReplaceWith(head);
            foreach (XAttribute attribute in source.Root.Attributes().Where(attribute => attribute.Name.LocalName is "class" or "dir" or "lang"))
                created.Root.SetAttributeValue(attribute.Name, attribute.Value);
            proposed.SetContentXml(addedId, created);
        }, cancellationToken);
        return addedId;
    }
    /// <summary>Removes a chapter and its navigation entries. Remaining content links must still resolve; otherwise the removal is rolled back.</summary>
    public void RemoveChapter(int index, CancellationToken cancellationToken = default) => Mutate(proposed => {
        if (proposed.Spine.Count <= 1) throw new InvalidOperationException("A book must retain at least one chapter.");
        string id = proposed.Spine[index].ManifestId;
        string path = proposed.Manifest.Single(item => item.Id == id).Reference.ContainerPath!;
        EpubDocument reading = ReadCompleteNavigation(proposed, cancellationToken);
        IEnumerable<EpubNavigationEntry> Keep(IEnumerable<EpubNavigationItem> items) {
            foreach (var item in items) {
                if (item.Target == path) { foreach (var child in Keep(item.Children)) yield return child; }
                else yield return new EpubNavigationEntry(item.Label, (item.Target ?? throw new InvalidDataException("Navigation target missing.")) +
                    (item.Fragment == null ? string.Empty : "#" + Uri.EscapeDataString(item.Fragment)), Keep(item.Children), item.SemanticType);
            }
        }
        proposed.SetNavigation(Keep(reading.TableOfContents), Keep(reading.PageList), Keep(reading.Landmarks));
        proposed.RemoveSpineItem(index); proposed.RemoveResource(id);
    }, cancellationToken);

    private static XElement ParseEditorBody(string xml) {
        using var reader = XmlReader.Create(new StringReader(xml), new XmlReaderSettings {
            DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 8L * 1024 * 1024
        });
        XElement body = XElement.Load(reader, LoadOptions.PreserveWhitespace);
        if (body.Name != XName.Get("body", "http://www.w3.org/1999/xhtml"))
            throw new ArgumentException("Chapter editing requires an XHTML body element.", nameof(xml));
        return body;
    }
    private static string RelativePackageHref(string owner, string target) {
        static Uri PackageUri(string path) => new Uri("epub://package/" + string.Join("/", path.Split('/').Select(Uri.EscapeDataString)));
        return PackageUri(owner).MakeRelativeUri(PackageUri(target)).OriginalString;
    }
}
