using System.Threading;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    private byte[] PrepareSplitNavigation(string path, string sourcePath, string newPath, string boundaryId, string title,
        ContentReferenceMap map, CancellationToken token) {
        XDocument navigation = ParseXml(_entries[path], _maximumEntryBytes);
        var entry = new EpubNavigationEntry(title, EncodePath(newPath) + "#" + Uri.EscapeDataString(boundaryId));
        if (navigation.Root?.Name == Html + "html") {
            XElement toc = navigation.Descendants(Html + "nav").Single(element => HasToken((string?)element.Attribute(Ops + "type"), "toc"));
            string? oldBase = navigation.Root!.Element(Html + "head")?.Elements(Html + "base").Attributes("href").FirstOrDefault()?.Value;
            XElement? primary = toc.Descendants(Html + "a").FirstOrDefault(anchor => anchor.Ancestors(Html + "li").Any() &&
                EpubReference.Resolve(path, oldBase, (string?)anchor.Attribute("href") ?? string.Empty).ContainerPath == sourcePath);
            RewriteMovedXml(navigation, path, path, string.Empty, string.Empty, token, map);
            string owner = HtmlContentLinkOwner(navigation, path);
            XElement added = BuildHtmlNodes(new[] { entry }, owner, 0, newPath).Single();
            if (primary != null) {
                if (EpubReference.Resolve(path, oldBase, primary.Attribute("href")!.Value).ContainerPath == newPath)
                    primary.SetAttributeValue("href", RelativeHref(owner, sourcePath));
                primary.Ancestors(Html + "li").First().AddAfterSelf(added);
            } else toc.Element(Html + "ol")!.Add(added);
        } else {
            XElement? primary = navigation.Descendants(Ncx + "navPoint").Elements(Ncx + "content").FirstOrDefault(content =>
                EpubReference.Resolve(path, (string?)content.Attribute("src") ?? string.Empty).ContainerPath == sourcePath);
            RewriteMovedXml(navigation, path, path, string.Empty, string.Empty, token, map);
            int order = 0;
            XElement added = BuildNcxNodes(new[] { entry }, path, 0, ref order, NavigationIds(navigation), newPath).Single();
            if (primary != null) {
                if (EpubReference.Resolve(path, primary.Attribute("src")!.Value).ContainerPath == newPath)
                    primary.SetAttributeValue("src", RelativeHref(path, sourcePath));
                primary.Parent!.AddAfterSelf(added);
            } else navigation.Root!.Element(Ncx + "navMap")!.Add(added);
            NormalizeNcxPlayOrder(navigation, path);
        }
        return SerializeXml(navigation, _maximumEntryBytes);
    }
}
