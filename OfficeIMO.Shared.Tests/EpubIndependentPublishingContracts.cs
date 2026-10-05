using OfficeIMO.Epub;
using System.IO.Compression;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubIndependentPublishingContracts {
    [Fact]
    public void IdpfPublicationRetainsRefinedMetadataAndNavigationAcrossEditorialMetadataChanges() {
        var book = LoadIndependent(out var entries);
        var original = book.Read();
        Assert.Equal("Curry, Charles Madison", Assert.Single(original.Metadata, entry => entry.Id == "curry").FileAs);
        Assert.Equal("Clippinger, Erle Elsworth", Assert.Single(original.Metadata, entry => entry.Id == "clippinger").FileAs);
        Assert.Contains(original.Metadata, entry => entry.Refines == "#t2" && entry.Property == "title-type" && entry.Value == "subtitle");
        Assert.Equal(92, original.PageList.Count);
        Assert.Equal("169", original.PageList.First().Label);
        Assert.Equal("260", original.PageList.Last().Label);
        Assert.Equal(2, original.Landmarks.Count);
        Assert.Equal(new[] { "cover", "nav", "s04" }, book.Spine.Select(entry => entry.ManifestId));
        Assert.Contains(Flatten(original.TableOfContents), entry => entry.Label == "Abram S. Isaacs" && entry.Children.Count != 0);

        book.AddContributor("editorial-reviewer", new EpubContributorMetadata { Name = "OfficeIMO test editor", MarcRoles = new[] { "edt" } });
        book.AddCollection("test-series", new EpubCollectionMetadata { Name = "Test editions", Position = new uint[] { 1 } });
        book.SetPrimaryTitle("t1", new EpubTitleMetadata { Text = "Children's Literature — test edition", FileAs = "Children's Literature" });
        book.AddTitle("test-edition", new EpubTitleMetadata { Text = "Editorial test edition", Kind = EpubTitleKind.Edition });
        book.AddSubject("test-subject", new EpubSubjectMetadata { Text = "Education", Language = "en" });
        book.SetPublicationDetails(new EpubPublicationDetails { Description = "An independent-producer editing fixture." });
        EpubWriteResult result = book.Write();
        Assert.False(result.HasLoss);
        using var saved = new ZipArchive(new MemoryStream(result.Bytes), ZipArchiveMode.Read);
        Assert.Equal(entries.Keys.OrderBy(value => value, StringComparer.Ordinal), saved.Entries.Select(entry => entry.FullName).OrderBy(value => value, StringComparer.Ordinal));
        foreach (var entry in entries.Where(entry => entry.Key != book.PackagePath)) {
            using Stream source = saved.GetEntry(entry.Key)!.Open();
            using var bytes = new MemoryStream();
            source.CopyTo(bytes);
            Assert.Equal(entry.Value, bytes.ToArray());
        }
        var reopened = EpubDocument.Load(new MemoryStream(result.Bytes));
        Assert.Equal(original.PageList.Select(PageTarget), reopened.PageList.Select(PageTarget));
        Assert.Equal(Flatten(original.TableOfContents).Select(PageTarget), Flatten(reopened.TableOfContents).Select(PageTarget));
        Assert.Equal("Curry, Charles Madison", Assert.Single(reopened.Metadata, entry => entry.Id == "curry").FileAs);
        Assert.Equal("edt", Assert.Single(reopened.Metadata, entry => entry.Id == "editorial-reviewer").Role);
        Assert.Equal("Children's Literature — test edition", Assert.Single(reopened.Metadata, entry => entry.Id == "t1").Value);
        Assert.Contains(reopened.Metadata, entry => entry.Refines == "#t2" && entry.Property == "title-type" && entry.Value == "subtitle");
    }

    [Fact]
    public void IndependentStaticDerivativeSplitRetainsTextAndRepairsBothNavigationFormats() {
        var book = LoadIndependent(out _);
        var original = book.GetContentXml("s04");
        var html = System.Xml.Linq.XNamespace.Get("http://www.w3.org/1999/xhtml");
        var ncx = System.Xml.Linq.XNamespace.Get("http://www.daisy.org/z3986/2005/ncx/");
        byte[] scripted = book.Write().Bytes;
        Assert.Throws<NotSupportedException>(() => book.SplitChapter("s04", "pgepubid00508", "later", "EPUB/parts/later.xhtml", "Later stories"));
        Assert.Equal(scripted, book.Write().Bytes);
        // Explicit test-only static derivative: retain the independently produced chapter bytes,
        // but remove the TOC disclosure script and reveal its formerly collapsible entries.
        book = LoadIndependent(out _, staticNavigation: true);
        int oldPoints = book.GetContentXml("ncx").Descendants(ncx + "navPoint").Count();
        book.SplitChapter("s04", "pgepubid00508", "later", "EPUB/parts/later.xhtml", "Later stories");
        var reopened = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        Assert.Equal(original.Root!.Element(html + "body")!.Value,
            reopened.GetContentXml("s04").Root!.Element(html + "body")!.Value + reopened.GetContentXml("later").Root!.Element(html + "body")!.Value);
        Assert.Equal(92, reopened.Read().PageList.Count);
        Assert.Contains(reopened.Read().PageList, item => item.Target == "EPUB/parts/later.xhtml");
        Assert.Equal(oldPoints + 1, reopened.GetContentXml("ncx").Descendants(ncx + "navPoint").Count());
        Assert.Equal(new[] { "cover", "nav", "s04", "later" }, reopened.Spine.Select(item => item.ManifestId));
    }

    private static EpubPublication LoadIndependent(out Dictionary<string, byte[]> entries, bool staticNavigation = false) {
        string directory = Path.Combine(AppContext.BaseDirectory, "idpf", "childrens-literature");
        entries = Directory.GetFiles(directory, "*", SearchOption.AllDirectories).ToDictionary(
            path => path.Substring(directory.Length + 1).Replace('\\', '/'), File.ReadAllBytes, StringComparer.Ordinal);
        if (staticNavigation) {
            var html = XNamespace.Get("http://www.w3.org/1999/xhtml");
            var opf = XNamespace.Get("http://www.idpf.org/2007/opf");
            var navigation = XDocument.Parse(System.Text.Encoding.UTF8.GetString(entries["EPUB/nav.xhtml"]), LoadOptions.PreserveWhitespace);
            navigation.Descendants(html + "script").Remove();
            navigation.Descendants().Attributes("hidden").Remove();
            entries["EPUB/nav.xhtml"] = System.Text.Encoding.UTF8.GetBytes(navigation.ToString(SaveOptions.DisableFormatting));
            var package = XDocument.Parse(System.Text.Encoding.UTF8.GetString(entries["EPUB/package.opf"]), LoadOptions.PreserveWhitespace);
            package.Descendants(opf + "item").Single(item => (string?)item.Attribute("id") == "nav").SetAttributeValue("properties", "nav");
            entries["EPUB/package.opf"] = System.Text.Encoding.UTF8.GetBytes(package.ToString(SaveOptions.DisableFormatting));
        }
        using var input = new MemoryStream();
        using (var archive = new ZipArchive(input, ZipArchiveMode.Create, true)) {
            foreach (var entry in entries.OrderBy(entry => entry.Key == "mimetype" ? 0 : 1).ThenBy(entry => entry.Key, StringComparer.Ordinal)) {
                using Stream output = archive.CreateEntry(entry.Key, entry.Key == "mimetype" ? CompressionLevel.NoCompression : CompressionLevel.Optimal).Open();
                output.Write(entry.Value, 0, entry.Value.Length);
            }
        }
        return EpubPublication.Load(new MemoryStream(input.ToArray()));
    }

    private static string PageTarget(EpubNavigationItem item) => item.Label + "|" + item.Target + "|" + item.Fragment;
    private static IEnumerable<EpubNavigationItem> Flatten(IEnumerable<EpubNavigationItem> items) {
        foreach (var item in items) {
            yield return item;
            foreach (var child in Flatten(item.Children)) yield return child;
        }
    }
}
