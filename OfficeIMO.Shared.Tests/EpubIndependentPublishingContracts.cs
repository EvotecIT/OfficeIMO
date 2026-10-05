using OfficeIMO.Epub;
using System.IO.Compression;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubIndependentPublishingContracts {
    [Fact]
    public void IdpfPublicationRetainsRefinedMetadataAndNavigationAcrossEditorialMetadataChanges() {
        string directory = Path.Combine(AppContext.BaseDirectory, "idpf", "childrens-literature");
        var entries = Directory.GetFiles(directory, "*", SearchOption.AllDirectories).ToDictionary(
            path => path.Substring(directory.Length + 1).Replace('\\', '/'), File.ReadAllBytes, StringComparer.Ordinal);
        using var input = new MemoryStream();
        using (var archive = new ZipArchive(input, ZipArchiveMode.Create, true)) {
            foreach (var entry in entries.OrderBy(entry => entry.Key == "mimetype" ? 0 : 1).ThenBy(entry => entry.Key, StringComparer.Ordinal)) {
                using Stream output = archive.CreateEntry(entry.Key, entry.Key == "mimetype" ? CompressionLevel.NoCompression : CompressionLevel.Optimal).Open();
                output.Write(entry.Value, 0, entry.Value.Length);
            }
        }
        var book = EpubPublication.Load(new MemoryStream(input.ToArray()));
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

    private static string PageTarget(EpubNavigationItem item) => item.Label + "|" + item.Target + "|" + item.Fragment;
    private static IEnumerable<EpubNavigationItem> Flatten(IEnumerable<EpubNavigationItem> items) {
        foreach (var item in items) {
            yield return item;
            foreach (var child in Flatten(item.Children)) yield return child;
        }
    }
}
