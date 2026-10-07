using System.Threading;
using OfficeIMO.Epub;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubEditionComparisonContracts {
    [Theory]
    [InlineData(EpubVersion.Epub2)]
    [InlineData(EpubVersion.Epub3)]
    public void UnchangedRoundtripHasNoChangesAndComparisonDoesNotMutateEitherEdition(EpubVersion version) {
        var before = Book(version); byte[] bytes = before.Write().Bytes;
        var after = EpubPublication.Load(new MemoryStream(bytes));
        var report = before.CompareTo(after);
        Assert.True(!report.HasChanges, $"metadata={report.MetadataChanged}; spine={report.ReadingOrderChanged}; package={report.PackageStructureChanged}; resources={string.Join(",", report.Resources.Select(r => r.ManifestId + ":" + r.Kind))}");
        Assert.Equal(bytes, before.Write().Bytes); Assert.Equal(bytes, after.Write().Bytes);
    }
    [Fact]
    public void TextMarkupMetadataAndOrderAreReportedSeparately() {
        var before = Book(); var after = EpubPublication.Load(new MemoryStream(before.Write().Bytes));
        after.Title = "Second edition"; after.MoveSpineItem(0, 1);
        var xml = after.GetContentXml("one"); xml.Descendants().Single(e => (string?)e.Attribute("id") == "paragraph").Value = "Revised";
        after.SetContentXml("one", xml);
        var report = before.CompareTo(after);
        Assert.True(report.MetadataChanged); Assert.True(report.ReadingOrderChanged);
        var paragraph = report.TextChanges.Single(item => item.Locator == "#paragraph");
        Assert.Equal("Original", paragraph.PreviousText); Assert.Equal("Revised", paragraph.CurrentText);
        var change = Assert.Single(report.Resources);
        Assert.Equal("one", change.ManifestId);
        Assert.Equal(EpubEditionChangeKind.Xml | EpubEditionChangeKind.Text, change.Kind);
        xml.Descendants().Single(e => (string?)e.Attribute("id") == "paragraph").Value = "Original";
        xml.Descendants().Single(e => (string?)e.Attribute("id") == "paragraph").SetAttributeValue("class", "emphasis");
        after.SetContentXml("one", xml);
        Assert.Equal(EpubEditionChangeKind.Xml, Assert.Single(before.CompareTo(after).Resources).Kind);
    }
    [Fact]
    public void RenamesAndBinaryResourceChangesRetainManifestIdentity() {
        var before = Book(); before.AddResource("asset", "EPUB/data.png", "image/png", new byte[] { 1 });
        var after = EpubPublication.Load(new MemoryStream(before.Write().Bytes));
        after.RenameResource("asset", "EPUB/new.png"); after.UpdateResource("asset", new byte[] { 2 });
        var change = Assert.Single(before.CompareTo(after).Resources);
        Assert.Equal("asset", change.ManifestId);
        Assert.Equal("EPUB/data.png", change.PreviousPath); Assert.Equal("EPUB/new.png", change.CurrentPath);
        Assert.Equal(EpubEditionChangeKind.Location | EpubEditionChangeKind.Declaration | EpubEditionChangeKind.Binary, change.Kind);
    }
    [Fact]
    public void AddedAndRemovedResourcesAndCancellationAreExplicit() {
        var before = Book(); before.AddResource("old", "EPUB/old.bin", "application/octet-stream", new byte[] { 1 });
        var after = EpubPublication.Load(new MemoryStream(before.Write().Bytes));
        after.RemoveResource("old"); after.AddResource("new", "EPUB/new.bin", "application/octet-stream", new byte[] { 2 });
        Assert.Equal(new[] { EpubEditionChangeKind.Added, EpubEditionChangeKind.Removed }, before.CompareTo(after).Resources.Select(item => item.Kind));
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => before.CompareTo(after, cancellation.Token));
    }
    [Fact]
    public void XmlSerializationIsDistinguishedFromContentChanges() {
        var before = Book(); var after = EpubPublication.Load(new MemoryStream(before.Write().Bytes));
        string xml = System.Text.Encoding.UTF8.GetString(after.GetResourceBytes("one"));
        after.UpdateResource("one", System.Text.Encoding.UTF8.GetBytes(xml.Replace("\"", "'")));
        Assert.Equal(EpubEditionChangeKind.Serialization, Assert.Single(before.CompareTo(after).Resources).Kind);
    }
    [Fact]
    public void UnmanifestedExtensionChangesAreNotOverlooked() {
        byte[] source = Book().Write().Bytes;
        EpubPublication LoadWithExtension(byte value) {
            using var stream = new MemoryStream(); stream.Write(source, 0, source.Length); stream.Position = 0;
            using (var archive = new System.IO.Compression.ZipArchive(stream, System.IO.Compression.ZipArchiveMode.Update, true)) {
                using var entry = archive.CreateEntry("vendor/state.bin").Open(); entry.WriteByte(value);
            }
            return EpubPublication.Load(new MemoryStream(stream.ToArray()));
        }
        var change = Assert.Single(LoadWithExtension(1).CompareTo(LoadWithExtension(2)).Resources);
        Assert.Null(change.ManifestId); Assert.Equal("vendor/state.bin", change.PreviousPath); Assert.Equal(EpubEditionChangeKind.Binary, change.Kind);
    }

    [Fact]
    public void LongTextDifferencesAreDetectedBeyondTheExcerptBoundary() {
        var before = Book(); var xml = before.GetContentXml("one");
        xml.Descendants().Single(e => (string?)e.Attribute("id") == "paragraph").Value = new string('a', 5000) + "old";
        before.SetContentXml("one", xml);
        var after = EpubPublication.Load(new MemoryStream(before.Write().Bytes));
        xml.Descendants().Single(e => (string?)e.Attribute("id") == "paragraph").Value = new string('a', 5000) + "new";
        after.SetContentXml("one", xml);
        var change = before.CompareTo(after).TextChanges.Single(item => item.Locator == "#paragraph");
        Assert.True(change.IsTruncated); Assert.Equal(4096, change.PreviousText!.Length);
        Assert.Equal(change.PreviousText, change.CurrentText);
    }

    [Fact]
    public void ChangedTextBlockLimitRejectsInsteadOfReturningAPartialReport() {
        var before = EpubPublication.Create("Many paragraphs", "en");
        before.AddChapter("one", "EPUB/one.xhtml", "One", string.Concat(Enumerable.Range(0, 10_000).Select(i => "<p id='p" + i + "'>Before</p>")));
        var after = EpubPublication.Load(new MemoryStream(before.Write().Bytes));
        var xml = after.GetContentXml("one");
        foreach (var paragraph in xml.Descendants(XName.Get("p", "http://www.w3.org/1999/xhtml"))) paragraph.Value = "After";
        after.SetContentXml("one", xml);
        Assert.Throws<InvalidDataException>(() => before.CompareTo(after));
        Assert.Contains("Before", before.GetContentXml("one").Root!.Value);
        Assert.Contains("After", after.GetContentXml("one").Root!.Value);
    }

    private static EpubPublication Book(EpubVersion version = EpubVersion.Epub3) {
        var book = EpubPublication.Create("Book", "en", version: version);
        book.AddChapter("one", "EPUB/one.xhtml", "One", "<p id='paragraph'>Original</p>");
        book.AddChapter("two", "EPUB/two.xhtml", "Two", "<p>Second</p>");
        return book;
    }
}
