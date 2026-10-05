using System.IO.Compression;
using System.Text.Json.Nodes;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookProjectRevisionTests {
    [Fact]
    public void NamedRevisionsSurviveReopenAndRestoreIsUndoable() {
        var project = BookProject.Create("First edition");
        var first = project.CreateRevision("Editorial baseline");
        project.SetMetadata("Second edition", "en", "Author");
        var second = project.CreateRevision("Copyedited");
        var reopened = BookProject.LoadProject(project.ToProjectBytes());
        Assert.Equal(new[] { first, second }, reopened.Revisions);
        reopened.RestoreRevision(first.Id);
        Assert.Equal("First edition", reopened.Publication.Title);
        reopened.Undo(); Assert.Equal("Second edition", reopened.Publication.Title);
        reopened.Redo(); Assert.Equal("First edition", reopened.Publication.Title);
        reopened.RemoveRevision(first.Id);
        var saved = BookProject.LoadProject(reopened.ToProjectBytes());
        Assert.Equal(second, Assert.Single(saved.Revisions));
        Assert.Equal("First edition", saved.Publication.Title);
    }

    [Fact]
    public void CancelledCaptureAndRestoreLeavePublicationAndHistoryUnchanged() {
        var project = BookProject.Create("First");
        var revision = project.CreateRevision("First");
        project.SetMetadata("Second", "en", "");
        byte[] before = project.ToProjectBytes();
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => project.CreateRevision("Cancelled", cancellation.Token));
        Assert.Throws<OperationCanceledException>(() => project.RestoreRevision(revision.Id, cancellation.Token));
        Assert.Equal(before, project.ToProjectBytes());
        project.Undo(); Assert.Equal("First", project.Publication.Title);
    }

    [Fact]
    public void RevisionCountNeverEvictsExistingSnapshots() {
        var project = BookProject.Create("Book");
        for (int i = 0; i < BookProject.MaximumRevisionCount; i++) project.CreateRevision("Revision " + i);
        Assert.Throws<InvalidOperationException>(() => project.CreateRevision("Overflow"));
        Assert.Equal(BookProject.MaximumRevisionCount, project.Revisions.Count);
        var reopened = BookProject.LoadProject(project.ToProjectBytes());
        reopened.RemoveRevision(reopened.Revisions[0].Id);
        reopened.CreateRevision("Replacement");
        Assert.Equal(BookProject.MaximumRevisionCount, reopened.Revisions.Count);
    }

    [Fact]
    public void RevisionByteExhaustionRetainsEarlierSnapshots() {
        var project = BookProject.Create("Bounded history");
        byte[] payload = new byte[16 * 1024 * 1024]; new Random(42).NextBytes(payload);
        project.Publication.AddResource("attachment", "EPUB/attachment.bin", "application/octet-stream", payload);
        var first = project.CreateRevision("Baseline");
        int capacity = (int)(BookProject.MaximumRevisionBytes / first.PublicationBytes);
        for (int i = 1; i < capacity; i++) project.CreateRevision("Revision " + i);
        Assert.Throws<InvalidOperationException>(() => project.CreateRevision("Overflow"));
        Assert.Equal(capacity, project.Revisions.Count);
        Assert.Equal(first, project.Revisions[0]);
        project.RemoveRevision(first.Id);
        project.CreateRevision("Replacement");
        Assert.Equal(capacity, project.Revisions.Count);
    }

    [Theory]
    [InlineData("hash")]
    [InlineData("missing")]
    [InlineData("extra")]
    [InlineData("duplicate")]
    [InlineData("version")]
    [InlineData("length")]
    public void CorruptOrUndeclaredHistoryIsRejected(string kind) {
        var project = BookProject.Create("Book"); project.CreateRevision("Original");
        byte[] changed = Rewrite(project.ToProjectBytes(), (entries, record) => {
            var revisions = record["Revisions"]!.AsArray();
            if (kind == "hash") revisions[0]!["Sha256"] = new string('0', 64);
            if (kind == "length") revisions[0]!["PublicationBytes"] = 1;
            if (kind == "duplicate") revisions.Add(revisions[0]!.DeepClone());
            if (kind == "version") record["Version"] = 1;
            if (kind == "missing") entries.Remove(entries.Keys.Single(name => name.StartsWith("revisions/")));
            if (kind == "extra") entries.Add("unknown.bin", new byte[] { 1 });
        });
        Assert.Throws<InvalidDataException>(() => BookProject.LoadProject(changed));
    }

    [Fact]
    public void VersionOneProjectsRemainReadable() {
        var project = BookProject.Create("Legacy");
        byte[] legacy = Rewrite(project.ToProjectBytes(), (_, record) => { record["Version"] = 1; record.Remove("Revisions"); });
        var reopened = BookProject.LoadProject(legacy);
        Assert.Equal("Legacy", reopened.Publication.Title); Assert.Empty(reopened.Revisions);
    }

    private static byte[] Rewrite(byte[] bytes, Action<Dictionary<string, byte[]>, JsonObject> edit) {
        using var input = new ZipArchive(new MemoryStream(bytes));
        var entries = input.Entries.ToDictionary(entry => entry.FullName, entry => { using var data = new MemoryStream(); using var stream = entry.Open(); stream.CopyTo(data); return data.ToArray(); });
        var record = JsonNode.Parse(entries["project.json"])!.AsObject();
        edit(entries, record); entries["project.json"] = System.Text.Encoding.UTF8.GetBytes(record.ToJsonString());
        using var output = new MemoryStream();
        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, true))
            foreach (var entry in entries) { using var stream = zip.CreateEntry(entry.Key).Open(); stream.Write(entry.Value); }
        return output.ToArray();
    }
}
