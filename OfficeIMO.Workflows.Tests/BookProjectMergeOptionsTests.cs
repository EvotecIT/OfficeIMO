using OfficeIMO.Epub;
using System.Xml.Linq;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookProjectMergeOptionsTests {
    private static readonly XNamespace Html = "http://www.w3.org/1999/xhtml";

    [Fact]
    public void ScopedMergeOptionsParticipateInUndoRedoAndProjectReopen() {
        var project = Project();byte[] before = project.Export().Bytes;
        project.MergeChapters("one", "two", "start", Options());
        var merged = project.Export().Bytes;
        Assert.Single(project.Publication.Spine);
        Assert.Equal(new[] { "body", "second-body" }, project.Publication.GetContentXml("one").Root!.Element(Html + "body")!
            .Elements(Html + "div").Select(e => (string?)e.Attribute("id")));
        Assert.Contains("#second-body", project.Publication.GetContentXml("one").ToString());
        project.Undo();Assert.Equal(before, project.Export().Bytes);
        project.Redo();Assert.Equal(merged, project.Export().Bytes);
        var reopened = BookProject.LoadProject(project.ToProjectBytes());
        Assert.Equal(merged, reopened.Export().Bytes);
        Assert.Equal("start", reopened.Publication.Read().TableOfContents[1].Fragment);
    }

    [Fact]
    public void FailedReconciliationDoesNotConsumeTheRedoTransaction() {
        var project = Project();project.MergeChapters("one", "two", "start", Options());
        byte[] merged = project.Export().Bytes;project.Undo();byte[] before = project.ToProjectBytes();
        Assert.Throws<InvalidDataException>(() => project.MergeChapters("one", "two", "start", new EpubChapterMergeOptions { PreserveBodyScopes = true }));
        Assert.Equal(before, project.ToProjectBytes());
        project.Redo();Assert.Equal(merged, project.Export().Bytes);
    }

    private static EpubChapterMergeOptions Options() => new() {
        PreserveBodyScopes = true, SecondChapterIdMap = new Dictionary<string, string> { ["body"] = "second-body" }
    };

    private static BookProject Project() {
        var publication = EpubPublication.Create("Book", "en");
        publication.AddChapter("one", "EPUB/one.xhtml", "One", "<h1>One</h1><p><a href='two.xhtml#body'>Next</a></p>");
        publication.AddChapter("two", "EPUB/two.xhtml", "Two", "<h1>Two</h1><p>Second</p>");
        foreach (string id in new[] { "one", "two" }) {
            var xml = publication.GetContentXml(id);xml.Root!.Element(Html + "body")!.SetAttributeValue("id", "body");
            publication.SetContentXml(id, xml);
        }
        return BookProject.FromEpub(publication.Write().Bytes);
    }
}
