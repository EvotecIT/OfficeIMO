using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>
/// Speaker notes are created, rewritten and read without the SDK, writing the same notes slides as through it.
/// </summary>
public class PowerPointNotesWithoutSdkTests {
    [Fact]
    public void NotesEditsWriteTheSameNotesWithoutTheSdk() {
        byte[] source = CreateDeck();

        byte[] throughSdk;
        using (var presentation = PowerPointDocument.Load(new MemoryStream(source))) {
            _ = presentation.OpenXmlDocument;
            Edit(presentation);
            throughSdk = presentation.ToBytes();
        }
        byte[] withoutSdk;
        using (var presentation = PowerPointDocument.Load(new MemoryStream(source))) {
            Edit(presentation);
            Assert.False(presentation.IsSdkLoaded, string.Join("; ", presentation.Diagnostics.Select(diagnostic => diagnostic.Message)));
            withoutSdk = presentation.ToBytes();
        }

        Assert.Equal(PowerPointTablesWithoutSdkTests.Parts(throughSdk, "ppt/notesSlides/notesSlide"), PowerPointTablesWithoutSdkTests.Parts(withoutSdk, "ppt/notesSlides/notesSlide"));
        Assert.Equal(SlideRelationships(throughSdk), SlideRelationships(withoutSdk));
    }

    [Fact]
    public void ReadingNotesMatchesTheSdk() {
        byte[] edited;
        using (var presentation = PowerPointDocument.Load(new MemoryStream(CreateDeck()))) {
            Edit(presentation);
            edited = presentation.ToBytes();
        }

        using var loaded = PowerPointDocument.Load(new MemoryStream(edited));
        string[] read = loaded.Slides.Select(Describe).ToArray();
        Assert.False(loaded.IsSdkLoaded);
        _ = loaded.OpenXmlDocument;
        Assert.Equal(loaded.Slides.Select(Describe).ToArray(), read);
    }

    [Fact]
    public void NotesOnAPresentationWithoutANotesMasterAddTheSameMasterWithoutTheSdk() {
        byte[] source = WithoutNotesMaster(CreateDeck());

        byte[] throughSdk;
        using (var presentation = PowerPointDocument.Load(new MemoryStream(source))) {
            _ = presentation.OpenXmlDocument;
            presentation.Slides[0].Notes.Text = "Speaker notes";
            throughSdk = presentation.ToBytes();
        }
        byte[] withoutSdk;
        using (var presentation = PowerPointDocument.Load(new MemoryStream(source))) {
            presentation.Slides[0].Notes.Text = "Speaker notes";
            Assert.False(presentation.IsSdkLoaded, string.Join("; ", presentation.Diagnostics.Select(diagnostic => diagnostic.Message)));
            withoutSdk = presentation.ToBytes();
        }

        Assert.Equal(PowerPointTablesWithoutSdkTests.Parts(throughSdk, "ppt/notesMasters/notesMaster"), PowerPointTablesWithoutSdkTests.Parts(withoutSdk, "ppt/notesMasters/notesMaster"));
        using var reopened = PowerPointDocument.Load(new MemoryStream(withoutSdk));
        var presentationPart = reopened.OpenXmlDocument.PresentationPart!;
        var master = presentationPart.NotesMasterPart!;
        Assert.NotNull(master.ThemePart?.Theme);
        Assert.Equal(presentationPart.GetIdOfPart(master), Assert.Single(presentationPart.Presentation!.NotesMasterIdList!.ChildElements).GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships").Value);
        Assert.Same(master, presentationPart.SlideParts.Single(slide => slide.NotesSlidePart != null).NotesSlidePart!.NotesMasterPart);
        Assert.Equal("Speaker notes", reopened.Slides[0].Notes.Text);
    }

    [Fact]
    public void NotesTextKeptFromARemovedSlideStaysItsOwn() {
        byte[] source;
        using (var created = PowerPointDocument.Create()) {
            created.AddSlide().Notes.Text = "Old notes";
            created.AddSlide();
            source = created.ToBytes();
        }

        byte[] saved;
        using (var presentation = PowerPointDocument.Load(new MemoryStream(source))) {
            PowerPointParagraph paragraph = presentation.Slides[0].Notes.Paragraphs[0];
            PowerPointTextRun run = paragraph.Runs[0];
            presentation.RemoveSlide(0);
            // The remaining slide's new notes can take the removed notes' part name.
            presentation.Slides[0].Notes.Text = "New notes";
            run.Text = "Edited";
            Assert.Throws<InvalidOperationException>(() => run.SetHyperlink("https://example.com"));
            Assert.False(presentation.IsSdkLoaded);

            _ = presentation.OpenXmlDocument;
            Assert.Equal("Edited", run.Text);
            Assert.Equal("Edited", paragraph.Text);
            saved = presentation.ToBytes();
        }

        using var reopened = PowerPointDocument.Load(new MemoryStream(saved));
        Assert.Equal("New notes", reopened.Slides[0].Notes.Text);
        var notesPart = reopened.OpenXmlDocument.PresentationPart!.SlideParts.Single().NotesSlidePart!;
        Assert.Empty(notesPart.HyperlinkRelationships);
    }

    [Fact]
    public void NotesTooLargeForThePackageReadTheSameEverywhere() {
        string text = new('x', 17 * 1024 * 1024);
        byte[] source;
        using (var created = PowerPointDocument.Create()) {
            created.AddSlide().Notes.Text = text;
            source = created.ToBytes();
        }

        using var presentation = PowerPointDocument.Load(new MemoryStream(source));
        PowerPointSlide slide = presentation.Slides[0];
        Assert.True(slide.HasSpeakerNotes);
        Assert.Equal(text.Length, slide.GetSpeakerNotesText().Length);
        Assert.Single(slide.Notes.Paragraphs);
        Assert.True(slide.Notes.TryGetText(out string read));
        Assert.Equal(text.Length, read.Length);
    }

    private static void Edit(PowerPointDocument presentation) {
        PowerPointSlide[] slides = presentation.Slides.ToArray();
        slides[0].Notes.Text = "First point\nSecond point";
        IReadOnlyList<PowerPointParagraph> paragraphs = slides[1].Notes.SetParagraphs(new[] { "Opening", "Closing" });
        paragraphs[1].Runs[0].Bold = true;
        slides[1].Notes.Text = slides[1].Notes.Text + "\nAdded";
        // The third slide keeps no notes.
    }

    private static string Describe(PowerPointSlide slide) {
        bool hasText = slide.Notes.TryGetText(out string text);
        bool hasExisting = slide.Notes.TryGetExistingText(out string existing);
        return string.Join("|", slide.HasSpeakerNotes, slide.GetSpeakerNotesText(), hasText, text, hasExisting, existing,
            string.Join("/", slide.Notes.Paragraphs.Select(paragraph => paragraph.Text + ":" + paragraph.Runs.Any(run => run.Bold))));
    }

    private static byte[] CreateDeck() {
        using var presentation = PowerPointDocument.Create();
        for (int index = 0; index < 3; index++) presentation.AddSlide().AddTextBox("Slide " + index);
        return presentation.ToBytes();
    }

    // Each slide's relationships as id, type and the part they resolve to; the SDK writes absolute targets where the
    // package writes relative ones.
    private static string[] SlideRelationships(byte[] package) {
        using var archive = new ZipArchive(new MemoryStream(package));
        var result = new List<string>();
        foreach (ZipArchiveEntry entry in archive.Entries.Where(entry => entry.FullName.StartsWith("ppt/slides/_rels/", StringComparison.Ordinal)).OrderBy(entry => entry.FullName, StringComparer.Ordinal)) {
            XElement root;
            using (Stream stream = entry.Open()) root = XElement.Load(stream);
            foreach (XElement relationship in root.Elements().OrderBy(element => (string?)element.Attribute("Id"), StringComparer.Ordinal)) {
                string target = (string)relationship.Attribute("Target")!;
                string resolved = target.StartsWith("/", StringComparison.Ordinal) ? target : new Uri(new Uri("http://package/ppt/slides/"), target).AbsolutePath;
                result.Add(entry.FullName + "|" + (string?)relationship.Attribute("Id") + "|" + (string?)relationship.Attribute("Type") + "|" + resolved);
            }
        }
        return result.ToArray();
    }

    // The deck without its notes master: the parts, their relationships, content types and the presentation's list.
    private static byte[] WithoutNotesMaster(byte[] package) {
        var output = new MemoryStream();
        output.Write(package, 0, package.Length);
        using (var archive = new ZipArchive(output, ZipArchiveMode.Update, leaveOpen: true)) {
            string[] removed = archive.Entries.Where(entry => entry.FullName.StartsWith("ppt/notesMasters/", StringComparison.Ordinal)).Select(entry => entry.FullName).ToArray();
            foreach (string name in removed) archive.GetEntry(name)!.Delete();
            Rewrite(archive, "ppt/_rels/presentation.xml.rels", root => root.Elements().Where(element => ((string?)element.Attribute("Type"))?.EndsWith("/notesMaster", StringComparison.Ordinal) == true).Remove());
            Rewrite(archive, "ppt/presentation.xml", root => root.Elements().Where(element => element.Name.LocalName == "notesMasterIdLst").Remove());
            Rewrite(archive, "[Content_Types].xml", root => root.Elements().Where(element => ((string?)element.Attribute("PartName"))?.StartsWith("/ppt/notesMasters/", StringComparison.Ordinal) == true).Remove());
        }
        return output.ToArray();
    }

    private static void Rewrite(ZipArchive archive, string name, Action<XElement> change) {
        ZipArchiveEntry entry = archive.GetEntry(name)!;
        XElement root;
        using (Stream stream = entry.Open()) root = XElement.Load(stream);
        change(root);
        entry.Delete();
        using var writer = new StreamWriter(archive.CreateEntry(name).Open(), new UTF8Encoding(false));
        writer.Write(root.ToString(SaveOptions.DisableFormatting));
    }
}
