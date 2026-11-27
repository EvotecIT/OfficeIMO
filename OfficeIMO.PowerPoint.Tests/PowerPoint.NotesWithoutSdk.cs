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

        Assert.Equal(PowerPointTablesWithoutSdkTests.Parts(throughSdk, "ppt/notesSlides/"), PowerPointTablesWithoutSdkTests.Parts(withoutSdk, "ppt/notesSlides/"));
        Assert.Equal(PowerPointTablesWithoutSdkTests.Parts(throughSdk, "ppt/slides/_rels/"), PowerPointTablesWithoutSdkTests.Parts(withoutSdk, "ppt/slides/_rels/"));
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
