using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>
/// Text box edits made without the SDK keep run formatting valid, write auto-fit in schema order, and survive a
/// read-only snapshot.
/// </summary>
public class PowerPointTextBoxWithoutSdkTests {
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    [Fact]
    public void ReplacingTextKeepsAPrefixTheRunFormattingInherited() {
        byte[] source = Rewrite(CreateDeck(), slide => slide.Replace("<a:r>",
            "<a:r xmlns:x=\"urn:officeimo:test\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">")
            .Replace("<a:t>Text", "<a:rPr lang=\"en-US\" mc:Ignorable=\"x\" x:flag=\"1\"/><a:t>Text"));
        byte[] saved;
        using (var presentation = PowerPointDocument.Load(new MemoryStream(source))) {
            presentation.Slides[0].TextBoxes.Single().Text = "Replaced";
            Assert.False(presentation.IsSdkLoaded);
            saved = presentation.ToBytes();
        }

        XElement slide = XElement.Parse(SlideXml(saved));
        XElement properties = slide.Descendants(A + "rPr").Single(element => element.Attribute(Mc + "Ignorable") != null);
        Assert.Equal("urn:officeimo:test", properties.GetNamespaceOfPrefix("x")?.NamespaceName);
    }

    [Fact]
    public void AutoFitGoesBeforeTheBodyExtensions() {
        byte[] source = Rewrite(CreateDeck(), slide => slide.Replace("<a:bodyPr/>",
            "<a:bodyPr><a:scene3d><a:camera prst=\"orthographicFront\"/><a:lightRig rig=\"threePt\" dir=\"t\"/></a:scene3d></a:bodyPr>"));
        byte[] saved;
        using (var presentation = PowerPointDocument.Load(new MemoryStream(source))) {
            PowerPointTextBox box = presentation.Slides[0].TextBoxes.Single();
            box.TextAutoFit = PowerPointTextAutoFit.Shape;
            saved = presentation.ToBytes();
        }

        XElement body = XElement.Parse(SlideXml(saved)).Descendants(A + "bodyPr").Single();
        Assert.Equal(new[] { "spAutoFit", "scene3d" }, body.Elements().Select(element => element.Name.LocalName));
    }

    [Fact]
    public void AReadOnlySnapshotKeepsEditsMadeWithoutTheSdk() {
        using var presentation = PowerPointDocument.Load(new MemoryStream(CreateDeck()), new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
        presentation.Slides[0].TextBoxes.Single().Name = "Renamed";

        using var snapshot = PowerPointDocument.Load(new MemoryStream(presentation.ToBytes()));
        Assert.Equal("Renamed", snapshot.Slides[0].TextBoxes.Single().Name);
    }

    [Fact]
    public void TheRunCallbackOfAddParagraphFormatsARunInTheParagraph() {
        using var presentation = PowerPointDocument.Load(new MemoryStream(CreateDeck()));
        PowerPointTextBox box = presentation.Slides[0].TextBoxes.Single();

        PowerPointParagraph paragraph = box.AddParagraph("First", configure: added => added.Text = "Second", run: run => run.Bold = true);
        PowerPointParagraph opened = box.AddParagraph("Third", configure: added => {
            added.Text = "Fourth";
            _ = presentation.OpenXmlDocument;
        }, run: run => run.Bold = true);

        Assert.Contains(paragraph.Runs, run => run.Bold);
        Assert.Contains(opened.Runs, run => run.Bold);
    }

    private static byte[] CreateDeck() {
        using var presentation = PowerPointDocument.Create();
        presentation.AddSlide().AddTextBox("Text");
        return presentation.ToBytes();
    }

    private static string SlideXml(byte[] package) {
        using var archive = new ZipArchive(new MemoryStream(package));
        using var reader = new StreamReader(archive.GetEntry("ppt/slides/slide1.xml")!.Open(), Encoding.UTF8);
        return reader.ReadToEnd();
    }

    private static byte[] Rewrite(byte[] package, Func<string, string> change) {
        var output = new MemoryStream();
        output.Write(package, 0, package.Length);
        using (var archive = new ZipArchive(output, ZipArchiveMode.Update, leaveOpen: true)) {
            ZipArchiveEntry entry = archive.GetEntry("ppt/slides/slide1.xml")!;
            string text;
            using (var reader = new StreamReader(entry.Open(), Encoding.UTF8)) text = reader.ReadToEnd();
            string changed = change(text);
            Assert.NotEqual(text, changed);
            entry.Delete();
            using var writer = new StreamWriter(archive.CreateEntry("ppt/slides/slide1.xml").Open(), new UTF8Encoding(false));
            writer.Write(changed);
        }
        return output.ToArray();
    }
}
