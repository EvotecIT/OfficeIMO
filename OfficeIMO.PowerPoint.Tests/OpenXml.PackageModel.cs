using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests;

/// <summary>
/// The presentation model built on the Open XML layer must write packages the Open XML SDK reads and validates, keep
/// every part an edit does not touch byte for byte, and keep the rest of an edited part as it was written.
/// </summary>
public class OpenXmlPackageModelTests {
    [Fact]
    public void CreatesAPresentationTheSdkValidatesAndReads() {
        byte[] package = CreateDeck(5);

        using var stream = new MemoryStream(package);
        using PresentationDocument document = PresentationDocument.Open(stream, false);
        Assert.Empty(new OpenXmlValidator().Validate(document));
        List<P.SlideId> slides = document.PresentationPart!.Presentation!.SlideIdList!.Elements<P.SlideId>().ToList();
        Assert.Equal(5, slides.Count);
        var third = (SlidePart)document.PresentationPart.GetPartById(slides[2].RelationshipId!.Value!);
        Assert.Equal(new[] { "Slide 3", "First point of slide 3", "Second point", "Third & last <point>" },
            third.Slide!.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>().Select(paragraph => paragraph.InnerText).ToArray());
        Assert.Equal("Title and Content", third.SlideLayoutPart!.SlideLayout!.CommonSlideData!.Name!.Value);
    }

    [Fact]
    public void EditingOneSlideKeepsEveryOtherPartAndTheRestOfTheSlide() {
        byte[] original = CreateDeck(20);

        byte[] edited;
        using (var source = new MemoryStream(original))
        using (PowerPointPackageModel model = PowerPointPackageModel.Open(source)) {
            PowerPointModelSlide slide = model.GetSlide(9);
            PowerPointModelSlide.SetText(slide.FindTitle()!.Value, "Edited title");
            slide.AddTextBox(457200, 4572000, 6096000, 914400, "Added later");
            var output = new MemoryStream();
            model.Save(output);
            edited = output.ToArray();
        }

        using (var stream = new MemoryStream(edited))
        using (PresentationDocument document = PresentationDocument.Open(stream, false)) {
            Assert.Empty(new OpenXmlValidator().Validate(document));
            P.SlideId tenth = document.PresentationPart!.Presentation!.SlideIdList!.Elements<P.SlideId>().ElementAt(9);
            var slide = (SlidePart)document.PresentationPart.GetPartById(tenth.RelationshipId!.Value!);
            Assert.Equal(new[] { "Edited title", "First point of slide 10", "Second point", "Third & last <point>", "Added later" },
                slide.Slide!.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>().Select(paragraph => paragraph.InnerText).ToArray());
        }
        Assert.Equal(new[] { "ppt/slides/slide10.xml" }, ChangedMembers(original, edited));
    }

    public static IEnumerable<object[]> CorpusPresentations() => OpenXmlSlideTextReaderTests.CorpusPresentations();

    [Theory]
    [MemberData(nameof(CorpusPresentations))]
    public void OpenAndSaveKeepsCorpusPresentationsByteForByte(string relativePath) {
        byte[] original = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, relativePath));

        byte[] saved;
        using (var source = new MemoryStream(original))
        using (OfficeOpcEditablePackage package = OfficeOpcEditablePackage.Open(source, leaveOpen: true)) {
            // Parsing every XML part without changing it must not change what is written.
            foreach (string part in package.PartNames.ToList()) {
                if (part.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) || part.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)) package.GetXml(part);
            }
            var output = new MemoryStream();
            package.Save(output);
            saved = output.ToArray();
        }

        Assert.Empty(ChangedMembers(original, saved));
    }

    [Theory]
    [MemberData(nameof(CorpusPresentations))]
    public void AddingATextBoxToCorpusPresentationsKeepsTheirOtherContent(string relativePath) {
        byte[] original = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, relativePath));
        int originalErrors = CountValidationErrors(original);

        byte[] edited;
        string slidePart;
        using (var source = new MemoryStream(original)) {
            PowerPointPackageModel model;
            try {
                model = PowerPointPackageModel.Open(source);
            } catch (InvalidDataException) {
                return;
            }
            using (model) {
                if (model.SlideCount == 0) return;
                PowerPointModelSlide slide = model.GetSlide(0);
                slide.AddTextBox(0, 0, 914400, 914400, "Added & kept");
                slidePart = model.GetSlidePartName(0);
                var output = new MemoryStream();
                model.Save(output);
                edited = output.ToArray();
            }
        }

        Assert.Equal(new[] { slidePart.Substring(1) }, ChangedMembers(original, edited), StringComparer.OrdinalIgnoreCase);
        Assert.True(CountValidationErrors(edited) <= originalErrors);
        XElement before = LoadPart(original, slidePart);
        XElement after = LoadPart(edited, slidePart);
        XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
        XElement added = after.Descendants(p + "sp").Last();
        Assert.Contains("Added & kept", added.Value, StringComparison.Ordinal);
        added.Remove();
        Assert.True(XNode.DeepEquals(Normalize(before), Normalize(after)), "The edited slide changed outside the added shape.");
    }

    internal static byte[] CreateDeck(int slides) {
        using PowerPointPackageModel model = PowerPointPackageModel.Create();
        for (int index = 1; index <= slides; index++) {
            PowerPointModelSlide slide = model.AddSlide();
            slide.AddTitle("Slide " + index);
            slide.AddTextBox(838200, 1825625, 10515600, 4351338, "First point of slide " + index, "Second point", "Third & last <point>");
        }
        var output = new MemoryStream();
        model.Save(output);
        return output.ToArray();
    }

    private static int CountValidationErrors(byte[] package) {
        using var stream = new MemoryStream(package);
        using PresentationDocument document = PresentationDocument.Open(stream, false);
        return new OpenXmlValidator().Validate(document).Count();
    }

    // Members whose stored bytes differ, or that exist in only one package.
    private static List<string> ChangedMembers(byte[] first, byte[] second) {
        var changed = new List<string>();
        using OfficeZipPackageReader left = OfficeZipPackageReader.Open(new MemoryStream(first));
        using OfficeZipPackageReader right = OfficeZipPackageReader.Open(new MemoryStream(second));
        Assert.Equal(left.Count, right.Count);
        for (int index = 0; index < left.Count; index++) {
            string name = left.GetName(index);
            Assert.Equal(name, right.GetName(index));
            Assert.True(left.TryReadStored(name, out OfficeZipStoredMember a));
            Assert.True(right.TryReadStored(name, out OfficeZipStoredMember b));
            using (a)
            using (b) {
                bool same = a.Crc == b.Crc && a.Method == b.Method
                    && a.Compressed.Array.AsSpan(0, a.Compressed.Length).SequenceEqual(b.Compressed.Array.AsSpan(0, b.Compressed.Length));
                if (!same) changed.Add(name);
            }
        }
        return changed;
    }

    private static XElement LoadPart(byte[] package, string partName) {
        using OfficeZipPackageReader zip = OfficeZipPackageReader.Open(new MemoryStream(package));
        Assert.True(zip.TryRead(partName.Substring(1), out OfficePooledBytes content));
        using (content) return XElement.Load(new MemoryStream(content.Array, 0, content.Length), LoadOptions.PreserveWhitespace);
    }

    // Declarations may be written differently; elements, attributes and text must match.
    private static XElement Normalize(XElement element) {
        var copy = new XElement(element);
        foreach (XElement descendant in copy.DescendantsAndSelf()) descendant.Attributes().Where(attribute => attribute.IsNamespaceDeclaration).Remove();
        return copy;
    }
}
