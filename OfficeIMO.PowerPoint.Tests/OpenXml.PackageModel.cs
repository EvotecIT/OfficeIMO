using System.Text;
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

    [Fact]
    public void NewShapesTakeAnIdAboveEveryDrawingObjectOnTheSlide() {
        byte[] deck = Rewrite(CreateDeck(1), "ppt/slides/slide1.xml", xml => xml.Replace("</p:spTree>",
            "<p:grpSp><p:nvGrpSpPr><p:cNvPr id=\"20\" name=\"Group\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>"
            + "<p:pic><p:nvPicPr><p:cNvPr id=\"40\" name=\"Picture\"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill/><p:spPr/></p:pic>"
            + "</p:grpSp><p:cxnSp><p:nvCxnSpPr><p:cNvPr id=\"30\" name=\"Connector\"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr/></p:cxnSp></p:spTree>"));

        byte[] edited;
        using (var source = new MemoryStream(deck))
        using (PowerPointPackageModel model = PowerPointPackageModel.Open(source)) {
            model.GetSlide(0).AddTextBox(0, 0, 914400, 914400, "Added");
            var output = new MemoryStream();
            model.Save(output);
            edited = output.ToArray();
        }

        XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
        XElement added = LoadPart(edited, "/ppt/slides/slide1.xml").Element(p + "cSld")!.Element(p + "spTree")!.Elements(p + "sp").Last();
        Assert.Equal("41", added.Element(p + "nvSpPr")!.Element(p + "cNvPr")!.Attribute("id")!.Value);
    }

    [Fact]
    public void OpensAPresentationWithoutASlideListAndAddsSlidesToIt() {
        byte[] deck = Rewrite(CreateDeck(0), "ppt/presentation.xml",
            xml => System.Text.RegularExpressions.Regex.Replace(xml, "<p:sldIdLst\\s*/>|<p:sldIdLst>.*?</p:sldIdLst>", string.Empty));
        Assert.DoesNotContain("sldIdLst", Encoding.UTF8.GetString(LoadPartBytes(deck, "ppt/presentation.xml")));

        byte[] edited;
        using (var source = new MemoryStream(deck))
        using (PowerPointPackageModel model = PowerPointPackageModel.Open(source)) {
            Assert.Equal(0, model.SlideCount);
            model.AddSlide().AddTitle("First");
            var output = new MemoryStream();
            model.Save(output);
            edited = output.ToArray();
        }

        Assert.Equal(0, CountValidationErrors(edited));
    }

    [Fact]
    public void FindsPartsWhoseNamesHoldAnEncodedPercentSign() {
        byte[] deck = CreateDeck(1);
        deck = Rename(deck, "ppt/slides/slide1.xml", "ppt/slides/slide%25201.xml");
        deck = Rename(deck, "ppt/slides/_rels/slide1.xml.rels", "ppt/slides/_rels/slide%25201.xml.rels");
        deck = Rewrite(deck, "ppt/_rels/presentation.xml.rels", xml => xml.Replace("slides/slide1.xml", "slides/slide%25201.xml"));
        deck = Rewrite(deck, "[Content_Types].xml", xml => xml.Replace("/ppt/slides/slide1.xml", "/ppt/slides/slide%25201.xml"));

        using var source = new MemoryStream(deck);
        using PowerPointPackageModel model = PowerPointPackageModel.Open(source);
        Assert.NotNull(model.GetSlide(0).FindTitle());
    }

    private static byte[] Rewrite(byte[] package, string member, Func<string, string> change) {
        string xml = Encoding.UTF8.GetString(LoadPartBytes(package, member));
        return Replace(package, member, member, Encoding.UTF8.GetBytes(change(xml)));
    }

    private static byte[] Rename(byte[] package, string member, string newName) =>
        Replace(package, member, newName, LoadPartBytes(package, member));

    private static byte[] Replace(byte[] package, string member, string newName, byte[] content) {
        var stream = new MemoryStream();
        stream.Write(package, 0, package.Length);
        using (var archive = new System.IO.Compression.ZipArchive(stream, System.IO.Compression.ZipArchiveMode.Update, leaveOpen: true)) {
            archive.GetEntry(member)!.Delete();
            using Stream entry = archive.CreateEntry(newName).Open();
            entry.Write(content, 0, content.Length);
        }
        return stream.ToArray();
    }

    private static byte[] LoadPartBytes(byte[] package, string member) {
        using OfficeZipPackageReader zip = OfficeZipPackageReader.Open(new MemoryStream(package));
        Assert.True(zip.TryRead(member, out OfficePooledBytes content));
        using (content) return content.Array.AsSpan(0, content.Length).ToArray();
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

    // Members whose content differs. On .NET 8 and later untouched members are copied as stored, so their compressed
    // bytes must match too; older targets save through ZipArchive and recompress them.
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
            Assert.True(left.TryRead(name, out OfficePooledBytes leftContent));
            Assert.True(right.TryRead(name, out OfficePooledBytes rightContent));
            using (a)
            using (b)
            using (leftContent)
            using (rightContent) {
                bool same = leftContent.Array.AsSpan(0, leftContent.Length).SequenceEqual(rightContent.Array.AsSpan(0, rightContent.Length));
#if NET8_0_OR_GREATER
                same = same && a.Crc == b.Crc && a.Method == b.Method
                    && a.Compressed.Array.AsSpan(0, a.Compressed.Length).SequenceEqual(b.Compressed.Array.AsSpan(0, b.Compressed.Length));
#endif
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
