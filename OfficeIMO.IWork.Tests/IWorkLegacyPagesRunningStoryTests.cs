using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkLegacyPagesRunningStoryTests {
    private const string StoryLoss = "IWORK_LEGACY_PAGES_HEADER_FOOTER_OMITTED";
    private static readonly XNamespace Sf = "http://developer.apple.com/namespaces/sf";
    private static readonly XNamespace Sfa = "http://developer.apple.com/namespaces/sfa";
    private static readonly XNamespace Sl = "http://developer.apple.com/namespaces/sl";

    [Fact]
    public void Pages09DefaultRunningStoriesSurviveEditableWordSaveAndReopen() {
        using IWorkSourceDocument source = IWorkSourceDocument.Load(Fixture("header+footer.pages"));
        IWorkPagesProjection projection = source.ReadPages();
        IWorkPagesSection section = Assert.Single(projection.Sections);
        string expectedHeader = ReferenceStory("header");
        string expectedFooter = ReferenceStory("footer");
        Assert.Equal("Header.", expectedHeader);
        Assert.Equal("Footer.", expectedFooter);
        Assert.Equal(expectedHeader, Assert.Single(section.DefaultPageHeaderContents).PlainText);
        Assert.Equal(expectedFooter, Assert.Single(section.DefaultPageFooterContents).PlainText);
        Assert.False(section.DifferentFirstPage);
        Assert.False(section.DifferentOddAndEvenPages);
        Assert.Empty(section.FirstPageHeaderContents);
        Assert.Empty(section.EvenPageFooterContents);
        Assert.Equal(string.Empty, Assert.Single(projection.Body.Paragraphs).Text);
        Assert.DoesNotContain(projection.Diagnostics, diagnostic => diagnostic.Code == StoryLoss);

        using var result = source.ToWordResult(PartialConversion());
        using var output = new MemoryStream();
        result.Value.Save(output);
        byte[] savedBytes = output.ToArray();
        using (var saved = WordprocessingDocument.Open(new MemoryStream(savedBytes), false)) {
            var main = saved.MainDocumentPart!;
            W.Header header = Assert.Single(main.HeaderParts).Header!;
            W.Footer footer = Assert.Single(main.FooterParts).Footer!;
            Assert.Equal(expectedHeader, header.InnerText);
            Assert.Equal(expectedFooter, footer.InnerText);
            Assert.Equal("20", Assert.Single(header.Descendants<W.Run>(), run => run.InnerText.Length > 0)
                .RunProperties!.FontSize!.Val!.Value);
            Assert.Equal("000000", Assert.Single(footer.Descendants<W.Run>(), run => run.InnerText.Length > 0)
                .RunProperties!.Color!.Val!.Value);
            Assert.DoesNotContain(expectedHeader, main.Document!.Body!.InnerText);
            Assert.DoesNotContain(expectedFooter, main.Document!.Body!.InnerText);
        }

        using WordDocument reopened = WordDocument.Load(new MemoryStream(savedBytes));
        Assert.Equal(expectedHeader, Assert.Single(reopened.PageSetup.Header.Paragraphs,
            paragraph => paragraph.Text.Length > 0).Text);
        reopened.PageSetup.Header.AddParagraph("Edited header");
        using var edited = new MemoryStream();
        reopened.Save(edited);
        edited.Position = 0;
        using var editedDocument = WordprocessingDocument.Open(edited, false);
        Assert.Contains(Assert.Single(editedDocument.MainDocumentPart!.HeaderParts).Header!.Elements<W.Paragraph>(),
            paragraph => paragraph.InnerText == "Edited header");
        Assert.Equal(expectedHeader, Assert.Single(section.DefaultPageHeaderContents).PlainText);
    }

    [Theory]
    [InlineData("first-page")]
    [InlineData("even-page")]
    [InlineData("invalid-selection")]
    [InlineData("multiple-sections")]
    [InlineData("unresolved-style")]
    public void Pages09UnqualifiedRunningStorySelectionReportsOmissionAndRetainsBody(string selection) {
        using MemoryStream input = DerivedInput(root => {
            XElement body = Body(root);
            body.Descendants(Sf + "p").Single().Value = "Body control";
            XElement section = body.Element(Sf + "section")!;
            XElement style = SectionProperties(root);
            if (selection == "first-page") SetFlag(style, Sf, "pageMasterFirstPageDifferent", "1");
            else if (selection == "even-page") SetFlag(style, Sf, "pageMasterLeftRightPagesDifferent", "1");
            else if (selection == "invalid-selection") SetFlag(style, Sf, "pageMasterFirstPageDifferent", "invalid");
            else if (selection == "unresolved-style") section.SetAttributeValue(Sf + "style", "missing-style");
            else body.Add(new XElement(section));
        });
        using IWorkSourceDocument source = IWorkSourceDocument.Load(input);
        IWorkPagesProjection projection = source.ReadPages();
        Assert.Empty(projection.Sections);
        Assert.All(projection.Body.Paragraphs, paragraph => Assert.Equal("Body control", paragraph.Text));
        Assert.NotEmpty(projection.Body.Paragraphs);
        Assert.Single(projection.Diagnostics, diagnostic => diagnostic.Code == StoryLoss
            && diagnostic.LossKind == OfficeConversionLossKind.Omission);
        using var result = source.ToWordResult(PartialConversion());
        Assert.Single(result.Report.Diagnostics, diagnostic => diagnostic.Code == StoryLoss
            && diagnostic.LossKind == OfficeConversionLossKind.Omission);
        Assert.Equal(projection.Body.Paragraphs.Select(paragraph => paragraph.Text),
            result.Value.OpenXmlDocument.MainDocumentPart!.Document!.Body!.Elements<W.Paragraph>()
                .Select(paragraph => paragraph.InnerText));
    }

    [Fact]
    public void Pages09DisabledAndInactiveRunningStoriesRemainUnselected() {
        using MemoryStream input = DerivedInput(root => {
            SetFlag(root.Element(Sl + "publication-info")!, Sl, "kSFWPHasHeadersProperty", "0");
            foreach (XElement story in root.Element(Sf + "headers")!.Elements()
                .Concat(root.Element(Sf + "footers")!.Elements())) {
                if ((string?)story.Attribute(Sf + "name") is string name && !name.Contains("Odd"))
                    story.Descendants(Sf + "p").Single().Value = "Inactive story";
            }
        });
        using IWorkSourceDocument source = IWorkSourceDocument.Load(input);
        IWorkPagesProjection projection = source.ReadPages();
        IWorkPagesSection section = Assert.Single(projection.Sections);
        Assert.Empty(section.SelectedHeaderContents);
        Assert.Equal("Footer.", Assert.Single(section.SelectedFooterContents).PlainText);
        Assert.Empty(section.FirstPageFooterContents);
        Assert.Empty(section.EvenPageFooterContents);
        Assert.DoesNotContain(projection.Diagnostics, diagnostic => diagnostic.Code == StoryLoss);
        using var result = source.ToWordResult(PartialConversion());
        using var output = new MemoryStream();
        result.Value.Save(output);
        output.Position = 0;
        using var saved = WordprocessingDocument.Open(output, false);
        Assert.All(saved.MainDocumentPart!.HeaderParts, part => Assert.Equal(string.Empty, part.Header!.InnerText));
        Assert.Equal("Footer.", Assert.Single(saved.MainDocumentPart.FooterParts).Footer!.InnerText);
    }

    [Theory]
    [InlineData("missing")]
    [InlineData("duplicate")]
    public void Pages09UnresolvedDefaultHeaderRetainsItsQualifiedFooter(string storage) {
        using MemoryStream input = DerivedInput(root => {
            XElement headers = root.Element(Sf + "headers")!;
            XElement selected = headers.Elements(Sf + "header").Single(element =>
                (string?)element.Attribute(Sf + "name") == "SFWPDefaultOddHeaderIdentifier");
            if (storage == "missing") selected.Remove();
            else headers.Add(new XElement(selected));
        });
        using IWorkSourceDocument source = IWorkSourceDocument.Load(input);
        IWorkPagesProjection projection = source.ReadPages();
        IWorkPagesSection section = Assert.Single(projection.Sections);
        Assert.Empty(section.SelectedHeaderContents);
        Assert.Equal("Footer.", Assert.Single(section.SelectedFooterContents).PlainText);
        Assert.Single(projection.Diagnostics, diagnostic => diagnostic.Code == StoryLoss
            && diagnostic.LossKind == OfficeConversionLossKind.Omission);
    }

    [Fact]
    public void Pages09RunningStoriesShareTheBodyProjectionTextBudget() {
        using IWorkSourceDocument boundary = IWorkSourceDocument.Load(Fixture("header+footer.pages"),
            new IWorkLoadOptions { MaxProjectedTextCharacters = 14 });
        Assert.Equal(2, boundary.ReadPages().Sections.Single().SelectedHeaderContents.Count
            + boundary.ReadPages().Sections.Single().SelectedFooterContents.Count);
        using IWorkSourceDocument limited = IWorkSourceDocument.Load(Fixture("header+footer.pages"),
            new IWorkLoadOptions { MaxProjectedTextCharacters = 13 });
        Assert.Throws<InvalidDataException>(() => limited.ReadPages());
    }

    private static string ReferenceStory(string kind) {
        XNamespace style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
        XNamespace text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
        XDocument reference = XDocument.Load(Fixture("header+footer.pages.fodt"));
        return reference.Descendants(style + kind).Single().Descendants(text + "p").Single().Value.Trim();
    }

    private static IWorkConversionOptions PartialConversion() => new() {
        Mode = IWorkConversionMode.EditableOnly,
        AllowPartialEditableReconstruction = true
    };

    private static MemoryStream DerivedInput(Action<XElement> change) {
        var stream = new MemoryStream();
        byte[] bytes = File.ReadAllBytes(Fixture("header+footer.pages"));
        stream.Write(bytes, 0, bytes.Length);
        stream.Position = 0;
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true)) {
            ZipArchiveEntry entry = zip.GetEntry("index.xml")!;
            XDocument document;
            using (Stream xml = entry.Open()) document = XDocument.Load(xml, LoadOptions.PreserveWhitespace);
            change(document.Root!);
            entry.Delete();
            using Stream output = zip.CreateEntry("index.xml").Open();
            document.Save(output);
        }
        stream.Position = 0;
        return stream;
    }

    private static XElement Body(XElement root) => root.Element(Sf + "text-storage")!.Element(Sf + "text-body")!;

    private static XElement SectionProperties(XElement root) => root.Element(Sl + "stylesheet")!
        .Descendants(Sf + "sectionstyle").Single(style => (string?)style.Attribute(Sf + "ident") == "section-style-10")
        .Element(Sf + "property-map")!;

    private static void SetFlag(XElement owner, XNamespace ns, string property, string value) =>
        owner.Element(ns + property)!.Elements().Single().SetAttributeValue(Sfa + "number", value);

    private static string Fixture(string name) => Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus",
        "legacy-pages-running-stories", name);
}
