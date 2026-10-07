using OfficeIMO.Epub;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubDescendantLanguageContracts {
    [Theory]
    [InlineData("en_US", null, false)]
    [InlineData("en", "pl", false)]
    [InlineData(" ", null, false)]
    [InlineData("", "", true)]
    [InlineData("fr-CA", "fr-ca", true)]
    [InlineData(null, "x-house", true)]
    public void DescendantDeclarationsAreCheckedWithoutMutatingImportedBytes(string? language, string? xmlLanguage, bool valid) {
        var book = Book(); var xml = book.GetContentXml("one");
        var target = xml.Descendants().Single(e => (string?)e.Attribute("id") == "passage");
        target.SetAttributeValue("lang", language); target.SetAttributeValue(XNamespace.Xml + "lang", xmlLanguage);
        book.SetContentXml("one", xml); byte[] bytes = book.Write().Bytes;
        var loaded = EpubPublication.Load(new MemoryStream(bytes));
        var check = loaded.Preflight().Checks.Single(c => c.Code == "document-language");
        Assert.Equal(valid ? EpubPreflightStatus.Passed : EpubPreflightStatus.Failed, check.Status);
        if (!valid) Assert.Contains(check.Diagnostics, d => d.Path == "EPUB/one.xhtml" && d.Message.Contains("passage"));
        Assert.Equal(bytes, loaded.Write().Bytes);
    }

    [Fact]
    public void MergeCannotHideInvalidLanguageByMovingItOffTheRoot() {
        var book = Book(); book.AddChapter("two", "EPUB/two.xhtml", "Second", "<h1>Second</h1>");
        var second = book.GetContentXml("two"); second.Root!.SetAttributeValue(XNamespace.Xml + "lang", "en_US"); book.SetContentXml("two", second);
        Assert.Equal(EpubPreflightStatus.Failed, Language(book).Status);
        book.MergeChapters("one", "two", "second-start", new EpubChapterMergeOptions { PreserveSecondChapterLanguageAndDirection = true });
        Assert.Equal(EpubPreflightStatus.Failed, Language(book).Status);
    }

    [Fact]
    public void SvgAndForeignXmlLanguageAreCheckedButUnrelatedLangAttributesAreNot() {
        var book = Book(); var xml = book.GetContentXml("one"); XNamespace custom = "urn:example";
        xml.Root!.Element(XName.Get("body", "http://www.w3.org/1999/xhtml"))!.Add(new XElement(custom + "annotation", new XAttribute("lang", "not_a_language")));
        book.SetContentXml("one", xml);
        Assert.Equal(EpubPreflightStatus.Passed, Language(book).Status);
        book.AddResource("svg", "EPUB/page.svg", "image/svg+xml", System.Text.Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' xml:lang='en'><text id='label' xml:lang='en_US'>Words</text></svg>"));
        Assert.Contains(Language(book).Diagnostics, d => d.Path == "EPUB/page.svg" && d.Message.Contains("label"));
        var annotation = xml.Descendants(custom + "annotation").Single(); annotation.SetAttributeValue(XNamespace.Xml + "lang", "en_US"); book.SetContentXml("one", xml);
        Assert.Contains(Language(book).Diagnostics, d => d.Path == "EPUB/one.xhtml" && d.Message.Contains("annotation"));
    }

    private static EpubPreflightCheck Language(EpubPublication book) => book.Preflight().Checks.Single(c => c.Code == "document-language");
    private static EpubPublication Book() {
        var book = EpubPublication.Create("Language preflight", "en");
        book.AddChapter("one", "EPUB/one.xhtml", "Chapter", "<h1>Chapter</h1><p id='passage'>Passage</p>"); return book;
    }
}
