using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookOnixTranslationLanguageTests {
    [Theory]
    [InlineData("subject")]
    [InlineData("edition")]
    [InlineData("text")]
    [InlineData("xhtml")]
    [InlineData("source-title")]
    public void RepeatedTranslationsRequireEveryLanguageBeforeMutation(string field) {
        var project = BookOnixTests.Project();
        byte[] before = project.ToProjectBytes();
        Assert.Throws<ArgumentException>(() => project.ExportOnix(Options(field, true, false), BookOnixTests.TestSchema()));
        Assert.Equal(before, project.ToProjectBytes());
    }

    [Theory]
    [InlineData("subject", "SubjectHeadingText")]
    [InlineData("edition", "EditionStatement")]
    [InlineData("text", "Text")]
    [InlineData("xhtml", "Text")]
    [InlineData("source-title", "SourceTitle")]
    public void SingletonMayOmitLanguageAndTranslationsPreserveTheirLanguages(string field, string element) {
        var project = BookOnixTests.Project();
        byte[] before = project.ToProjectBytes();
        XNamespace ns = BookProject.OnixNamespace;
        var single = project.ExportOnix(Options(field, false, false), BookOnixTests.TestSchema());
        Assert.Null(XDocument.Load(new MemoryStream(single.Bytes)).Descendants(ns + element).Single().Attribute("language"));
        var translated = project.ExportOnix(Options(field, true, true), BookOnixTests.TestSchema());
        Assert.Equal(new[] { "eng", "pol" }, XDocument.Load(new MemoryStream(translated.Bytes)).Descendants(ns + element).Select(e => (string?)e.Attribute("language")));
        Assert.Equal(single.Publication.Bytes, translated.Publication.Bytes);
        Assert.Equal(before, project.ToProjectBytes());
    }

    private static BookOnixExportOptions Options(string field, bool repeated, bool complete) {
        string? firstLanguage = repeated ? "eng" : null;
        string? secondLanguage = complete ? "pol" : null;
        var options = BookOnixTests.Options();
        if (field == "subject") return options with { Subjects = [new() { Scheme = BookOnixSubjectScheme.Keywords,
            Headings = repeated ? [new("Learning", firstLanguage), new("Nauka", secondLanguage)] : [new("Learning")] }] };
        if (field == "edition") return options with { Edition = new() {
            Statements = repeated ? [new("Revised edition", firstLanguage), new("Wydanie poprawione", secondLanguage)] : [new("Revised edition")]
        } };
        var format = field == "xhtml" ? BookOnixCollateralTextFormat.Xhtml : BookOnixCollateralTextFormat.PlainText;
        string first = field == "xhtml" ? "<p>First</p>" : "First";
        string second = field == "xhtml" ? "<p>Drugi</p>" : "Drugi";
        BookOnixCollateralTextValue[] values = repeated ? [new(first, firstLanguage) { Format = format }, new(second, secondLanguage) { Format = format }]
            : [new(first) { Format = format }];
        return options with { CollateralTexts = [new() { Type = BookOnixTextType.Description,
            Audiences = [BookOnixContentAudience.Unrestricted],
            Texts = field == "source-title" ? [new("Description")] : values,
            SourceTitles = field == "source-title" ? values : []
        }] };
    }
}
