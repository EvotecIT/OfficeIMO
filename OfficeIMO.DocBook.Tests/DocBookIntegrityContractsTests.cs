using System;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.DocBook.Tests;

public sealed class DocBookIntegrityContractsTests {
    [Fact]
    public void IndependentlyProducedDocumentPreservesBytesAndProjectsItsContent() {
        byte[] bytes = System.IO.File.ReadAllBytes(System.IO.Path.Combine(AppContext.BaseDirectory, "Fixtures", "pandoc-3.12-common-structure.docbook"));
        using var stream = new System.IO.MemoryStream(bytes);
        var document = DocBookDocument.Load(stream);
        Assert.True(document.Validate().IsValid);
        Assert.Contains(document.Validate().Diagnostics, item => item.Code == "DB003"); // Producer declares 5.0; bounded profile is 5.2.
        Assert.Equal("Producer document", document.Title);
        var model = document.ToOfficeDocumentModel().Value;
        Assert.Contains(model.Links, link => link.Uri == "https://example.com/guide");
        Assert.Equal(2, Assert.Single(model.Tables).Rows.Count);
        using var written = new System.IO.MemoryStream();
        document.Write(written);
        Assert.Equal(bytes, written.ToArray());
    }
    [Theory]
    [InlineData(DocBookProfile.DocBook45)]
    [InlineData(DocBookProfile.DocBook52)]
    public void TypedBodyAddedAfterSectionsIsInsertedBeforeSubdivisions(DocBookProfile profile) {
        var document = DocBookDocument.CreateArticle(profile);
        document.Title = "Guide";
        var first = document.AddSection("First");
        first.AddSection("Nested").AddParagraph("Child");
        first.AddParagraph("Section introduction");
        document.AddSection("Second").AddParagraph("Second body");
        document.AddParagraph("Article introduction");
        document.Root.AddProgramListing("code");
        Assert.True(document.Validate().IsValid);
        Assert.Equal(new[] { profile == DocBookProfile.DocBook52 ? "info" : "articleinfo", "para", "programlisting", "section", "section" },
            document.Xml.Root!.Elements().Select(node => node.Name.LocalName));
        Assert.Equal(new[] { "title", "para", "section" }, first.Children.Select(node => node.Name));
    }

    [Fact]
    public void RawBodyAfterSectionsIsReportedWithoutRewritingNativeXml() {
        const string xml = "<article xmlns='http://docbook.org/ns/docbook' version='5.2'><title>Guide</title><section><title>Section</title><para>Body</para></section><para>Late</para></article>";
        var document = DocBookDocument.Parse(xml);
        Assert.Contains(document.Validate().Diagnostics, item => item.Code == "DB024" && item.Severity == DocBookDiagnosticSeverity.Error);
        Assert.Equal(xml, document.ToDocBook());
    }

    [Fact]
    public void TableProjectionSeparatesParagraphsWithoutSplittingInlineRuns() {
        const string xml = "<article xmlns='http://docbook.org/ns/docbook' version='5.2'><title>Guide</title><informaltable><tgroup cols='1'><tbody><row><entry><para>one<emphasis> and</emphasis></para><para>two</para></entry></row></tbody></tgroup></informaltable></article>";
        var document = DocBookDocument.Parse(xml);
        var result = document.ToOfficeDocumentModel();
        Assert.Equal("one and\ntwo", Assert.Single(Assert.Single(result.Value.Tables).Rows)[0]);
        Assert.Equal(xml, document.ToDocBook());
    }

    [Fact]
    public void BlockSeparatorsConsumeTheProjectionTextBudget() {
        const string xml = "<article xmlns='http://docbook.org/ns/docbook' version='5.2'><informaltable><tgroup cols='1'><tbody><row><entry><para>one</para><para>two</para></entry></row></tbody></tgroup></informaltable></article>";
        var result = DocBookDocument.Parse(xml).ToOfficeDocumentModel(options: new DocBookConversionOptions { MaxTotalTextCharacters = 6 });
        Assert.Contains(result.Diagnostics, item => item.Code == "DB123");
    }

    [Fact]
    public void ChangeTrackingPreservesUndoAndDeclarationEdits() {
        const string xml = "<?xml version='1.0' encoding='utf-8'?><article xmlns='http://docbook.org/ns/docbook' version='5.2'><para>Before</para></article>";
        var document = DocBookDocument.Parse(xml);
        Assert.False(document.IsModified);
        var paragraph = document.Xml.Descendants().Single(node => node.Name.LocalName == "para");
        paragraph.Value = "After";
        Assert.True(document.IsModified);
        Assert.True(document.IsModified);
        paragraph.Value = "Before";
        Assert.False(document.IsModified);
        Assert.Equal(xml, document.ToDocBook());
        document.Xml.Declaration!.Standalone = "yes";
        Assert.True(document.IsModified);
        Assert.Contains("standalone=\"yes\"", document.ToDocBook());
        document.Xml.Declaration.Standalone = null;
        Assert.False(document.IsModified);
        document.Xml.Declaration = new XDeclaration("1.0", "utf-8", "no");
        Assert.True(document.IsModified);
    }

    [Fact]
    public void CalsOrderingStillRejectsEarlierPhasesAfterLaterOnes() {
        const string xml = "<article xmlns='http://docbook.org/ns/docbook' version='5.2'><informaltable><tgroup cols='1'><tbody><row><entry>x</entry></row></tbody><colspec colname='late'/></tgroup></informaltable></article>";
        var result = DocBookDocument.Parse(xml).Validate();
        Assert.Contains(result.Diagnostics, item => item.Code == "DB020");
    }
}
