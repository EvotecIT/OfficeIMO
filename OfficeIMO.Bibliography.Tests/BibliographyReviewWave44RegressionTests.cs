using System.Xml.Linq;

namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave44RegressionTests {
    [Theory]
    [InlineData("Record", false)]
    [InlineData("RECORD", true)]
    public void Accepted_EndNote_record_names_survive_strict_canonical_edits(string recordName, bool withAttributes) {
        string attributes = withAttributes ? " custom=\"value\"" : string.Empty;
        string source = "<xml><records><" + recordName + attributes + "><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><titles><title>Before</title></titles></" + recordName + "></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        document.Items[0].Title = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyDocument reopened = BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document;

        XElement record = Assert.Single(XDocument.Parse(written.Content).Root!.Element("records")!.Elements());
        Assert.Equal(recordName, record.Name.LocalName);
        Assert.Equal("After", Assert.Single(reopened.Items).Title);
        BibliographyNativeField carrier = Assert.Single(reopened.Items[0].NativeFields, field => field.Name == "@record-attributes");
        Assert.StartsWith("<" + recordName, carrier.Value, StringComparison.Ordinal);
    }

    [Fact]
    public void Additional_EndNote_URLs_keep_their_position_among_native_fields() {
        const string source = "<xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><custom>first</custom><urls><related-urls><url>primary</url><url>extra</url></related-urls></urls><other>last</other></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        BibliographyItem item = Assert.Single(document.Items);
        string[] originalOrder = item.NativeFields.Select(field => field.Name + "=" + field.Value).ToArray();
        item.Title = "Edited";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document.Items);

        Assert.Equal(originalOrder, reopened.NativeFields.Select(field => field.Name + "=" + field.Value));
    }
}
