namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave55RegressionTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EndNote_custom_root_detection_observes_cancellation_inside_large_tokens(bool attribute) {
        string large = new string('x', 64 * 1024 * 1024);
        string source = attribute
            ? "<library metadata=\"" + large + "\"><records/></library>"
            : "<library><metadata>" + large + "</metadata><records/></library>";
        var options = new BibliographyReadOptions { MaximumInputCharacters = source.Length + 1, MaximumValueLength = source.Length + 1 };

        BibliographyCancellationTest.AssertObserved(token => BibliographyDocument.Parse(source, options, token));
    }

    [Theory]
    [InlineData("title")]
    [InlineData("pages")]
    [InlineData("date")]
    public void EndNote_typed_element_writes_observe_cancellation_inside_large_values(string owner) {
        string value = new string('x', 64 * 1024 * 1024);
        var document = new BibliographyDocument(BibliographyFormat.EndNoteXml);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.Book };
        if (owner == "title") item.Title = value;
        else if (owner == "pages") item.Pages = value;
        else item.Dates.Add(new BibliographyDate { Role = BibliographyDateRole.Issued, Literal = value });
        document.Items.Add(item);

        BibliographyCancellationTest.AssertObserved(token =>
            EndNoteXmlCodec.Write(document, new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical }, new BibliographyConversionReport(), token));
    }

    [Theory]
    [InlineData("element", false)]
    [InlineData("element", true)]
    [InlineData("records-element", false)]
    [InlineData("records-element", true)]
    public void EndNote_native_extension_chunks_keep_surrogate_pairs_together(string kind, bool comment) {
        string value = new string('x', 4095) + "😀";
        string xml = comment ? "<metadata><!--" + value + "--></metadata>" : "<metadata>" + value + "</metadata>";
        var document = new BibliographyDocument(BibliographyFormat.EndNoteXml);
        document.NativeEntries.Add(new BibliographyNativeEntry(BibliographyFormat.EndNoteXml, kind, xml, "metadata"));
        document.Items.Add(new BibliographyItem { Key = "1", Type = BibliographyItemType.Book, Title = "Title" });

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyNativeEntry reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document.NativeEntries,
            entry => entry.Kind == kind && entry.Name == "metadata");

        Assert.Contains("😀", reopened.Value, StringComparison.Ordinal);
    }
}
