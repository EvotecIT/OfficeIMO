namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyResourceBoundRegressionTests {
    [Fact]
    public void ReferenceResolutionPreservesDistinctNativeFieldsAndReplacementOrder() {
        var source = new BibliographyDocument(BibliographyFormat.BibLatex);
        var child = new BibliographyItem { Key = "child", Type = BibliographyItemType.Book };
        child.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.BibLatex, "xdata", "parent"));
        child.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.BibLatex, "f0001", "child value"));
        var parent = new BibliographyItem { Key = "parent", NativeType = "xdata" };
        for (int index = 0; index < 1_000; index++)
            parent.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.BibLatex, $"f{index:D4}", $"value {index}"));
        source.Items.Add(child);
        source.Items.Add(parent);

        BibliographyItem resolved = source.ResolveReferences().Document.Items[0];

        Assert.Equal(1_000, resolved.NativeFields.Count(field => field.Name.StartsWith("f", StringComparison.Ordinal)));
        Assert.Equal("value 1", Assert.Single(resolved.NativeFields, field => field.Name == "f0001").Value);
        Assert.Equal("f0000", resolved.NativeFields.First(field => field.Name.StartsWith("f", StringComparison.Ordinal)).Name);
        Assert.Equal("child value", child.NativeFields[1].Value);

        BibliographyItem withoutReplacement = source.ResolveReferences(new BibliographyReferenceOptions { XDataOverridesExistingFields = false }).Document.Items[0];
        Assert.Equal("child value", Assert.Single(withoutReplacement.NativeFields, field => field.Name == "f0001").Value);
    }

    [Theory]
    [InlineData("")]
    [InlineData(" second-field-align=\"flush\"")]
    public void LayoutStopsWhenCumulativeIntermediateBudgetIsExceeded(string alignment) {
        const string header = "<style xmlns=\"http://purl.org/net/xbiblio/csl\" version=\"1.0\" class=\"in-text\"><citation><layout><text variable=\"title\"/></layout></citation>";
        string fields = string.Concat(Enumerable.Repeat("<text variable=\"title\"/>", 200));
        CslStyle style = CslStyle.Parse(header + "<bibliography" + alignment + "><layout>" + fields + "</layout></bibliography></style>");
        BibliographyDocument source = BibliographyDocument.Parse("[{\"id\":\"one\",\"type\":\"book\",\"title\":\"Alpha\"}]", BibliographyFormat.CslJson).Document;
        var processor = new CslProcessor(source, style, new CslRenderOptions { MaximumIntermediateCharacters = 20, MaximumRenderingOperations = 100 });

        InvalidDataException error = Assert.Throws<InvalidDataException>(() => processor.RenderBibliography());
        Assert.Contains("MaximumIntermediateCharacters", error.Message);
    }
}
