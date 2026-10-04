namespace OfficeIMO.AsciiDoc.Tests;

public sealed class AsciiDocAttributeSubstitutionTests {
    [Fact]
    public void DocumentAttributes_RespectSourceOrderSetAndUnset() {
        const string source = ":product: OfficeIMO\n:edition: {product} Pro\n:product!:\n";
        AsciiDocDocumentAttributes attributes = AsciiDocDocument.ParseResult(source).Document.GetAttributes(
            new Dictionary<string, string> { ["initial"] = "yes" });

        Assert.False(attributes.Contains("product"));
        Assert.Equal("{product} Pro", attributes.GetValueOrDefault("edition"));
        Assert.Equal("yes", attributes.GetValueOrDefault("INITIAL"));
    }

    [Fact]
    public void ManyAttributeSnapshotsRetainEarlierValuesAfterLaterUpdatesAndUnsets() {
        var source = new System.Text.StringBuilder();
        for (int index = 0; index < 512; index++)
            source.Append(":key").Append(index).Append(": value").Append(index).Append('\n');
        source.Append(":KEY0: updated\n:key1!:\n");

        AsciiDocBlockContext[] contexts = AsciiDocDocument.Parse(source.ToString())
            .GetBlockContexts().ToArray();
        Assert.Equal(514, contexts.Length);
        Assert.Equal("value0", contexts[0].Attributes.GetValueOrDefault("KEY0"));
        Assert.False(contexts[0].Attributes.Contains("key511"));
        Assert.Equal("value0", contexts[511].Attributes.GetValueOrDefault("key0"));
        Assert.Equal(512, contexts[511].Attributes.Count);
        Assert.Equal("updated", contexts[512].Attributes.GetValueOrDefault("key0"));
        Assert.Equal(511, contexts[513].Attributes.Count);
        Assert.False(contexts[513].Attributes.Contains("KEY1"));
        Assert.Equal(511, contexts[513].Attributes.Values.Count);
    }

    [Fact]
    public void Substitution_IsRecursiveCaseInsensitiveAndBounded() {
        AsciiDocDocumentAttributes attributes = AsciiDocDocument.ParseResult(
            ":product: OfficeIMO\n:edition: {PRODUCT} Pro\n").Document.GetAttributes();

        AsciiDocAttributeSubstitutionResult result = AsciiDocAttributeSubstitutor.Substitute("Use {edition}.", attributes);

        Assert.Equal("Use OfficeIMO Pro.", result.Value);
        Assert.Empty(result.Diagnostics);
    }

    [Fact]
    public void UndefinedAndCyclicReferences_AreDiagnosedWithoutDataLoss() {
        AsciiDocDocumentAttributes attributes = AsciiDocDocument.ParseResult(":a: {b}\n:b: {a}\n").Document.GetAttributes();

        AsciiDocAttributeSubstitutionResult result = AsciiDocAttributeSubstitutor.Substitute("{a} {missing}", attributes);

        Assert.Equal("{a} {missing}", result.Value);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "ADOCEVAL001");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "ADOCEVAL002");
    }

    [Fact]
    public void EscapedReference_RemainsLiteralWithoutDiagnostic() {
        AsciiDocDocumentAttributes attributes = AsciiDocDocument.ParseResult(":name: value\n").Document.GetAttributes();

        AsciiDocAttributeSubstitutionResult result = AsciiDocAttributeSubstitutor.Substitute("\\{name} {name}", attributes);

        Assert.Equal("{name} value", result.Value);
        Assert.Empty(result.Diagnostics);
    }
}
