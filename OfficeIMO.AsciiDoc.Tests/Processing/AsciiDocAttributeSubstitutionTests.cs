namespace OfficeIMO.AsciiDoc.Tests;

public sealed class AsciiDocAttributeSubstitutionTests {
    [Fact]
    public void DocumentAttributes_RespectSourceOrderSetAndUnset() {
        const string source = ":product: OfficeIMO\n:edition: {product} Pro\n:product!:\n";
        AsciiDocDocumentAttributes attributes = AsciiDocDocument.Parse(source).Document.GetAttributes(
            new Dictionary<string, string> { ["initial"] = "yes" });

        Assert.False(attributes.Contains("product"));
        Assert.Equal("{product} Pro", attributes.GetValueOrDefault("edition"));
        Assert.Equal("yes", attributes.GetValueOrDefault("INITIAL"));
    }

    [Fact]
    public void Substitution_IsRecursiveCaseInsensitiveAndBounded() {
        AsciiDocDocumentAttributes attributes = AsciiDocDocument.Parse(
            ":product: OfficeIMO\n:edition: {PRODUCT} Pro\n").Document.GetAttributes();

        AsciiDocAttributeSubstitutionResult result = AsciiDocAttributeSubstitutor.Substitute("Use {edition}.", attributes);

        Assert.Equal("Use OfficeIMO Pro.", result.Value);
        Assert.Empty(result.Diagnostics);
    }

    [Fact]
    public void UndefinedAndCyclicReferences_AreDiagnosedWithoutDataLoss() {
        AsciiDocDocumentAttributes attributes = AsciiDocDocument.Parse(":a: {b}\n:b: {a}\n").Document.GetAttributes();

        AsciiDocAttributeSubstitutionResult result = AsciiDocAttributeSubstitutor.Substitute("{a} {missing}", attributes);

        Assert.Equal("{a} {missing}", result.Value);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "ADOCEVAL001");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "ADOCEVAL002");
    }

    [Fact]
    public void EscapedReference_RemainsLiteralWithoutDiagnostic() {
        AsciiDocDocumentAttributes attributes = AsciiDocDocument.Parse(":name: value\n").Document.GetAttributes();

        AsciiDocAttributeSubstitutionResult result = AsciiDocAttributeSubstitutor.Substitute("\\{name} {name}", attributes);

        Assert.Equal("{name} value", result.Value);
        Assert.Empty(result.Diagnostics);
    }
}
