using System.Text.Json;

namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyCslNumericIdTests {
    [Theory]
    [InlineData("123", "123")]
    [InlineData("1.0", "1")]
    public void NumericCslIdentityIsTypedAndRetainsItsJsonRepresentation(string raw, string key) {
        string source = "[{\"id\":" + raw + ",\"type\":\"book\",\"title\":\"Before\"}]";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.CslJson).Document;
        Assert.Equal(key, document.Items.Single().Key);
        Assert.Equal(source, document.Write().Content);
        document.Items.Single().Title = "After";
        BibliographyWriteResult result = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        using JsonDocument json = JsonDocument.Parse(result.Content);
        Assert.Equal(JsonValueKind.Number, json.RootElement[0].GetProperty("id").ValueKind);
        Assert.Equal(raw, json.RootElement[0].GetProperty("id").GetRawText());
        Assert.Equal(key, BibliographyDocument.Parse(result.Content, BibliographyFormat.CslJson).Document.Items.Single().Key);
        document.Items.Single().Key = "new-key";
        using JsonDocument edited = JsonDocument.Parse(document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }).Content);
        Assert.Equal("new-key", edited.RootElement[0].GetProperty("id").GetString());
    }
}
