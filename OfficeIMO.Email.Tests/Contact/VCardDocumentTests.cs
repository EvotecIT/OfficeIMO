using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class VCardDocumentTests {
    [Fact]
    public void Parse_MultipleVersionsPreservesGroupsRepeatedValuesAndMedia() {
        const string source = "BEGIN:VCARD\r\nVERSION:4.0\r\nFN:Team contact\r\nKIND:group\r\n" +
            "MEMBER:urn:uuid:a\r\nMEMBER:urn:uuid:b\r\nitem1.EMAIL;TYPE=work;PREF=1:team@example.com\r\n" +
            "PHOTO:data:image/png;base64,AAEC\r\nX-VENDOR-FIELD:retained\r\nEND:VCARD\r\n" +
            "BEGIN:VCARD\r\nVERSION:2.1\r\nFN:Legacy contact\r\nTEL;HOME:12345\r\nEND:VCARD\r\n";

        VCardDocument document = VCardDocument.Parse(source);
        ContentLineComponent first = document.Cards[0];
        first.AddProperty("EMAIL", "alternate@example.com").SetParameter("TYPE", "home");

        VCardDocument reparsed = VCardDocument.Parse(document.Serialize());
        ContentLineComponent reparsedFirst = reparsed.Cards[0];

        Assert.Equal(2, reparsed.Cards.Count);
        Assert.Equal(VCardVersion.V4_0, VCardDocument.GetVersion(reparsedFirst));
        Assert.Equal(VCardVersion.V2_1, VCardDocument.GetVersion(reparsed.Cards[1]));
        Assert.Equal(new[] { "urn:uuid:a", "urn:uuid:b" },
            reparsedFirst.GetProperties("MEMBER").Select(property => property.Value));
        Assert.Equal("item1", reparsedFirst.GetProperties("EMAIL").First().Group);
        Assert.Equal("data:image/png;base64,AAEC", reparsedFirst.GetFirstProperty("PHOTO")!.Value);
        Assert.Equal("retained", reparsedFirst.GetFirstProperty("X-VENDOR-FIELD")!.Value);
        Assert.Equal(2, reparsedFirst.GetProperties("EMAIL").Count());
    }

    [Theory]
    [InlineData(VCardVersion.V2_1, "2.1")]
    [InlineData(VCardVersion.V3_0, "3.0")]
    [InlineData(VCardVersion.V4_0, "4.0")]
    public void AddCard_WritesRequestedVersion(VCardVersion version, string expected) {
        var document = new VCardDocument();
        document.Cards.Clear();
        ContentLineComponent card = document.AddCard(version);
        card.AddProperty("FN", "Versioned contact");

        VCardDocument reparsed = VCardDocument.Parse(document.Serialize());

        Assert.Equal(expected, reparsed.Cards.Single().GetFirstProperty("VERSION")!.Value);
    }

    [Fact]
    public void Parse_RejectsMissingOrUnsupportedVersion() {
        Assert.Throws<InvalidDataException>(() => VCardDocument.Parse(
            "BEGIN:VCARD\r\nFN:No version\r\nEND:VCARD\r\n"));
        Assert.Throws<InvalidDataException>(() => VCardDocument.Parse(
            "BEGIN:VCARD\r\nVERSION:5.0\r\nFN:Future\r\nEND:VCARD\r\n"));
    }

    [Fact]
    public void ValidationAppliesVersionSpecificContractsWithoutRejectingExtensions() {
        var document = new VCardDocument();
        ContentLineComponent card = document.Cards.Single();
        card.AddProperty("FN", "Ada Lovelace");
        card.AddProperty("ANNIVERSARY", "18350708");
        card.AddProperty("EMAIL", "ada@example.com").SetParameter("PREF", "101");
        card.AddProperty("X-FUTURE-PROPERTY", "retained");

        IReadOnlyList<ContentLineValidationIssue> issues = document.Validate();

        Assert.Contains(issues, issue => issue.Code == "VCARD4_PREF_INVALID" &&
            issue.Severity == ContentLineValidationSeverity.Error);
        Assert.Contains(issues, issue => issue.Code == "VCARD_PROPERTY_REQUIRED" && issue.PropertyName == "N");
        Assert.DoesNotContain(issues, issue => issue.PropertyName == "X-FUTURE-PROPERTY");
    }

    [Fact]
    public void ValidationFlagsV3AnniversaryAndV4LegacyEncoding() {
        var v3 = new VCardDocument();
        VCardDocument.SetVersion(v3.Cards.Single(), VCardVersion.V3_0);
        v3.Cards.Single().AddProperty("FN", "Version three");
        v3.Cards.Single().AddProperty("N", "Three;Version;;;");
        v3.Cards.Single().AddProperty("ANNIVERSARY", "20260717");

        var v4 = new VCardDocument();
        v4.Cards.Single().AddProperty("FN", "Version four");
        v4.Cards.Single().AddProperty("N", "Four;Version;;;");
        v4.Cards.Single().AddProperty("NOTE", "legacy=20text")
            .SetParameter("ENCODING", "QUOTED-PRINTABLE").SetParameter("CHARSET", "windows-1252");

        Assert.Contains(v3.Validate(), issue => issue.Code == "VCARD_PROPERTY_VERSION_MISMATCH");
        Assert.Contains(v4.Validate(), issue => issue.Code == "VCARD4_ENCODING_FORBIDDEN");
        Assert.Contains(v4.Validate(), issue => issue.Code == "VCARD4_CHARSET_FORBIDDEN");
    }

    [Fact]
    public void GroupAndTextHelpersCreateInteroperableV4Card() {
        var document = new VCardDocument();
        ContentLineComponent group = document.AddGroup("Engineering, Europe",
            new[] { "urn:uuid:alice", "urn:uuid:bob" });
        group.AddVCardText("NOTE", "Line one\nLine two; retained");

        VCardDocument reparsed = VCardDocument.Parse(document.Serialize());
        ContentLineComponent reparsedGroup = reparsed.Cards.Single();

        Assert.Equal("Engineering, Europe", reparsedGroup.GetVCardText("FN"));
        Assert.Equal("Line one\nLine two; retained", reparsedGroup.GetVCardText("NOTE"));
        Assert.Equal(2, reparsedGroup.GetProperties("MEMBER").Count());
        Assert.DoesNotContain(reparsed.Validate(),
            issue => issue.Severity == ContentLineValidationSeverity.Error);
    }

    [Fact]
    public void SetVersionKeepsVersionAsTheFirstCardProperty() {
        var document = new VCardDocument();
        ContentLineComponent card = document.Cards.Single();
        card.Properties.Insert(0, new ContentLineProperty("FN", "Reordered"));

        VCardDocument.SetVersion(card, VCardVersion.V3_0);

        Assert.Equal("VERSION", card.Properties[0].Name);
        Assert.Equal("3.0", card.Properties[0].Value);
        Assert.DoesNotContain(document.Validate(), issue => issue.Code == "VCARD_VERSION_ORDER");
    }

    [Fact]
    public void V4AllowsMultipleFormattedNamesButV3DoesNot() {
        var v4 = new VCardDocument();
        v4.Cards.Single().AddProperty("FN", "Primary name");
        v4.Cards.Single().AddProperty("FN", "Nom français").SetParameter("LANGUAGE", "fr");
        var v3 = new VCardDocument();
        VCardDocument.SetVersion(v3.Cards.Single(), VCardVersion.V3_0);
        v3.Cards.Single().AddProperty("FN", "Primary name");
        v3.Cards.Single().AddProperty("FN", "Second name");

        Assert.DoesNotContain(v4.Validate(), issue =>
            issue.Code == "VCARD_PROPERTY_CARDINALITY" && issue.PropertyName == "FN");
        Assert.Contains(v3.Validate(), issue =>
            issue.Code == "VCARD_PROPERTY_CARDINALITY" && issue.PropertyName == "FN");
    }

    [Fact]
    public void ValidationAndSerializationRejectMissingOrMutatedCardRoots() {
        var empty = new VCardDocument();
        empty.Cards.Clear();
        var mutated = new VCardDocument();
        mutated.Cards.Single().Name = "VCALENDAR";

        Assert.Contains(empty.Validate(), issue => issue.Code == "VCARD_ROOT_REQUIRED");
        Assert.Contains(mutated.Validate(), issue => issue.Code == "VCARD_ROOT_INVALID");
        Assert.Throws<InvalidDataException>(() => empty.ToBytes());
        Assert.Throws<InvalidDataException>(() => mutated.ToBytes());
    }
}
