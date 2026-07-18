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
    public void ValidationChecksEveryV4EncodingAndPreferenceParameterValue() {
        var document = new VCardDocument();
        ContentLineComponent card = document.Cards.Single();
        card.AddProperty("FN", "Version four");
        card.AddProperty("N", "Four;Version;;;");
        ContentLineProperty note = card.AddProperty("NOTE", "legacy=20text");
        note.Parameters.Add(new ContentLineParameter("ENCODING", "8BIT", "QUOTED-PRINTABLE"));
        ContentLineProperty photo = card.AddProperty("PHOTO", "legacy-photo");
        photo.Parameters.Add(new ContentLineParameter("ENCODING", "8BIT"));
        photo.Parameters.Add(new ContentLineParameter("ENCODING", "QP"));
        ContentLineProperty email = card.AddProperty("EMAIL", "four@example.test");
        email.Parameters.Add(new ContentLineParameter("PREF", "1"));
        email.Parameters.Add(new ContentLineParameter("PREF", "101"));

        IReadOnlyList<ContentLineValidationIssue> issues = document.Validate();

        Assert.Contains(issues, issue => issue.Code == "VCARD4_ENCODING_FORBIDDEN" &&
            issue.PropertyName == "NOTE");
        Assert.Contains(issues, issue => issue.Code == "VCARD4_ENCODING_FORBIDDEN" &&
            issue.PropertyName == "PHOTO");
        Assert.Contains(issues, issue => issue.Code == "VCARD4_PREF_CARDINALITY" &&
            issue.PropertyName == "EMAIL");
        Assert.Contains(issues, issue => issue.Code == "VCARD4_PREF_INVALID" &&
            issue.PropertyName == "EMAIL");
    }

    [Theory]
    [InlineData("8BIT")]
    [InlineData("b")]
    [InlineData("BASE64")]
    public void ValidationRejectsEveryV4EncodingParameter(string encoding) {
        var document = new VCardDocument();
        ContentLineComponent card = document.Cards.Single();
        card.AddProperty("FN", "Version four");
        card.AddProperty("N", "Four;Version;;;");
        card.AddProperty("PHOTO", "legacy-photo").SetParameter("ENCODING", encoding);

        Assert.Contains(document.Validate(), issue =>
            issue.Code == "VCARD4_ENCODING_FORBIDDEN" && issue.PropertyName == "PHOTO");
    }

    [Fact]
    public void ValidationRejectsNestedCardComponentsWithoutRemovingThem() {
        var document = new VCardDocument();
        ContentLineComponent card = document.Cards.Single();
        card.AddProperty("FN", "Nested card");
        card.AddProperty("N", "Card;Nested;;;");
        ContentLineComponent nested = card.AddComponent("X-VENDOR-COMPONENT");
        nested.AddProperty("X-VALUE", "retained");

        Assert.Contains(document.Validate(), issue =>
            issue.Code == "VCARD_COMPONENT_NESTING_INVALID" &&
            issue.ComponentName == "X-VENDOR-COMPONENT");
        Assert.Same(nested, Assert.Single(card.Components));
        Assert.Contains("BEGIN:X-VENDOR-COMPONENT", document.Serialize(), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(VCardVersion.V3_0, "BDAY", "not-a-date", null)]
    [InlineData(VCardVersion.V3_0, "BDAY", "2026-02-30", null)]
    [InlineData(VCardVersion.V3_0, "BDAY", "2026-07-18T09:00:00Z", null)]
    [InlineData(VCardVersion.V3_0, "BDAY", "2026-07-18T09:00:00,Z", "date-time")]
    [InlineData(VCardVersion.V3_0, "BDAY", "2026-07-18T09:00:00.5Z", "date-time")]
    [InlineData(VCardVersion.V4_0, "BDAY", "not-a-date", null)]
    [InlineData(VCardVersion.V4_0, "BDAY", "20230229", null)]
    [InlineData(VCardVersion.V4_0, "BDAY", "1985-13", null)]
    [InlineData(VCardVersion.V4_0, "BDAY", "--0230", null)]
    [InlineData(VCardVersion.V4_0, "BDAY", "---00", null)]
    [InlineData(VCardVersion.V4_0, "BDAY", "T240000", null)]
    [InlineData(VCardVersion.V4_0, "BDAY", "T--61", null)]
    [InlineData(VCardVersion.V4_0, "BDAY", "T102200+2400", null)]
    [InlineData(VCardVersion.V4_0, "BDAY", "20260718", "date")]
    [InlineData(VCardVersion.V4_0, "ANNIVERSARY", "not-a-date", null)]
    public void ValidationRejectsInvalidVersionSpecificDateValues(VCardVersion version,
        string propertyName, string value, string? valueType) {
        var document = new VCardDocument();
        ContentLineComponent card = document.Cards.Single();
        VCardDocument.SetVersion(card, version);
        card.AddProperty("FN", "Versioned contact");
        card.AddProperty("N", "Contact;Versioned;;;");
        ContentLineProperty property = card.AddProperty(propertyName, value);
        if (valueType != null) property.SetParameter("VALUE", valueType);

        Assert.Contains(document.Validate(), issue => issue.Code == "VCARD_DATE_VALUE_INVALID" &&
            issue.PropertyName == propertyName);
    }

    [Theory]
    [InlineData(VCardVersion.V3_0, "BDAY", "20260718", null)]
    [InlineData(VCardVersion.V3_0, "BDAY", "2026-07-18", "date")]
    [InlineData(VCardVersion.V3_0, "BDAY", "2026-07-18T09:00:00Z", "date-time")]
    [InlineData(VCardVersion.V3_0, "BDAY", "1953-10-15T23:10:00,5Z", "date-time")]
    [InlineData(VCardVersion.V3_0, "BDAY", "19531015T231000,125-0600", "date-time")]
    [InlineData(VCardVersion.V4_0, "BDAY", "--0415", null)]
    [InlineData(VCardVersion.V4_0, "BDAY", "19961022T140000", null)]
    [InlineData(VCardVersion.V4_0, "BDAY", "--1022T1400", null)]
    [InlineData(VCardVersion.V4_0, "BDAY", "---22T14", null)]
    [InlineData(VCardVersion.V4_0, "BDAY", "1985-04", null)]
    [InlineData(VCardVersion.V4_0, "BDAY", "1985", null)]
    [InlineData(VCardVersion.V4_0, "BDAY", "---12", null)]
    [InlineData(VCardVersion.V4_0, "BDAY", "T102200", null)]
    [InlineData(VCardVersion.V4_0, "BDAY", "T1022", null)]
    [InlineData(VCardVersion.V4_0, "BDAY", "T10", null)]
    [InlineData(VCardVersion.V4_0, "BDAY", "T-2200", null)]
    [InlineData(VCardVersion.V4_0, "BDAY", "T--00", null)]
    [InlineData(VCardVersion.V4_0, "BDAY", "T102200Z", null)]
    [InlineData(VCardVersion.V4_0, "BDAY", "T102200-0800", "date-and-or-time")]
    [InlineData(VCardVersion.V4_0, "BDAY", "circa 1800", "text")]
    [InlineData(VCardVersion.V4_0, "ANNIVERSARY", "19960415", null)]
    public void ValidationAcceptsVersionSpecificDateValues(VCardVersion version,
        string propertyName, string value, string? valueType) {
        var document = new VCardDocument();
        ContentLineComponent card = document.Cards.Single();
        VCardDocument.SetVersion(card, version);
        card.AddProperty("FN", "Versioned contact");
        card.AddProperty("N", "Contact;Versioned;;;");
        ContentLineProperty property = card.AddProperty(propertyName, value);
        if (valueType != null) property.SetParameter("VALUE", valueType);

        Assert.DoesNotContain(document.Validate(), issue => issue.Code == "VCARD_DATE_VALUE_INVALID" &&
            issue.PropertyName == propertyName);
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
    public void LegacyQuotedPrintableSoftBreakRoundTripsAsOneContentLine() {
        const string source = "BEGIN:VCARD\r\nVERSION:2.1\r\n" +
            "FN;CHARSET=UTF-8;ENCODING=QUOTED-PRINTABLE:Jos=C3=\r\n =A9\r\nEND:VCARD\r\n";

        VCardDocument parsed = VCardDocument.Parse(source);
        string serialized = parsed.Serialize();
        VCardDocument reparsed = VCardDocument.Parse(serialized);

        Assert.Equal("Jos=C3=A9", parsed.Cards.Single().GetFirstProperty("FN")!.Value);
        Assert.Equal("Jos=C3=A9", reparsed.Cards.Single().GetFirstProperty("FN")!.Value);
        Assert.DoesNotContain("=\r\n =A9", serialized, StringComparison.Ordinal);
    }

    [Fact]
    public void LegacyQuotedPrintableHeaderAcceptsEscapedQuoteBeforeComma() {
        const string source = "BEGIN:VCARD\r\nVERSION:2.1\r\n" +
            "FN;ENCODING=QUOTED-PRINTABLE;X-NAME=\"Doe\\\", John\":Alpha=\r\n" +
            "Beta\r\nEND:VCARD\r\n";

        ContentLineProperty formattedName = VCardDocument.Parse(source).Cards.Single()
            .GetFirstProperty("FN")!;

        Assert.Equal("AlphaBeta", formattedName.Value);
        Assert.Equal("Doe\", John", formattedName.GetParameter("X-NAME")!.Values.Single());
        ContentLineProperty reparsed = VCardDocument.Parse(VCardDocument.Parse(source).Serialize())
            .Cards.Single().GetFirstProperty("FN")!;
        Assert.Equal("AlphaBeta", reparsed.Value);
        Assert.Equal("Doe\", John", reparsed.GetParameter("X-NAME")!.Values.Single());
    }

    [Fact]
    public void LegacyQuotedPrintableFoldingDoesNotTurnEncodedEqualsIntoASoftBreak() {
        const string header = "FN;CHARSET=UTF-8;ENCODING=QUOTED-PRINTABLE:";
        int prefixLength = 75 - Encoding.UTF8.GetByteCount(header) - 1;
        string expected = new string('A', prefixLength) + "=C3=A9";
        var document = new VCardDocument();
        ContentLineComponent card = document.Cards.Single();
        VCardDocument.SetVersion(card, VCardVersion.V2_1);
        card.AddProperty("FN", expected)
            .SetParameter("CHARSET", "UTF-8")
            .SetParameter("ENCODING", "QUOTED-PRINTABLE");

        string serialized = document.Serialize();
        string actual = VCardDocument.Parse(serialized).Cards.Single().GetFirstProperty("FN")!.Value;

        Assert.DoesNotContain("=\r\n ", serialized, StringComparison.Ordinal);
        Assert.Contains("\r\n =C3", serialized, StringComparison.Ordinal);
        Assert.Equal(expected, actual);
    }

    [Fact]
    public void UnrelatedParameterTextDoesNotEnableQuotedPrintableSoftBreaks() {
        const string source = "BEGIN:VCARD\r\nVERSION:3.0\r\n" +
            "X-TEST;X-NOTE=\"ENCODING=QP\":value=\r\nFN:Next property\r\nEND:VCARD\r\n";

        ContentLineComponent card = VCardDocument.Parse(source).Cards.Single();

        Assert.Equal("value=", card.GetFirstProperty("X-TEST")!.Value);
        Assert.Equal("Next property", card.GetFirstProperty("FN")!.Value);
    }

    [Fact]
    public void ParameterEncodingIsVersionAwareAcrossLegacyAndV4Cards() {
        const string source = "BEGIN:VCARD\r\nVERSION:3.0\r\n" +
            "FN;X-LITERAL=alpha^nbeta:Legacy\r\nEND:VCARD\r\n" +
            "BEGIN:VCARD\r\nVERSION:4.0\r\n" +
            "FN;X-QUOTE=alpha^'beta:Modern\r\nEND:VCARD\r\n";

        VCardDocument parsed = VCardDocument.Parse(source);
        Assert.Equal("alpha^nbeta", parsed.Cards[0].GetFirstProperty("FN")!
            .Parameters.Single().Values.Single());
        Assert.Equal("alpha\"beta", parsed.Cards[1].GetFirstProperty("FN")!
            .Parameters.Single().Values.Single());

        string serialized = parsed.Serialize();
        Assert.Contains("X-LITERAL=alpha^nbeta", serialized, StringComparison.Ordinal);
        Assert.DoesNotContain("X-LITERAL=alpha^^nbeta", serialized, StringComparison.Ordinal);
        Assert.Contains("X-QUOTE=alpha^'beta", serialized, StringComparison.Ordinal);

        VCardDocument reparsed = VCardDocument.Parse(serialized);
        Assert.Equal("alpha^nbeta", reparsed.Cards[0].GetFirstProperty("FN")!
            .Parameters.Single().Values.Single());
        Assert.Equal("alpha\"beta", reparsed.Cards[1].GetFirstProperty("FN")!
            .Parameters.Single().Values.Single());
    }

    [Fact]
    public void LegacyParameterWriterEscapesQuotesAndRejectsLineBreaks() {
        var quoted = new VCardDocument();
        VCardDocument.SetVersion(quoted.Cards.Single(), VCardVersion.V3_0);
        quoted.Cards.Single().AddProperty("FN", "Legacy").SetParameter("X-NAME", "a\"b");
        var multiline = new VCardDocument();
        VCardDocument.SetVersion(multiline.Cards.Single(), VCardVersion.V2_1);
        multiline.Cards.Single().AddProperty("FN", "Legacy").SetParameter("X-NAME", "a\nb");

        string serialized = quoted.Serialize();
        Assert.Contains("X-NAME=\"a\\\"b\"", serialized, StringComparison.Ordinal);
        Assert.Equal("a\"b", VCardDocument.Parse(serialized).Cards.Single()
            .GetFirstProperty("FN")!.GetParameter("X-NAME")!.Values.Single());
        Assert.Throws<InvalidDataException>(() => multiline.Serialize());
    }

    [Fact]
    public void LegacyParameterWriterPreservesBackslashRunsBeforeQuotes() {
        string[] values = { "a\"b", "a\\\"b", "a\\\\\"b", "a\\\\\\\"b" };

        foreach (string value in values) {
            var document = new VCardDocument();
            VCardDocument.SetVersion(document.Cards.Single(), VCardVersion.V3_0);
            document.Cards.Single().AddProperty("FN", "Legacy").SetParameter("X-NAME", value);

            string serialized = document.Serialize();
            string reparsed = VCardDocument.Parse(serialized).Cards.Single()
                .GetFirstProperty("FN")!.GetParameter("X-NAME")!.Values.Single();

            Assert.Equal(value, reparsed);
        }
    }

    [Fact]
    public void SerializationRejectsLiteralLineBreaksInRawPropertyValues() {
        var document = new VCardDocument();
        document.Cards.Single().AddProperty("FN", "safe\nEND:VCARD\nBEGIN:VCARD");

        Assert.Throws<InvalidDataException>(() => document.Serialize());
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
