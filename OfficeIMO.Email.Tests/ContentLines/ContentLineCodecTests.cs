using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class ContentLineCodecTests {
    [Fact]
    public void LeadingUtf8BomIsIgnored() {
        byte[] payload = new UTF8Encoding(encoderShouldEmitUTF8Identifier: true).GetPreamble()
            .Concat(Encoding.UTF8.GetBytes(
                "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//BOM//EN\r\nEND:VCALENDAR\r\n"))
            .ToArray();

        IcsDocument document = IcsDocument.Load(payload);

        Assert.Single(document.Calendars);
        Assert.Equal("-//BOM//EN", document.Calendars[0].GetFirstProperty("PRODID")?.Value);
    }

    [Fact]
    public void MalformedInputTokensAreReportedAsInvalidData() {
        Assert.Throws<InvalidDataException>(() => IcsDocument.Parse(
            "BEGIN:VCAL ENDAR\r\nVERSION:2.0\r\nEND:VCAL ENDAR\r\n"));
        Assert.Throws<InvalidDataException>(() => VCardDocument.Parse(
            "BEGIN:VCARD\r\nVERSION:4.0\r\nBAD NAME:value\r\nEND:VCARD\r\n"));
    }

    [Fact]
    public void ComponentDelimitersRejectParameters() {
        Assert.Throws<InvalidDataException>(() => IcsDocument.Parse(
            "BEGIN;X-PARAM=value:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//Test//EN\r\nEND:VCALENDAR\r\n"));
        Assert.Throws<InvalidDataException>(() => VCardDocument.Parse(
            "BEGIN:VCARD\r\nVERSION:4.0\r\nFN:Test\r\nEND;X-PARAM=value:VCARD\r\n"));
    }

    [Fact]
    public void PublicContentLineApisRejectNonAsciiTokens() {
        Assert.Throws<ArgumentException>(() => new ContentLineComponent("VÉVENT"));
        Assert.Throws<ArgumentException>(() => new ContentLineProperty("NÅME", "value"));
        Assert.Throws<ArgumentException>(() => new ContentLineParameter("TYPÉ", "work"));

        var document = new VCardDocument();
        ContentLineProperty property = document.Cards.Single().AddProperty("EMAIL", "a@example.test");
        property.Group = "grøup";
        Assert.Throws<ArgumentException>(() => document.Serialize());
    }

    [Fact]
    public void ManyShortPhysicalFoldsUnfoldWithinTheConfiguredLinearBound() {
        var source = new StringBuilder("BEGIN:VCARD\r\nVERSION:4.0\r\nFN:first");
        for (int index = 0; index < 100_000; index++) source.Append("\r\n x");
        source.Append("\r\nEND:VCARD\r\n");

        VCardDocument document = VCardDocument.Parse(source.ToString(),
            new ContentLineReaderOptions(maxInputBytes: 1024 * 1024,
                maxUnfoldedLineBytes: 256 * 1024));

        Assert.Equal(100_005, document.Cards[0].GetFirstProperty("FN")?.Value.Length);
    }

    [Fact]
    public void ManyQuotedHeaderFoldsWithColonsParseWithinTheConfiguredLinearBound() {
        var source = new StringBuilder("BEGIN:VCARD\r\nVERSION:4.0\r\nFN;X-NAME=\"start");
        for (int index = 0; index < 10_000; index++) source.Append("\r\n :part");
        source.Append("\r\n end\":display name\r\nEND:VCARD\r\n");

        ContentLineProperty formattedName = VCardDocument.Parse(source.ToString(),
            new ContentLineReaderOptions(maxInputBytes: 256 * 1024,
                maxUnfoldedLineBytes: 128 * 1024)).Cards.Single().GetFirstProperty("FN")!;

        Assert.Equal("display name", formattedName.Value);
        Assert.StartsWith("start:part", formattedName.GetParameter("X-NAME")!.Values.Single());
    }

    [Fact]
    public void ManyDeferredQuotedPrintableSoftBreaksCompactWithinTheConfiguredLinearBound() {
        var source = new StringBuilder(
            "BEGIN:VCARD\r\nVERSION:2.1\r\nFN;ENCODING=QUOTED-PRINTABLE:start");
        for (int index = 0; index < 25_000; index++) source.Append("=\r\n x");
        source.Append("display name\r\nEND:VCARD\r\n");

        ContentLineProperty formattedName = VCardDocument.Parse(source.ToString(),
            new ContentLineReaderOptions(maxInputBytes: 256 * 1024,
                maxUnfoldedLineBytes: 64 * 1024)).Cards.Single().GetFirstProperty("FN")!;

        Assert.Equal("start" + new string('x', 25_000) + "display name", formattedName.Value);
    }

    [Fact]
    public void QuotedPrintableUnfoldingPreservesEqualsMarkersInThePropertyHeader() {
        const string source = "BEGIN:VCARD\r\nVERSION:2.1\r\n" +
            "FN;ENCODING=QUOTED-PRINTABLE;X-NAME=\"a=\r\n b\":value\r\nEND:VCARD\r\n";

        ContentLineProperty formattedName = VCardDocument.Parse(source)
            .Cards.Single().GetFirstProperty("FN")!;

        Assert.Equal("a=b", formattedName.GetParameter("X-NAME")!.Values.Single());
        Assert.Equal("value", formattedName.Value);
    }

    [Theory]
    [InlineData("")]
    [InlineData(" ")]
    public void QuotedPrintableSoftBreakMarkerDoesNotCountAgainstTheUnfoldedLineLimit(
        string continuationPrefix) {
        const string prefix = "FN;ENCODING=QUOTED-PRINTABLE:";
        string value = new string('a', 64 - prefix.Length - 1);
        string source = "BEGIN:VCARD\r\nVERSION:2.1\r\n" + prefix + value +
            "=\r\n" + continuationPrefix + "b\r\nEND:VCARD\r\n";

        ContentLineProperty formattedName = VCardDocument.Parse(source,
            new ContentLineReaderOptions(maxUnfoldedLineBytes: 64))
            .Cards.Single().GetFirstProperty("FN")!;

        Assert.Equal(value + "b", formattedName.Value);
    }

    [Fact]
    public void EmptyPhysicalFoldsDoNotDuplicateDeferredSoftBreakOffsets() {
        const string source = "BEGIN:VCARD\r\nVERSION:2.1\r\n" +
            "FN;ENCODING=QUOTED-PRINTABLE:A=\r\n \r\n B=\r\n \r\n C\r\nEND:VCARD\r\n";

        ContentLineProperty formattedName = VCardDocument.Parse(source)
            .Cards.Single().GetFirstProperty("FN")!;

        Assert.Equal("ABC", formattedName.Value);
    }

    [Fact]
    public void Parameters_RoundTripRfc6868AndRepeatedValues() {
        const string source = "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//Test//EN\r\n" +
            "BEGIN:VEVENT\r\nATTENDEE;CN=Dee^'Arcy^^Team^nLine;MEMBER=\"mailto:a@example.com\",\"mailto:b@example.com\":mailto:c@example.com\r\n" +
            "END:VEVENT\r\nEND:VCALENDAR\r\n";

        IcsDocument document = IcsDocument.Parse(source);
        ContentLineProperty attendee = document.GetComponents("VEVENT").Single().GetFirstProperty("ATTENDEE")!;

        Assert.Equal("Dee\"Arcy^Team\nLine", attendee.GetParameter("CN")!.Values.Single());
        Assert.Equal(new[] { "mailto:a@example.com", "mailto:b@example.com" },
            attendee.GetParameter("MEMBER")!.Values);

        IcsDocument reparsed = IcsDocument.Parse(document.Serialize());
        ContentLineProperty reparsedAttendee = reparsed.GetComponents("VEVENT").Single()
            .GetFirstProperty("ATTENDEE")!;
        Assert.Equal(attendee.GetParameter("CN")!.Values, reparsedAttendee.GetParameter("CN")!.Values);
        Assert.Equal(attendee.GetParameter("MEMBER")!.Values, reparsedAttendee.GetParameter("MEMBER")!.Values);
    }

    [Theory]
    [InlineData("\0")]
    [InlineData("\b")]
    [InlineData("\u001F")]
    [InlineData("\u007F")]
    public void WritersRejectProhibitedAsciiControlsInMutableParameterValues(string control) {
        var calendar = new IcsDocument();
        calendar.Calendars.Single().AddProperty("X-CONTROL", "value")
            .SetParameter("X-PARAM", "left" + control + "right");
        var contact = new VCardDocument();
        contact.Cards.Single().GetFirstProperty("VERSION")!.Value = "3.0";
        contact.Cards.Single().AddProperty("FN", "Control test")
            .SetParameter("X-PARAM", "left" + control + "right");

        Assert.Throws<InvalidDataException>(() => calendar.Serialize());
        Assert.Throws<InvalidDataException>(() => contact.Serialize());
    }

    [Fact]
    public void WritersAllowHorizontalTabsInMutableParameterValues() {
        var calendar = new IcsDocument();
        calendar.Calendars.Single().AddProperty("X-TAB", "value")
            .SetParameter("X-PARAM", "left\tright");
        var contact = new VCardDocument();
        contact.Cards.Single().GetFirstProperty("VERSION")!.Value = "3.0";
        contact.Cards.Single().AddProperty("FN", "Tab test")
            .SetParameter("X-PARAM", "left\tright");

        Assert.Contains("X-PARAM=left\tright", calendar.Serialize(), StringComparison.Ordinal);
        Assert.Contains("X-PARAM=left\tright", contact.Serialize(), StringComparison.Ordinal);
    }

    [Fact]
    public void Writer_FoldsUnicodeAtUtf8OctetBoundary() {
        var document = new IcsDocument();
        ContentLineComponent calendar = document.Calendars.Single();
        ContentLineComponent appointment = calendar.AddComponent("VEVENT");
        appointment.AddProperty("UID", "folding@example.com");
        appointment.AddProperty("SUMMARY", string.Concat(Enumerable.Repeat("Meeting 😀 zażółć ", 12)));

        string serialized = document.Serialize();
        string[] physicalLines = serialized.Split(new[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);

        Assert.Contains(physicalLines, line => line.StartsWith(" ", StringComparison.Ordinal));
        Assert.All(physicalLines, line => Assert.True(Encoding.UTF8.GetByteCount(line) <= 75, line));
        Assert.Equal(appointment.GetFirstProperty("SUMMARY")!.Value,
            IcsDocument.Parse(serialized).GetComponents("VEVENT").Single().GetFirstProperty("SUMMARY")!.Value);
    }

    [Fact]
    public void Reader_RejectsConfiguredInputAndLineLimits() {
        const string source = "BEGIN:VCARD\r\nVERSION:4.0\r\nFN:Long display name\r\nEND:VCARD\r\n";

        Assert.Throws<InvalidDataException>(() => VCardDocument.Parse(source,
            new ContentLineReaderOptions(maxInputBytes: 8)));
        Assert.Throws<InvalidDataException>(() => VCardDocument.Parse(source,
            new ContentLineReaderOptions(maxUnfoldedLineBytes: 8)));
    }

    [Fact]
    public void Writer_RejectsConfiguredOutputLimit() {
        var document = new VCardDocument();
        document.Cards.Single().AddProperty("FN", "A contact name");

        Assert.Throws<InvalidDataException>(() => document.ToBytes(
            new ContentLineWriterOptions(maxOutputBytes: 16)));
    }

    [Fact]
    public void ReaderRejectsNestingLimitsBeyondTheTraversableModelDepth() {
        ContentLineReaderOptions supported = new ContentLineReaderOptions(maxNestingDepth: 256);

        Assert.Equal(256, supported.MaxNestingDepth);
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new ContentLineReaderOptions(maxNestingDepth: 257));
    }

    [Fact]
    public void LegacyQuotedParameterEscapesDoNotSplitEmbeddedCommaOrQuote() {
        const string source = "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//Test//EN\r\n" +
            "BEGIN:VEVENT\r\nATTENDEE;CN=\"Doe, \\\"John\\\"\";X-PATH=\"C:\\Temp\":mailto:john@example.com\r\n" +
            "END:VEVENT\r\nEND:VCALENDAR\r\n";

        ContentLineProperty attendee = IcsDocument.Parse(source).GetComponents("VEVENT")
            .Single().GetFirstProperty("ATTENDEE")!;

        Assert.Equal("Doe, \"John\"", attendee.GetParameter("CN")!.Values.Single());
        Assert.Equal("C:\\Temp", attendee.GetParameter("X-PATH")!.Values.Single());
        Assert.Equal("mailto:john@example.com", attendee.Value);

        ContentLineProperty reparsed = IcsDocument.Parse(IcsDocument.Parse(source).Serialize())
            .GetComponents("VEVENT").Single().GetFirstProperty("ATTENDEE")!;
        Assert.Equal("C:\\Temp", reparsed.GetParameter("X-PATH")!.Values.Single());
    }

    [Theory]
    [InlineData("CN=\"Doe\\\", John\"", "Doe\", John")]
    [InlineData("CN=\"Doe\\\"; John\"", "Doe\"; John")]
    [InlineData("CN=\"key\\\": value\"", "key\": value")]
    public void LegacyEscapedQuotesBeforeDelimitersRemainInsideTheParameter(
        string parameter, string expected) {
        string source = "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//Test//EN\r\n" +
            "BEGIN:VEVENT\r\nATTENDEE;" + parameter + ":mailto:a@example.com\r\n" +
            "END:VEVENT\r\nEND:VCALENDAR\r\n";

        ContentLineProperty attendee = IcsDocument.Parse(source).GetComponents("VEVENT")
            .Single().GetFirstProperty("ATTENDEE")!;

        Assert.Equal(expected, attendee.GetParameter("CN")!.Values.Single());
        Assert.Equal("mailto:a@example.com", attendee.Value);
    }

    [Fact]
    public void QuotedRfcParametersPreserveLiteralAndTrailingBackslashes() {
        const string source = "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//Test//EN\r\n" +
            "BEGIN:VEVENT\r\nATTENDEE;X-DOUBLE=\"C:\\\\Temp\";X-TRAIL=\"C:\\\":mailto:a@example.com\r\n" +
            "END:VEVENT\r\nEND:VCALENDAR\r\n";

        ContentLineProperty attendee = IcsDocument.Parse(source).GetComponents("VEVENT")
            .Single().GetFirstProperty("ATTENDEE")!;

        Assert.Equal("C:\\\\Temp", attendee.GetParameter("X-DOUBLE")!.Values.Single());
        Assert.Equal("C:\\", attendee.GetParameter("X-TRAIL")!.Values.Single());
        ContentLineProperty reparsed = IcsDocument.Parse(IcsDocument.Parse(source).Serialize())
            .GetComponents("VEVENT").Single().GetFirstProperty("ATTENDEE")!;
        Assert.Equal(attendee.GetParameter("X-DOUBLE")!.Values,
            reparsed.GetParameter("X-DOUBLE")!.Values);
        Assert.Equal(attendee.GetParameter("X-TRAIL")!.Values,
            reparsed.GetParameter("X-TRAIL")!.Values);
    }

    [Fact]
    public void WriterCountsContinuationSpaceInConfiguredEncoding() {
        var document = new VCardDocument();
        document.Cards.Single().AddProperty("FN", "abcdefghijklmnop");
        var encoding = new UnicodeEncoding(bigEndian: false, byteOrderMark: false);
        var options = new ContentLineWriterOptions(foldAtOctets: 12, encoding: encoding);

        string serialized = encoding.GetString(document.ToBytes(options));
        string[] lines = serialized.Split(new[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);

        Assert.Contains(lines, line => line.StartsWith(" ", StringComparison.Ordinal));
        Assert.All(lines, line => Assert.True(encoding.GetByteCount(line) <= 12, line));
    }

    [Fact]
    public void WriterRejectsAFoldLimitThatCannotFitAUnicodeContinuation() {
        var document = new VCardDocument();
        document.Cards.Single().AddProperty("FN", "😀");

        Assert.Throws<InvalidDataException>(() => document.Serialize(
            new ContentLineWriterOptions(foldAtOctets: 4)));
    }
}
