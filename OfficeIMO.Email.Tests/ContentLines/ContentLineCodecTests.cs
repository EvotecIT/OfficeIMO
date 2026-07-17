using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class ContentLineCodecTests {
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
    public void LegacyQuotedParameterEscapesDoNotSplitEmbeddedCommaOrQuote() {
        const string source = "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//Test//EN\r\n" +
            "BEGIN:VEVENT\r\nATTENDEE;CN=\"Doe, \\\"John\\\"\":mailto:john@example.com\r\n" +
            "END:VEVENT\r\nEND:VCALENDAR\r\n";

        ContentLineProperty attendee = IcsDocument.Parse(source).GetComponents("VEVENT")
            .Single().GetFirstProperty("ATTENDEE")!;

        Assert.Equal("Doe, \"John\"", attendee.GetParameter("CN")!.Values.Single());
        Assert.Equal("mailto:john@example.com", attendee.Value);
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
}
