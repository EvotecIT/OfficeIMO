using OfficeIMO.Reader;
using System.Linq;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests;

[Collection("ReaderRegistryNonParallel")]
public sealed class ReaderContentLineTests {
    [Fact]
    public void CalendarAndVCardKindsAndCapabilitiesAreBuiltIn() {
        Assert.Equal(21, (int)ReaderInputKind.Calendar);
        Assert.Equal(22, (int)ReaderInputKind.VCard);
        Assert.Equal(ReaderInputKind.Calendar, OfficeDocumentReader.Default.DetectKind("meeting.ics"));
        Assert.Equal(ReaderInputKind.Calendar, OfficeDocumentReader.Default.DetectKind("meeting.vcs"));
        Assert.Equal(ReaderInputKind.VCard, OfficeDocumentReader.Default.DetectKind("contact.vcf"));
        Assert.Equal(ReaderInputKind.VCard, OfficeDocumentReader.Default.DetectKind("contact.vcard"));

        ReaderHandlerCapability calendar = Assert.Single(OfficeDocumentReader.Default.GetCapabilities(),
            capability => capability.Id == "officeimo.reader.calendar");
        ReaderHandlerCapability vcard = Assert.Single(OfficeDocumentReader.Default.GetCapabilities(),
            capability => capability.Id == "officeimo.reader.vcard");
        Assert.Equal(ReaderInputKind.Calendar, calendar.Kind);
        Assert.Equal(ReaderInputKind.VCard, vcard.Kind);
        Assert.Contains(".ics", calendar.Extensions);
        Assert.Contains(".vcs", calendar.Extensions);
        Assert.Contains(".vcf", vcard.Extensions);
        Assert.Contains(".vcard", vcard.Extensions);
    }

    [Fact]
    public void CalendarReadUsesNativeParserAndRetainsUnknownProperties() {
        byte[] bytes = Encoding.UTF8.GetBytes("BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//Reader//EN\r\n" +
            "BEGIN:VEVENT\r\nUID:reader-event\r\nSUMMARY:Reader meeting\r\nX-READER-UNKNOWN:kept\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n");

        ReaderDetectionResult detection = OfficeDocumentReader.Default.Detect(bytes, "renamed.bin");
        ReaderChunk[] chunks = OfficeDocumentReader.Default.Read(bytes, "meeting.ics").ToArray();
        OfficeDocumentReadResult result = OfficeDocumentReader.Default.ReadDocument(bytes, "meeting.ics");

        Assert.Equal(ReaderInputKind.Calendar, detection.Kind);
        Assert.Equal(ReaderInputKind.Calendar, result.Kind);
        Assert.NotEmpty(chunks);
        Assert.All(chunks, chunk => Assert.Equal(ReaderInputKind.Calendar, chunk.Kind));
        Assert.Contains(chunks, chunk => chunk.Text.Contains("SUMMARY:Reader meeting") &&
            chunk.Text.Contains("X-READER-UNKNOWN:kept"));
    }

    [Fact]
    public void VCardReadSupportsMultipleCardsAndContentDetection() {
        byte[] bytes = Encoding.UTF8.GetBytes("BEGIN:VCARD\r\nVERSION:4.0\r\nFN:First contact\r\nEND:VCARD\r\n" +
            "BEGIN:VCARD\r\nVERSION:3.0\r\nFN:Second contact\r\nEND:VCARD\r\n");

        ReaderDetectionResult detection = OfficeDocumentReader.Default.Detect(bytes, "renamed.bin");
        ReaderChunk[] chunks = OfficeDocumentReader.Default.Read(bytes, "contacts.vcf").ToArray();

        Assert.Equal(ReaderInputKind.VCard, detection.Kind);
        Assert.NotEmpty(chunks);
        Assert.All(chunks, chunk => Assert.Equal(ReaderInputKind.VCard, chunk.Kind));
        Assert.Contains(chunks, chunk => chunk.Text.Contains("FN:First contact") &&
            chunk.Text.Contains("FN:Second contact"));
    }

    [Fact]
    public void VCardReadRoundTripsLegacyEscapedQuoteParameters() {
        byte[] bytes = Encoding.UTF8.GetBytes("BEGIN:VCARD\r\nVERSION:3.0\r\n" +
            "FN;X-NAME=\"Doe\\\", John\":Legacy contact\r\nEND:VCARD\r\n");

        ReaderChunk[] chunks = OfficeDocumentReader.Default.Read(bytes, "legacy.vcf").ToArray();

        Assert.NotEmpty(chunks);
        Assert.Contains(chunks, chunk =>
            chunk.Text.Contains("X-NAME=\"Doe\\\", John\"", StringComparison.Ordinal) &&
            chunk.Text.Contains("Legacy contact", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("BEGIN:VCALENDARJUNK\r\nplain text\r\n", ReaderInputKind.Calendar)]
    [InlineData("BEGIN:VCARDINAL\r\nplain text\r\n", ReaderInputKind.VCard)]
    public void ContentDetectionRequiresAnExactContentLineRoot(
        string content, ReaderInputKind falsePositiveKind) {
        ReaderDetectionResult detection = OfficeDocumentReader.Default.Detect(
            Encoding.ASCII.GetBytes(content), "renamed.bin");

        Assert.NotEqual(falsePositiveKind, detection.Kind);
    }
}
