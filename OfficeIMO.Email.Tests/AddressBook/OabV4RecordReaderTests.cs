using OfficeIMO.Email;

namespace OfficeIMO.Email.AddressBook.Tests;

public sealed class OabV4RecordReaderTests {
    [Fact]
    public void SchemaDefinitionExposesSharedKnownProperty() {
        var definition = new OfflineAddressBookPropertyDefinition(
            ((uint)OabPropertyTags.DisplayName << 16) | (ushort)MapiPropertyType.Unicode, 0);

        Assert.Same(MapiKnownProperties.PidTag.DisplayName, definition.KnownProperty);
    }

    [Fact]
    public void RejectsNonCanonicalCompactIntegers() {
        var definition = new OfflineAddressBookPropertyDefinition(
            ((uint)OabPropertyTags.ObjectType << 16) | (ushort)MapiPropertyType.Integer32, 0);
        var envelope = new OabRecordEnvelope(7, new byte[] { 0x80, 0x81, 0x01 });

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            OabV4RecordReader.Parse(envelope, new[] { definition },
                OfflineAddressBookReaderOptions.Default, "entry"));

        Assert.Contains("Non-canonical compact OAB integer", exception.Message);
    }

    [Fact]
    public void RejectsPresentEmptyStrings() {
        var definition = new OfflineAddressBookPropertyDefinition(
            ((uint)OabPropertyTags.DisplayName << 16) | (ushort)MapiPropertyType.Unicode, 0);
        var envelope = new OabRecordEnvelope(6, new byte[] { 0x80, 0x00 });

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            OabV4RecordReader.Parse(envelope, new[] { definition },
                OfflineAddressBookReaderOptions.Default, "entry"));

        Assert.Contains("Empty OAB string value is marked present", exception.Message);
    }

    [Fact]
    public void RejectsAbsentPrimaryKeyProperties() {
        var definition = new OfflineAddressBookPropertyDefinition(
            ((uint)OabPropertyTags.EmailAddress << 16) | (ushort)MapiPropertyType.Unicode, 2);
        var envelope = new OabRecordEnvelope(5, new byte[] { 0x00 });

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            OabV4RecordReader.Parse(envelope, new[] { definition },
                OfflineAddressBookReaderOptions.Default, "entry"));

        Assert.Contains("Required OAB primary-key property", exception.Message);
    }

    [Fact]
    public void PreservesNonFatalPresenceAndTrailingDataDiagnostics() {
        var definition = new OfflineAddressBookPropertyDefinition(
            ((uint)OabPropertyTags.ObjectType << 16) | (ushort)MapiPropertyType.Integer32, 0);
        var envelope = new OabRecordEnvelope(7, new byte[] { 0x81, 0x06, 0x7F });

        OabParsedRecord parsed = OabV4RecordReader.Parse(envelope, new[] { definition },
            OfflineAddressBookReaderOptions.Default, "entry");

        Assert.Equal(2, parsed.Diagnostics.Count);
        Assert.Contains(parsed.Diagnostics, diagnostic => diagnostic.Code == "OAB_RECORD_UNUSED_PRESENCE_BITS");
        Assert.Contains(parsed.Diagnostics, diagnostic => diagnostic.Code == "OAB_RECORD_TRAILING_DATA");
    }
}
